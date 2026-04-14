import { graphClient } from "../auth.js";
import { setDocumentMetadata } from "./metadata.js";
import fs from "fs";
import path from "path";
import { Readable } from "stream";

export async function listDocumentLibraries(siteId: string) {
  const result = await graphClient
    .api(`/sites/${siteId}/drives`)
    .get();

  return result.value.map((drive: any) => ({
    id: drive.id,
    name: drive.name,
    url: drive.webUrl,
    itemCount: drive.quota?.used,
  }));
}

export async function listDocuments(
  siteId: string,
  driveId: string,
  folderId: string = "root",
  fields: "all" | "minimal" = "all"
) {
  const result = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${folderId}/children?$expand=listItem`)
    .get();

  if (fields === "minimal") {
    return result.value.map((item: any) => ({
      id: item.id,
      name: item.name,
      isFolder: !!item.folder,
      size: item.size,
    }));
  }

  return result.value.map((item: any) => ({
    id: item.id,
    listItemId: item.listItem?.id,
    name: item.name,
    url: item.webUrl,
    size: item.size,
    mimeType: item.file?.mimeType,
    isFolder: !!item.folder,
    lastModified: item.lastModifiedDateTime,
    createdBy: item.createdBy?.user?.displayName,
  }));
}

export async function uploadDocument(
  siteId: string,
  driveId: string,
  fileName: string,
  filePath: string,
  folderId: string = "root"
) {
  const fileSize = fs.statSync(filePath).size;
  const FOUR_MB = 4 * 1024 * 1024;

  if (fileSize <= FOUR_MB) {
    const fileContent = fs.readFileSync(filePath);
    const result = await graphClient
      .api(`/sites/${siteId}/drives/${driveId}/items/${folderId}:/${fileName}:/content`)
      .putStream(fileContent);

    return {
      id: result.id,
      name: result.name,
      url: result.webUrl,
      size: result.size,
    };
  }

  const session = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${folderId}:/${fileName}:/createUploadSession`)
    .post({
      item: { "@microsoft.graph.conflictBehavior": "rename" },
    });

  const uploadUrl = session.uploadUrl;
  const CHUNK_SIZE = 10 * 1024 * 1024;
  const fd = fs.openSync(filePath, "r");
  let offset = 0;
  let finalResult: any;

  try {
    while (offset < fileSize) {
      const length = Math.min(CHUNK_SIZE, fileSize - offset);
      const buffer = Buffer.alloc(length);
      fs.readSync(fd, buffer, 0, length, offset);

      const end = offset + length - 1;
      const response = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Length": length.toString(),
          "Content-Range": `bytes ${offset}-${end}/${fileSize}`,
        },
        body: buffer,
      });

      if (response.status === 200 || response.status === 201) {
        finalResult = await response.json();
      } else if (response.status !== 202) {
        const errorText = await response.text();
        throw new Error(`Upload chunk failed: ${response.status} ${errorText}`);
      }

      offset += length;
    }
  } finally {
    fs.closeSync(fd);
  }

  return {
    id: finalResult.id,
    name: finalResult.name,
    url: finalResult.webUrl,
    size: finalResult.size,
  };
}

export async function downloadDocument(
  siteId: string,
  driveId: string,
  itemId: string,
  localPath: string
) {
  const metadata = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}`)
    .select("name,size,file")
    .get();

  let targetPath = localPath;
  try {
    const stat = fs.statSync(localPath);
    if (stat.isDirectory()) {
      targetPath = path.join(localPath, metadata.name);
    }
  } catch {
    // Path does not exist — treat as full file path, ensure parent dir exists
    const parent = path.dirname(localPath);
    if (parent && !fs.existsSync(parent)) {
      fs.mkdirSync(parent, { recursive: true });
    }
  }

  const webStream: ReadableStream = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/content`)
    .getStream();

  const nodeStream = Readable.fromWeb(webStream as any);

  await new Promise<void>((resolve, reject) => {
    const writeStream = fs.createWriteStream(targetPath);
    nodeStream.pipe(writeStream);
    writeStream.on("finish", () => resolve());
    writeStream.on("error", reject);
    nodeStream.on("error", reject);
  });

  return {
    id: itemId,
    name: metadata.name,
    size: metadata.size,
    mimeType: metadata.file?.mimeType,
    savedTo: path.resolve(targetPath),
  };
}

export async function searchDocuments(siteId: string, query: string) {
  const safeQuery = query.replace(/'/g, "''");
  const result = await graphClient
    .api(`/sites/${siteId}/drive/root/search(q='${safeQuery}')`)
    .get();

  return result.value.map((item: any) => ({
    id: item.id,
    name: item.name,
    url: item.webUrl,
    size: item.size,
    lastModified: item.lastModifiedDateTime,
  }));
}

export async function deleteDocument(siteId: string, driveId: string, itemId: string) {
  await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}`)
    .delete();

  return { success: true, itemId };
}

export async function createFolder(siteId: string, driveId: string, folderName: string, parentFolderId: string = "root") {
  const result = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${parentFolderId}/children`)
    .post({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    });

  return {
    id: result.id,
    name: result.name,
    url: result.webUrl,
  };
}

export async function getDocumentVersions(siteId: string, driveId: string, itemId: string) {
  const result = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${itemId}/versions`)
    .get();

  return result.value.map((v: any) => ({
    id: v.id,
    lastModified: v.lastModifiedDateTime,
    modifiedBy: v.lastModifiedBy?.user?.displayName,
    size: v.size,
  }));
}

export async function uploadDocuments(
  siteId: string,
  driveId: string,
  listId: string | undefined,
  files: Array<{ fileName: string; filePath: string; fields?: string }>,
  folderId: string = "root"
) {
  const results: Array<{
    fileName: string;
    status: "ok" | "error";
    id?: string;
    error?: string;
    metadataStatus?: "ok" | "error" | "skipped";
    metadataError?: string;
  }> = [];

  for (const file of files) {
    const entry: typeof results[0] = { fileName: file.fileName, status: "ok" };

    try {
      const uploaded = await uploadDocument(siteId, driveId, file.fileName, file.filePath, folderId);
      entry.id = uploaded.id;

      if (file.fields && listId) {
        try {
          const fields = JSON.parse(file.fields);
          await setDocumentMetadata(siteId, listId!, uploaded.id, fields, driveId);
          entry.metadataStatus = "ok";
        } catch (metaErr: any) {
          entry.metadataStatus = "error";
          entry.metadataError = metaErr.message || String(metaErr);
        }
      } else {
        entry.metadataStatus = "skipped";
      }
    } catch (err: any) {
      entry.status = "error";
      entry.error = err.message || String(err);
      entry.metadataStatus = "skipped";
    }

    results.push(entry);
    await new Promise((r) => setTimeout(r, 500));
  }

  return results;
}
