import { graphClient } from "../auth.js";
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

export async function listDocuments(siteId: string, driveId: string, folderId: string = "root") {
  const result = await graphClient
    .api(`/sites/${siteId}/drives/${driveId}/items/${folderId}/children?$expand=listItem`)
    .get();

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
  const result = await graphClient
    .api(`/sites/${siteId}/drive/root/search(q='${query}')`)
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
