import { callSharePointRest } from "../auth.js";

// Helper: find Site Pages list ID for a site
async function getSitePagesListId(siteUrl: string): Promise<string> {
  const lists = await callSharePointRest(
    siteUrl,
    "/_api/web/lists?$select=Id&$filter=BaseTemplate eq 119",
    "GET"
  ) as any;

  if (!lists.d.results.length) {
    throw new Error("Site Pages list not found on this site");
  }
  return lists.d.results[0].Id;
}

export async function getPageCanvasContent(siteUrl: string, pageItemId: number) {
  const listId = await getSitePagesListId(siteUrl);

  const result = await callSharePointRest(
    siteUrl,
    `/_api/web/lists(guid'${listId}')/items(${pageItemId})?$select=Id,Title,FileLeafRef,CanvasContent1`,
    "GET"
  ) as any;

  return {
    id: result.d.Id,
    title: result.d.Title,
    fileName: result.d.FileLeafRef,
    canvasContent: result.d.CanvasContent1,
  };
}

export async function setPageCanvasContent(siteUrl: string, pageItemId: number, canvasContent: string) {
  const listId = await getSitePagesListId(siteUrl);

  // Update CanvasContent1 field using MERGE
  await callSharePointRest(
    siteUrl,
    `/_api/web/lists(guid'${listId}')/items(${pageItemId})`,
    "MERGE",
    {
      "__metadata": { "type": "SP.Data.SitePagesItem" },
      "CanvasContent1": canvasContent,
    }
  );

  return { success: true, pageItemId };
}

export async function copyPage(siteUrl: string, sourceFileName: string, targetFileName: string) {
  const sourcePath = `/sites/${new URL(siteUrl).pathname.split("/sites/")[1]}/SitePages/${sourceFileName}`;
  const targetPath = `/sites/${new URL(siteUrl).pathname.split("/sites/")[1]}/SitePages/${targetFileName}`;

  await callSharePointRest(
    siteUrl,
    `/_api/web/GetFileByServerRelativePath(decodedurl='${sourcePath}')/copyTo(strNewUrl='${targetPath}',bOverWrite=false)`,
    "POST"
  );

  return {
    success: true,
    source: sourceFileName,
    target: targetFileName,
  };
}

export async function listSitePages(siteUrl: string) {
  const listId = await getSitePagesListId(siteUrl);

  const allPages: any[] = [];
  let nextPath: string | null = `/_api/web/lists(guid'${listId}')/items?$select=Id,Title,FileLeafRef&$top=100`;

  while (nextPath) {
    const result: any = await callSharePointRest(siteUrl, nextPath, "GET");
    allPages.push(...result.d.results);

    if (result.d.__next) {
      const nextUrl = new URL(result.d.__next);
      nextPath = nextUrl.pathname + nextUrl.search;
    } else {
      nextPath = null;
    }
  }

  return allPages.map((page: any) => ({
    itemId: page.Id,
    title: page.Title,
    fileName: page.FileLeafRef,
  }));
}
