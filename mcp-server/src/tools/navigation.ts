import { callSharePointRest } from "../auth.js";

export async function getNavigation(siteUrl: string) {
  const result = await callSharePointRest(siteUrl, "/_api/web/navigation/topnavigationbar") as any;

  return result.d.results.map((item: any) => ({
    id: item.Id,
    title: item.Title,
    url: item.Url,
    isExternal: item.IsExternal,
  }));
}

export async function addNavigationLink(siteUrl: string, title: string, url: string) {
  const result = await callSharePointRest(
    siteUrl,
    "/_api/web/navigation/topnavigationbar",
    "POST",
    {
      "__metadata": { "type": "SP.NavigationNode" },
      Title: title,
      Url: url,
      IsExternal: false,
    }
  ) as any;

  return {
    id: result.d.Id,
    title: result.d.Title,
    url: result.d.Url,
  };
}

export async function deleteNavigationLink(siteUrl: string, linkId: number) {
  await callSharePointRest(
    siteUrl,
    `/_api/web/navigation/topnavigationbar(${linkId})`,
    "DELETE"
  );

  return { success: true, linkId };
}
