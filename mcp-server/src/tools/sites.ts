import { graphClient } from "../auth.js";

export async function listSites() {
  const result = await graphClient
    .api("/sites?search=*")
    .header("ConsistencyLevel", "eventual")
    .top(50)
    .get();

  return result.value.map((site: any) => ({
    id: site.id,
    name: site.displayName,
    url: site.webUrl,
    description: site.description,
  }));
}

export async function getSite(siteUrl: string) {
  // Parse URL like "contoso.sharepoint.com/sites/MySite"
  const url = new URL(siteUrl.startsWith("http") ? siteUrl : `https://${siteUrl}`);
  const hostname = url.hostname;
  const serverRelativePath = url.pathname;

  const result = await graphClient
    .api(`/sites/${hostname}:${serverRelativePath}`)
    .get();

  return {
    id: result.id,
    name: result.displayName,
    url: result.webUrl,
    description: result.description,
  };
}

export async function getSiteById(siteId: string) {
  const result = await graphClient
    .api(`/sites/${siteId}`)
    .get();

  return {
    id: result.id,
    name: result.displayName,
    url: result.webUrl,
    description: result.description,
  };
}
