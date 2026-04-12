import { callSharePointRest } from "../auth.js";

export async function createSite(
  title: string,
  urlSlug: string,
  template: "communication" | "team" = "communication"
) {
  const domain = process.env.SHAREPOINT_DOMAIN!;
  const adminUrl = `https://${domain.replace(".sharepoint.com", "-admin.sharepoint.com")}`;
  const siteUrl = `https://${domain}/sites/${urlSlug}`;

  const webTemplate = template === "communication"
    ? "SITEPAGEPUBLISHING#0"
    : "STS#3";

  const result = await callSharePointRest(
    adminUrl,
    "/_api/SPSiteManager/create",
    "POST",
    {
      request: {
        Title: title,
        Url: siteUrl,
        WebTemplate: webTemplate,
      },
    }
  ) as any;

  return {
    siteId: result.d?.SiteId || result.SiteId,
    url: siteUrl,
    title: title,
  };
}
