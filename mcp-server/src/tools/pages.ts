import { graphClient } from "../auth.js";

export async function listPages(siteId: string) {
  const result = await graphClient
    .api(`/sites/${siteId}/pages`)
    .get();

  return result.value.map((page: any) => ({
    id: page.id,
    name: page.name,
    title: page.title,
    url: page.webUrl,
    publishingState: page.publishingState?.level,
  }));
}

export async function createPage(
  siteId: string,
  title: string,
  name: string,
  layoutType: string = "article"
) {
  const result = await graphClient
    .api(`/sites/${siteId}/pages`)
    .post({
      "@odata.type": "#microsoft.graph.sitePage",
      name: name.endsWith(".aspx") ? name : `${name}.aspx`,
      title,
      pageLayout: layoutType,
      showComments: false,
      showRecommendedPages: false,
      titleArea: {
        enableGradientEffect: false,
        layout: "plain",
        showAuthor: false,
        showPublishedDate: false,
        showTextBlockAboveTitle: false,
        textAlignment: "left",
      },
    });

  return {
    id: result.id,
    name: result.name,
    title: result.title,
    url: result.webUrl,
    publishingState: result.publishingState?.level,
  };
}

export async function createPageWithContent(
  siteId: string,
  title: string,
  name: string,
  sections: { layout: string; columns: { width: number; html: string }[] }[]
) {
  const horizontalSections = sections.map((section, i) => ({
    layout: section.layout,
    id: String(i + 1),
    emphasis: "none",
    columns: section.columns.map((col, j) => ({
      id: String(j + 1),
      width: col.width,
      webparts: [{
        id: crypto.randomUUID(),
        innerHtml: col.html,
      }],
    })),
  }));

  const result = await graphClient
    .api(`/sites/${siteId}/pages`)
    .post({
      "@odata.type": "#microsoft.graph.sitePage",
      name: name.endsWith(".aspx") ? name : `${name}.aspx`,
      title,
      pageLayout: "article",
      showComments: false,
      showRecommendedPages: false,
      titleArea: {
        enableGradientEffect: false,
        layout: "plain",
        showAuthor: false,
        showPublishedDate: false,
        showTextBlockAboveTitle: false,
        textAlignment: "left",
      },
      canvasLayout: { horizontalSections },
    });

  return {
    id: result.id,
    name: result.name,
    title: result.title,
    url: result.webUrl,
  };
}

export async function addQuickLinksWebPart(
  siteId: string,
  pageId: string,
  links: { title: string; url: string }[]
) {
  const items = links.map((link) => ({
    sourceItem: {
      guId: crypto.randomUUID(),
      url: link.url,
      itemType: 2,
      title: link.title,
      thumbnailType: 3,
    },
  }));

  const searchablePlainTexts = links.map((link, i) => ({
    key: `items[${i}].title`,
    value: link.title,
  }));

  const linkEntries = links.map((link, i) => ({
    key: `items[${i}].sourceItem.url`,
    value: link.url,
  }));

  const result = await graphClient
    .api(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage`)
    .patch({
      canvasLayout: {
        horizontalSections: [
          {
            layout: "oneColumn",
            id: "1",
            emphasis: "none",
            columns: [
              {
                id: "1",
                width: 12,
                webparts: [
                  {
                    id: crypto.randomUUID(),
                    webPartType: "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd",
                    data: {
                      dataVersion: "2.2",
                      title: "Quick links",
                      properties: {
                        items,
                        isMigrated: true,
                        layoutId: "CompactCard",
                        shouldShowThumbnail: true,
                        hideWebPartWhenEmpty: true,
                      },
                      serverProcessedContent: {
                        searchablePlainTexts,
                        links: linkEntries,
                      },
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    });

  return result;
}

export async function publishPage(siteId: string, pageId: string) {
  await graphClient
    .api(`/sites/${siteId}/pages/${pageId}/microsoft.graph.sitePage/publish`)
    .post({});

  return { success: true, pageId };
}

export async function deletePage(siteId: string, pageId: string) {
  await graphClient
    .api(`/sites/${siteId}/pages/${pageId}`)
    .delete();

  return { success: true, pageId };
}
