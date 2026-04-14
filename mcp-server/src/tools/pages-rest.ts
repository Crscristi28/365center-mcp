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

// Known SharePoint web part IDs
const KNOWN_WEB_PARTS: Record<string, string> = {
  "daf0b71c-6de8-4ef7-b511-faae7c388708": "HighlightedContent",
  "c70391ea-0b10-4ee9-b2b4-006d3fcad0cd": "QuickLinks",
};

type HighlightedContentSummary = {
  instanceId: string;
  type: "HighlightedContent";
  zoneIndex: number;
  sectionIndex: number;
  title: string;
  filter: string;
  layout: string;
  maxItems: number;
};

type QuickLinksSummary = {
  instanceId: string;
  type: "QuickLinks";
  zoneIndex: number;
  sectionIndex: number;
  layout: string;
  links: Array<{ title: string; description: string; url: string }>;
};

type UnknownWebPartSummary = {
  instanceId: string;
  type: "Unknown";
  webPartId: string;
  zoneIndex: number;
  sectionIndex: number;
};

type WebPartSummary = HighlightedContentSummary | QuickLinksSummary | UnknownWebPartSummary;

type TextSection = {
  zoneIndex: number;
  sectionIndex: number;
  sectionFactor?: number;
  zoneEmphasis?: number;
  heading: string | null;
  content: string;
};

function decodeHtmlEntities(str: string): string {
  return str
    .replace(/&quot;/g, '"')
    .replace(/&#123;/g, "{")
    .replace(/&#125;/g, "}")
    .replace(/&#58;/g, ":")
    .replace(/&#91;/g, "[")
    .replace(/&#93;/g, "]")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCharCode(parseInt(n, 16)));
}

function stripHtmlTags(html: string): string {
  return html.replace(/<[^>]+>/g, "").trim();
}

export async function getPageCanvasSummary(siteUrl: string, pageItemId: number) {
  const listId = await getSitePagesListId(siteUrl);

  const result = await callSharePointRest(
    siteUrl,
    `/_api/web/lists(guid'${listId}')/items(${pageItemId})?$select=Id,Title,FileLeafRef,CanvasContent1`,
    "GET"
  ) as any;

  const rawCanvas: string = result.d.CanvasContent1 || "";

  // Extract all data-sp-controldata attribute values (encoded JSON)
  const controlMatches = Array.from(rawCanvas.matchAll(/data-sp-controldata="([^"]*)"/g));
  // Extract all data-sp-webpartdata attribute values (encoded JSON) in order
  const webpartMatches = Array.from(rawCanvas.matchAll(/data-sp-webpartdata="([^"]*)"/g));
  // Extract all RTE blocks (raw HTML inside data-sp-rte divs) in order
  const rteMatches = Array.from(rawCanvas.matchAll(/data-sp-rte="">([\s\S]*?)<\/div>/g));

  const webParts: WebPartSummary[] = [];
  const textSections: TextSection[] = [];
  let wpIdx = 0;
  let rteIdx = 0;

  for (const ctrlMatch of controlMatches) {
    const encodedControl = ctrlMatch[1];
    if (!encodedControl) continue;

    let ctrl: any;
    try {
      ctrl = JSON.parse(decodeHtmlEntities(encodedControl));
    } catch {
      continue;
    }

    const controlType = ctrl.controlType;
    const position = ctrl.position || {};
    const zoneIndex = position.zoneIndex ?? 0;
    const sectionIndex = position.sectionIndex ?? 1;
    const sectionFactor = position.sectionFactor;
    const emphasis = ctrl.emphasis || {};
    const zoneEmphasis = emphasis.zoneEmphasis;

    // Skip system controls: 0=pageSettingsSlice, 1=empty section, 14=background
    if (controlType === 0 || controlType === 1 || controlType === 14) {
      continue;
    }

    if (controlType === 4) {
      // RTE text block
      const rteMatch = rteMatches[rteIdx++];
      if (!rteMatch) continue;
      const rteHtml = rteMatch[1];

      const headingMatch = rteHtml.match(/<h\d[^>]*>([\s\S]*?)<\/h\d>/);
      const paragraphMatch = rteHtml.match(/<p[^>]*>([\s\S]*?)<\/p>/);

      const section: TextSection = {
        zoneIndex,
        sectionIndex,
        heading: headingMatch ? stripHtmlTags(headingMatch[1]) : null,
        content: paragraphMatch ? stripHtmlTags(paragraphMatch[1]) : "",
      };
      if (sectionFactor !== undefined) section.sectionFactor = sectionFactor;
      if (zoneEmphasis !== undefined) section.zoneEmphasis = zoneEmphasis;
      textSections.push(section);
    } else if (controlType === 3) {
      // Web part
      const wpMatch = webpartMatches[wpIdx++];
      if (!wpMatch) continue;
      const encodedWp = wpMatch[1];

      let wp: any;
      try {
        wp = JSON.parse(decodeHtmlEntities(encodedWp));
      } catch {
        continue;
      }

      const webPartId = wp.id || "";
      const type = KNOWN_WEB_PARTS[webPartId] || "Unknown";
      const instanceId = wp.instanceId || "";
      const serverContent = wp.serverProcessedContent || {};
      const searchablePlainTexts = serverContent.searchablePlainTexts || {};
      const properties = wp.properties || {};

      if (type === "HighlightedContent") {
        const filters = properties.query?.filters || [];
        webParts.push({
          instanceId,
          type: "HighlightedContent",
          zoneIndex,
          sectionIndex,
          title: searchablePlainTexts.title || "",
          filter: filters[0]?.value || "",
          layout: properties.layoutId || "",
          maxItems: properties.maxItemsPerPage ?? 8,
        });
      } else if (type === "QuickLinks") {
        const linksMap = serverContent.links || {};
        const links: Array<{ title: string; description: string; url: string }> = [];
        let i = 0;
        while (searchablePlainTexts[`items[${i}].title`] !== undefined) {
          links.push({
            title: searchablePlainTexts[`items[${i}].title`] || "",
            description: searchablePlainTexts[`items[${i}].description`] || "",
            url: linksMap[`items[${i}].sourceItem.url`] || "",
          });
          i++;
        }
        webParts.push({
          instanceId,
          type: "QuickLinks",
          zoneIndex,
          sectionIndex,
          layout: properties.layoutId || "",
          links,
        });
      } else {
        webParts.push({
          instanceId,
          type: "Unknown",
          webPartId,
          zoneIndex,
          sectionIndex,
        });
      }
    }
  }

  return {
    id: result.d.Id,
    title: result.d.Title,
    fileName: result.d.FileLeafRef,
    webParts,
    textSections,
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

// ============ PATCH WEB PART (string-replace, never JSON round-trip) ============

// HTML entity encoding for values going INTO data-sp-webpartdata attribute.
// SharePoint uses ASCII-special entities but leaves Unicode as literal UTF-8.
// Order matters: & must be first to avoid double-encoding.
function encodeForWebpartData(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\{/g, "&#123;")
    .replace(/\}/g, "&#125;")
    .replace(/:/g, "&#58;")
    .replace(/\[/g, "&#91;")
    .replace(/\]/g, "&#93;");
}

// HTML escape for plain text going INTO a html div (data-sp-htmlproperties block).
// Only escape characters that would break HTML parsing. Everything else stays literal.
function escapeForHtmlText(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Locate the data-sp-webpartdata="..." attribute value containing a given instanceId.
// Returns { attrStart, attrEnd } as indices into the canvas string, delimiting the
// value between the opening `data-sp-webpartdata="` and the closing `"`.
// Returns null if not found.
function locateWebpartDataAttr(canvas: string, instanceId: string): { attrStart: number; attrEnd: number } | null {
  const encodedInstanceIdPattern = `&quot;instanceId&quot;&#58;&quot;${instanceId}&quot;`;
  const instanceIdPos = canvas.indexOf(encodedInstanceIdPattern);
  if (instanceIdPos === -1) return null;

  const attrOpenMarker = 'data-sp-webpartdata="';
  const attrOpenPos = canvas.lastIndexOf(attrOpenMarker, instanceIdPos);
  if (attrOpenPos === -1) return null;
  const attrStart = attrOpenPos + attrOpenMarker.length;

  const attrEnd = canvas.indexOf('"', attrStart);
  if (attrEnd === -1) return null;
  if (instanceIdPos >= attrEnd) return null;

  return { attrStart, attrEnd };
}

// Locate the <div data-sp-prop-name="title" ...>TEXT</div> element for a specific
// web part. The htmlproperties div comes AFTER webpartdata attribute within the
// same <div data-sp-webpart=""> block. We scan forward from the end of webpartdata
// and stop before any next data-sp-webpart= (which would belong to a different web part).
function locateHtmlPropertiesTitle(
  canvas: string,
  webpartAttrEnd: number
): { titleStart: number; titleEnd: number } | null {
  const nextWebpartPos = canvas.indexOf('data-sp-webpart="', webpartAttrEnd);
  const searchLimit = nextWebpartPos === -1 ? canvas.length : nextWebpartPos;

  const openMarker = '<div data-sp-prop-name="title" data-sp-searchableplaintext="true">';
  const openPos = canvas.indexOf(openMarker, webpartAttrEnd);
  if (openPos === -1 || openPos >= searchLimit) return null;
  const titleStart = openPos + openMarker.length;

  const closeMarker = '</div>';
  const closePos = canvas.indexOf(closeMarker, titleStart);
  if (closePos === -1 || closePos >= searchLimit) return null;

  return { titleStart, titleEnd: closePos };
}

export type WebPartPatch = {
  instanceId: string;
  title?: string;
  filter?: string;
  maxItems?: number;
  layout?: string;
};

export type PatchWebPartResult = {
  success: boolean;
  pageItemId: number;
  patched: string[];
  notFound: string[];
  verifyFailed?: string[];
};

export async function patchPageCanvasWebpart(
  siteUrl: string,
  pageItemId: number,
  patches: WebPartPatch[],
  verify: boolean = true
): Promise<PatchWebPartResult> {
  // 1. Read current canvas
  const { canvasContent: originalCanvas } = await getPageCanvasContent(siteUrl, pageItemId);
  if (!originalCanvas) {
    throw new Error(`Page ${pageItemId} has empty canvasContent — nothing to patch`);
  }

  // 2. Use summary to get authoritative current values per instanceId
  // (string-replace needs old values to search for)
  const summary = await getPageCanvasSummary(siteUrl, pageItemId);
  const summaryByInstanceId = new Map<string, any>();
  for (const wp of summary.webParts) {
    summaryByInstanceId.set(wp.instanceId, wp);
  }

  let canvas = originalCanvas;
  const patched: string[] = [];
  const notFound: string[] = [];

  for (const patch of patches) {
    const currentWp = summaryByInstanceId.get(patch.instanceId);
    if (!currentWp) {
      notFound.push(patch.instanceId);
      continue;
    }

    const webpartAttr = locateWebpartDataAttr(canvas, patch.instanceId);
    if (!webpartAttr) {
      notFound.push(patch.instanceId);
      continue;
    }

    let webpartValue = canvas.slice(webpartAttr.attrStart, webpartAttr.attrEnd);
    const beforeAttr = canvas.slice(0, webpartAttr.attrStart);
    const afterAttr = canvas.slice(webpartAttr.attrEnd);

    if (patch.title !== undefined) {
      const oldTitle = currentWp.title || "";
      const encodedOld = encodeForWebpartData(oldTitle);
      const encodedNew = encodeForWebpartData(patch.title);
      // Match inside searchablePlainTexts only — don't touch the generic wp.title "Highlighted content"
      const oldPattern = `&quot;searchablePlainTexts&quot;&#58;&#123;&quot;title&quot;&#58;&quot;${encodedOld}&quot;`;
      const newPattern = `&quot;searchablePlainTexts&quot;&#58;&#123;&quot;title&quot;&#58;&quot;${encodedNew}&quot;`;
      webpartValue = webpartValue.replace(oldPattern, newPattern);
    }

    if (patch.filter !== undefined && currentWp.type === "HighlightedContent") {
      const oldFilter = currentWp.filter || "";
      const encodedOld = encodeForWebpartData(oldFilter);
      const encodedNew = encodeForWebpartData(patch.filter);
      const oldPattern = `&quot;filters&quot;&#58;[&#123;&quot;filterType&quot;&#58;1,&quot;value&quot;&#58;&quot;${encodedOld}&quot;`;
      const newPattern = `&quot;filters&quot;&#58;[&#123;&quot;filterType&quot;&#58;1,&quot;value&quot;&#58;&quot;${encodedNew}&quot;`;
      webpartValue = webpartValue.replace(oldPattern, newPattern);
    }

    if (patch.maxItems !== undefined && currentWp.type === "HighlightedContent") {
      const oldMax = currentWp.maxItems;
      const oldPattern = `&quot;maxItemsPerPage&quot;&#58;${oldMax},`;
      const newPattern = `&quot;maxItemsPerPage&quot;&#58;${patch.maxItems},`;
      webpartValue = webpartValue.replace(oldPattern, newPattern);
    }

    if (patch.layout !== undefined) {
      const oldLayout = currentWp.layout || "";
      const encodedOld = encodeForWebpartData(oldLayout);
      const encodedNew = encodeForWebpartData(patch.layout);
      const oldPattern = `&quot;layoutId&quot;&#58;&quot;${encodedOld}&quot;`;
      const newPattern = `&quot;layoutId&quot;&#58;&quot;${encodedNew}&quot;`;
      webpartValue = webpartValue.replace(oldPattern, newPattern);
    }

    canvas = beforeAttr + webpartValue + afterAttr;

    // Title also lives in data-sp-htmlproperties div (plain HTML, not entity-encoded).
    // Re-locate after the earlier string replace to get up-to-date indices.
    if (patch.title !== undefined) {
      const relocated = locateWebpartDataAttr(canvas, patch.instanceId);
      if (relocated) {
        const titleLoc = locateHtmlPropertiesTitle(canvas, relocated.attrEnd);
        if (titleLoc) {
          const escapedNewTitle = escapeForHtmlText(patch.title);
          canvas =
            canvas.slice(0, titleLoc.titleStart) +
            escapedNewTitle +
            canvas.slice(titleLoc.titleEnd);
        }
      }
    }

    patched.push(patch.instanceId);
  }

  // 3. Write the modified canvas back
  await setPageCanvasContent(siteUrl, pageItemId, canvas);

  // 4. Optional verification: read back and confirm new values are present
  let verifyFailed: string[] | undefined;
  if (verify && patched.length > 0) {
    const verifySummary = await getPageCanvasSummary(siteUrl, pageItemId);
    const verifyByInstanceId = new Map<string, any>();
    for (const wp of verifySummary.webParts) {
      verifyByInstanceId.set(wp.instanceId, wp);
    }

    const failed: string[] = [];
    for (const patch of patches) {
      if (!patched.includes(patch.instanceId)) continue;
      const wp = verifyByInstanceId.get(patch.instanceId);
      if (!wp) {
        failed.push(patch.instanceId);
        continue;
      }
      if (patch.title !== undefined && wp.title !== patch.title) {
        failed.push(patch.instanceId);
        continue;
      }
      if (patch.filter !== undefined && wp.filter !== patch.filter) {
        failed.push(patch.instanceId);
        continue;
      }
      if (patch.maxItems !== undefined && wp.maxItems !== patch.maxItems) {
        failed.push(patch.instanceId);
        continue;
      }
      if (patch.layout !== undefined && wp.layout !== patch.layout) {
        failed.push(patch.instanceId);
        continue;
      }
    }
    if (failed.length > 0) {
      verifyFailed = failed;
    }
  }

  return {
    success: notFound.length === 0 && (!verifyFailed || verifyFailed.length === 0),
    pageItemId,
    patched,
    notFound,
    ...(verifyFailed ? { verifyFailed } : {}),
  };
}

export async function fetchSitePageItemIdMap(siteUrl: string): Promise<Map<string, number>> {
  const listId = await getSitePagesListId(siteUrl);

  const allPages: any[] = [];
  let nextPath: string | null = `/_api/web/lists(guid'${listId}')/items?$select=Id,FileLeafRef&$top=100`;

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

  return new Map(allPages.map((page: any) => [page.FileLeafRef, page.Id]));
}
