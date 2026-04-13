#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { readFileSync } from "fs";
import { fileURLToPath } from "url";
import path from "path";
import { listSites, getSite, getSiteById } from "./tools/sites.js";
import { listDocumentLibraries, listDocuments, uploadDocument, uploadDocuments, downloadDocument, searchDocuments, deleteDocument, createFolder, getDocumentVersions } from "./tools/documents.js";
import { listColumns, createChoiceColumn, createTextColumn, setDocumentMetadata, getDocumentMetadata } from "./tools/metadata.js";
import { listPages, createPage, createPageWithContent, addQuickLinksWebPart, publishPage, deletePage } from "./tools/pages.js";
import { getNavigation, addNavigationLink, deleteNavigationLink } from "./tools/navigation.js";
import { getPageCanvasContent, setPageCanvasContent, copyPage, listSitePages } from "./tools/pages-rest.js";
import { createSite } from "./tools/sites-rest.js";
import { getSitePermissions, getGroupMembers, addUserToGroup, removeUserFromGroup } from "./tools/permissions.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const pkg = JSON.parse(readFileSync(path.join(__dirname, "..", "package.json"), "utf-8"));

const server = new McpServer({
  name: "365center-mcp",
  version: pkg.version,
});

// ============ SITES ============

server.tool(
  "list_sites",
  "List all SharePoint sites in the tenant. Returns siteId for each site — use siteId in all other tools. Site ID format: hostname,siteCollectionId,siteId (e.g. contoso.sharepoint.com,guid,guid).",
  {},
  async () => {
    const sites = await listSites();
    return { content: [{ type: "text", text: JSON.stringify(sites, null, 2) }] };
  }
);

server.tool(
  "get_site",
  "Get a SharePoint site by URL. URL format: contoso.sharepoint.com/sites/SiteName (without https://). Returns siteId needed for all other tools.",
  { siteUrl: z.string().describe("SharePoint site URL") },
  async ({ siteUrl }) => {
    const site = await getSite(siteUrl);
    return { content: [{ type: "text", text: JSON.stringify(site, null, 2) }] };
  }
);

server.tool(
  "get_site_by_id",
  "Get a SharePoint site by its ID. Site ID format: hostname,siteCollectionId,siteId (e.g. contoso.sharepoint.com,guid,guid).",
  { siteId: z.string().describe("SharePoint site ID") },
  async ({ siteId }) => {
    const site = await getSiteById(siteId);
    return { content: [{ type: "text", text: JSON.stringify(site, null, 2) }] };
  }
);

server.tool(
  "create_site",
  "Create a new SharePoint site. Template: 'communication' (default) for publishing/documentation sites, 'team' for collaboration. Returns siteId for use in all other tools. Site creation takes a few seconds. Uses delegated auth (device code flow on first use).",
  {
    title: z.string().describe("Display name for the new site"),
    urlSlug: z.string().describe("URL slug — the site will be at https://{domain}/sites/{urlSlug}"),
    template: z.enum(["communication", "team"]).optional().describe("Site template: 'communication' (default) or 'team'"),
  },
  async ({ title, urlSlug, template }) => {
    const result = await createSite(title, urlSlug, template || "communication");
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ DOCUMENTS ============

server.tool(
  "list_document_libraries",
  "List all document libraries (drives) in a SharePoint site. Returns driveId and listId for each library. Use driveId for file operations (upload, download, delete, list, versions). Use listId for metadata operations (list_columns, get/set_document_metadata, create columns). Call this first when working with documents — you need both IDs.",
  { siteId: z.string().describe("SharePoint site ID") },
  async ({ siteId }) => {
    const libraries = await listDocumentLibraries(siteId);
    return { content: [{ type: "text", text: JSON.stringify(libraries, null, 2) }] };
  }
);

server.tool(
  "list_documents",
  "List documents in a document library folder. Returns both driveItemId and listItemId for each document. Use driveItemId (+ driveId) for file operations (download, delete, versions). Use listItemId (or driveItemId + driveId) for metadata operations (get/set_document_metadata). The listId parameter for metadata tools can be the list display name (e.g. 'Dokumenty') or the list GUID.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    folderId: z.string().optional().describe("Folder ID (default: root)"),
  },
  async ({ siteId, driveId, folderId }) => {
    const docs = await listDocuments(siteId, driveId, folderId || "root");
    return { content: [{ type: "text", text: JSON.stringify(docs, null, 2) }] };
  }
);

server.tool(
  "upload_document",
  "Upload a single local file to a SharePoint document library. Root folder by default, or specific folder via folderId. After upload, use set_document_metadata to tag it. Supports up to 250 GB (auto session upload for >4 MB). For multiple files, prefer upload_documents — up to 30 files per call with inline metadata and built-in rate limit protection. NOTE: If using Highlighted Content web parts, filename prefix determines filtering. The 'Title includes the words' filter is SUBSTRING-based — e.g. 'Report' will also match 'ReportQ1' or 'ReportAnnual'. Use non-overlapping prefixes.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    fileName: z.string().describe("Name for the file in SharePoint"),
    filePath: z.string().describe("Local file path to upload"),
    folderId: z.string().optional().describe("Target folder ID (default: root)"),
  },
  async ({ siteId, driveId, fileName, filePath, folderId }) => {
    const result = await uploadDocument(siteId, driveId, fileName, filePath, folderId || "root");
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "download_document",
  "Download a document from a SharePoint document library to a local path. The localPath can be a full file path or a directory — if a directory, the original SharePoint filename is kept. Parent directories are created if missing.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    itemId: z.string().describe("Drive item ID (from list_documents or search_documents)"),
    localPath: z.string().describe("Local destination — either a full file path or a directory. If a directory, the original filename is preserved."),
  },
  async ({ siteId, driveId, itemId, localPath }) => {
    const result = await downloadDocument(siteId, driveId, itemId, localPath);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "upload_documents",
  "Batch upload up to 30 files to SharePoint with optional inline metadata per file. Files >4 MB use resumable upload. Built-in 500ms delay between files prevents HTTP 429 from Microsoft. Returns per-file status for upload and metadata independently. When setting metadata, listId is required. Choice values must match predefined choices exactly (case-sensitive). For large batches, split into multiple calls of max 30 files each. Same substring warning as upload_document — if using Highlighted Content web parts, filename prefixes determine which section shows which documents.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    listId: z.string().optional().describe("List ID — required only if setting metadata via fields"),
    folderId: z.string().optional().describe("Target folder ID (default: root)"),
    files: z.array(z.object({
      fileName: z.string().describe("Name for the file in SharePoint"),
      filePath: z.string().describe("Local file path to upload"),
      fields: z.string().optional().describe("JSON string of metadata fields to set after upload, e.g. '{\"Oblast\":\"Linka1\",\"WS\":\"WS9\"}'"),
    })).describe("Array of files to upload (max 30)"),
  },
  async ({ siteId, driveId, listId, folderId, files }) => {
    if (files.length > 30) {
      return { content: [{ type: "text", text: JSON.stringify({ error: "Maximum 30 files per call. Split into multiple calls." }) }] };
    }
    const results = await uploadDocuments(siteId, driveId, listId, files, folderId || "root");
    return { content: [{ type: "text", text: JSON.stringify(results, null, 2) }] };
  }
);

server.tool(
  "search_documents",
  "Search for documents across all libraries in a SharePoint site by keyword. Search indexes may take a few minutes to reflect newly uploaded documents.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    query: z.string().describe("Search query"),
  },
  async ({ siteId, query }) => {
    const results = await searchDocuments(siteId, query);
    return { content: [{ type: "text", text: JSON.stringify(results, null, 2) }] };
  }
);

server.tool(
  "delete_document",
  "PERMANENTLY delete a document from a SharePoint document library. This action is irreversible — the document goes to the site recycle bin but should be considered deleted. Always confirm with the user before deleting.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    itemId: z.string().describe("Drive item ID of the document to delete"),
  },
  async ({ siteId, driveId, itemId }) => {
    const result = await deleteDocument(siteId, driveId, itemId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "create_folder",
  "Create a new folder in a SharePoint document library. Folder is created in the root by default, or inside another folder if parentFolderId is provided.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    folderName: z.string().describe("Name of the new folder"),
    parentFolderId: z.string().optional().describe("Parent folder ID (default: root)"),
  },
  async ({ siteId, driveId, folderName, parentFolderId }) => {
    const result = await createFolder(siteId, driveId, folderName, parentFolderId || "root");
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "get_document_versions",
  "Get version history of a document (audit trail). Shows who modified the document, when, and version numbers. Useful for compliance and tracking changes.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    driveId: z.string().describe("Document library (drive) ID"),
    itemId: z.string().describe("Drive item ID"),
  },
  async ({ siteId, driveId, itemId }) => {
    const result = await getDocumentVersions(siteId, driveId, itemId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ METADATA ============

server.tool(
  "list_columns",
  "List custom metadata columns in a SharePoint list/library. listId can be display name (e.g. 'Documents') or GUID. Returns internal name, display name, type, and choices. Call this before set_document_metadata — internal column names (used in API) often differ from display names (shown in UI). Using wrong name fails silently.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    listId: z.string().describe("List or document library list ID"),
  },
  async ({ siteId, listId }) => {
    const columns = await listColumns(siteId, listId);
    return { content: [{ type: "text", text: JSON.stringify(columns, null, 2) }] };
  }
);

server.tool(
  "create_choice_column",
  "Create a choice/dropdown metadata column in a SharePoint list or document library. Use allowMultiple:true for multi-select checkboxes (e.g. document belongs to multiple areas). Column 'name' is the internal API name (no spaces/special chars), 'displayName' is what users see in SharePoint UI. The column must be created BEFORE setting metadata values on documents.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    listId: z.string().describe("List or document library list ID"),
    name: z.string().describe("Internal column name"),
    displayName: z.string().describe("Display name shown in UI"),
    choices: z.array(z.string()).describe("List of choices"),
    allowMultiple: z.boolean().optional().describe("Allow multiple selections (default: false)"),
  },
  async ({ siteId, listId, name, displayName, choices, allowMultiple }) => {
    const result = await createChoiceColumn(siteId, listId, name, displayName, choices, allowMultiple || false);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "create_text_column",
  "Create a single-line text metadata column in a SharePoint list or document library. Column 'name' is the internal API name (no spaces/special chars), 'displayName' is what users see in SharePoint UI. The column must be created BEFORE setting metadata values on documents.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    listId: z.string().describe("List or document library list ID"),
    name: z.string().describe("Internal column name"),
    displayName: z.string().describe("Display name shown in UI"),
  },
  async ({ siteId, listId, name, displayName }) => {
    const result = await createTextColumn(siteId, listId, name, displayName);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "set_document_metadata",
  "Set metadata fields on a document. 'fields' is a JSON string of key-value pairs using column internal names. Choice values must match predefined choices exactly (case-sensitive). Multi-select: use array of strings. Columns must exist first — check with list_columns. Accepts drive item ID (requires driveId) or numeric list item ID. Only change fields the user asked for — leave others untouched. TIP: Changing a Status field (e.g. to 'Pending Approval') can trigger Power Automate flows with 'When item modified' trigger — no HTTP webhooks or extra tools needed.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    listId: z.string().describe("List or document library list ID"),
    itemId: z.string().describe("Document ID — either numeric list item ID or drive item ID"),
    fields: z.string().describe("JSON string of key-value pairs, e.g. {\"Oblast\":\"Linka 1\",\"Status\":\"Platný\"}"),
    driveId: z.string().optional().describe("Drive ID — required when itemId is a drive item ID (non-numeric)"),
  },
  async ({ siteId, listId, itemId, fields, driveId }) => {
    const parsedFields = JSON.parse(fields);
    const result = await setDocumentMetadata(siteId, listId, itemId, parsedFields, driveId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "get_document_metadata",
  "Get all metadata fields of a document including custom columns. Accepts both drive item ID and numeric list item ID — if using drive item ID, provide driveId. Returns all field values including system fields and custom metadata.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    listId: z.string().describe("List or document library list ID"),
    itemId: z.string().describe("Document ID — either numeric list item ID or drive item ID"),
    driveId: z.string().optional().describe("Drive ID — required when itemId is a drive item ID (non-numeric)"),
  },
  async ({ siteId, listId, itemId, driveId }) => {
    const result = await getDocumentMetadata(siteId, listId, itemId, driveId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ PAGES ============

server.tool(
  "list_pages",
  "List all pages via Graph API. Returns page GUID, name, title, URL, and state. Use page GUID for publish_page, delete_page, add_quick_links. NOTE: Canvas tools (get/set_page_canvas_content) need NUMERIC item IDs from list_site_pages instead — GUIDs from this tool won't work there.",
  { siteId: z.string().describe("SharePoint site ID") },
  async ({ siteId }) => {
    const pages = await listPages(siteId);
    return { content: [{ type: "text", text: JSON.stringify(pages, null, 2) }] };
  }
);

server.tool(
  "create_page",
  "Create a new empty page in checkout/draft state. MUST call publish_page after to make it visible. 'name' becomes URL slug (e.g. 'my-page' → my-page.aspx). For pages with text sections, use create_page_with_content. For pages with web parts (Highlighted Content etc.), create empty page first, then set_page_canvas_content, then publish_page.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    title: z.string().describe("Page title"),
    name: z.string().describe("Page file name (without .aspx)"),
  },
  async ({ siteId, title, name }) => {
    const result = await createPage(siteId, title, name);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "create_page_with_content",
  "Create a page with HTML text sections. Draft state — MUST call publish_page after. Sections is a JSON string array. Layouts: oneColumn (12), twoColumns (6+6), threeColumns (4+4+4), oneThirdLeftColumn (4+8), oneThirdRightColumn (8+4), fullWidth (12). Example: [{\"layout\":\"twoColumns\",\"columns\":[{\"width\":6,\"html\":\"<h2>Left</h2>\"},{\"width\":6,\"html\":\"<h2>Right</h2>\"}]}]. NOTE: Only supports text/HTML. For web parts like Highlighted Content, use create_page + set_page_canvas_content instead.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    title: z.string().describe("Page title"),
    name: z.string().describe("Page file name (without .aspx)"),
    sections: z.string().describe("JSON array of sections: [{\"layout\":\"oneColumn\",\"columns\":[{\"width\":12,\"html\":\"<h2>Title</h2><p>Text</p>\"}]}]"),
  },
  async ({ siteId, title, name, sections }) => {
    const parsedSections = JSON.parse(sections);
    const result = await createPageWithContent(siteId, title, name, parsedSections);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "add_quick_links",
  "Add a Quick Links web part to a SharePoint page using PATCH. WARNING: this replaces the entire page canvas layout — any existing content on the page will be overwritten. Best used on empty pages created with create_page, or when Quick Links is the only content needed. For pages with existing content, use create_page_with_content instead. Page must be in checkout/draft state (not published).",
  {
    siteId: z.string().describe("SharePoint site ID"),
    pageId: z.string().describe("Page ID"),
    links: z.array(z.object({
      title: z.string().describe("Link title"),
      url: z.string().describe("Link URL"),
    })).describe("Array of links to add"),
  },
  async ({ siteId, pageId, links }) => {
    const result = await addQuickLinksWebPart(siteId, pageId, links);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "publish_page",
  "Publish a page to make it visible. MUST call after: create_page, create_page_with_content, or set_page_canvas_content — without publishing, the page appears EMPTY to users. This is the most commonly forgotten step.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    pageId: z.string().describe("Page ID"),
  },
  async ({ siteId, pageId }) => {
    const result = await publishPage(siteId, pageId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "delete_page",
  "PERMANENTLY delete a SharePoint page. This action is irreversible. Always confirm with the user before deleting. The page will be removed from Site Pages.",
  {
    siteId: z.string().describe("SharePoint site ID"),
    pageId: z.string().describe("Page ID"),
  },
  async ({ siteId, pageId }) => {
    const result = await deletePage(siteId, pageId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ NAVIGATION ============

server.tool(
  "get_navigation",
  "Get the top navigation menu links of a SharePoint site. Uses SharePoint REST API. Returns link ID, title, and URL for each navigation item. The siteUrl must be the full URL with https:// (e.g. https://contoso.sharepoint.com/sites/MySite).",
  {
    siteUrl: z.string().describe("Full SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/MySite)"),
  },
  async ({ siteUrl }) => {
    const result = await getNavigation(siteUrl);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "add_navigation_link",
  "Add a link to top navigation. REST API, delegated auth. siteUrl must include https://. url = link target (internal or external). Check get_navigation first to avoid duplicates.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    title: z.string().describe("Navigation link title"),
    url: z.string().describe("Navigation link URL"),
  },
  async ({ siteUrl, title, url }) => {
    const result = await addNavigationLink(siteUrl, title, url);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "delete_navigation_link",
  "Remove a link from the top navigation menu of a SharePoint site. Use get_navigation first to find the linkId. The siteUrl must be full URL with https://.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    linkId: z.number().describe("Navigation link ID (from get_navigation)"),
  },
  async ({ siteUrl, linkId }) => {
    const result = await deleteNavigationLink(siteUrl, linkId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ PERMISSIONS ============

server.tool(
  "get_permissions",
  "Get all SharePoint groups (Visitors, Members, Owners, custom) for a site. Returns group ID, title, and description. Use group ID with get_group_members to see who is in each group, or with add_user_to_group/remove_user_from_group to manage membership. Uses SharePoint REST API with delegated auth.",
  { siteUrl: z.string().describe("Full SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/MySite)") },
  async ({ siteUrl }) => {
    const result = await getSitePermissions(siteUrl);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "get_group_members",
  "Get all members of a SharePoint group. Use get_permissions first to find the groupId. Returns user ID, name, email, and login name for each member.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    groupId: z.number().describe("SharePoint group ID (from get_permissions)"),
  },
  async ({ siteUrl, groupId }) => {
    const result = await getGroupMembers(siteUrl, groupId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "add_user_to_group",
  "Add a user to a SharePoint group (Visitors=read, Members=edit, Owners=admin). Use get_permissions first to find the groupId. The user must have a valid M365 account. Always confirm with the user before changing permissions.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    groupId: z.number().describe("SharePoint group ID (from get_permissions)"),
    userEmail: z.string().describe("User email address (must be valid M365 account)"),
  },
  async ({ siteUrl, groupId, userEmail }) => {
    const result = await addUserToGroup(siteUrl, groupId, userEmail);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "remove_user_from_group",
  "Remove a user from a SharePoint group. Use get_group_members first to find the userId. This action is irreversible — always confirm with the user before removing permissions.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    groupId: z.number().describe("SharePoint group ID"),
    userId: z.number().describe("User ID (from get_group_members)"),
  },
  async ({ siteUrl, groupId, userId }) => {
    const result = await removeUserFromGroup(siteUrl, groupId, userId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ REST API — CANVAS CONTENT ============

server.tool(
  "get_page_canvas_content",
  "Read raw CanvasContent1 of a page via REST API. Returns full HTML including all web parts. Use to inspect existing page structure before modifying. pageItemId = numeric ID from list_site_pages (NOT GUID from list_pages). Uses delegated auth. ALWAYS read before calling set_page_canvas_content — understand what's there before replacing it.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/MySite)"),
    pageItemId: z.number().describe("Numeric item ID from Site Pages list (use list_site_pages to find it)"),
  },
  async ({ siteUrl, pageItemId }) => {
    const result = await getPageCanvasContent(siteUrl, pageItemId);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "set_page_canvas_content",
  "Replace entire page canvas via REST API. Overwrites ALL existing content. Uses delegated auth. pageItemId = numeric ID from list_site_pages. MUST call publish_page after — page reverts to draft and appears EMPTY without it. Canvas HTML rules: use pre-encoded entities directly (&#123; &#125; &#58; &quot;) — never JSON.stringify then replace (double-escaping on file round-trips). Omit titleHTML property (causes escaping corruption) — titles render from searchablePlainTexts. Read current canvas with get_page_canvas_content first.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    pageItemId: z.number().describe("Numeric item ID from Site Pages list"),
    canvasContent: z.string().describe("Raw HTML/JSON canvas content string — get format from get_page_canvas_content on an existing page"),
  },
  async ({ siteUrl, pageItemId, canvasContent }) => {
    const result = await setPageCanvasContent(siteUrl, pageItemId, canvasContent);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "copy_page",
  "Copy an existing SharePoint page to create a new one. Useful for using template pages — create one page with the desired layout in SharePoint editor, then copy it programmatically. Both source and target are file names in SitePages folder (e.g. 'template.aspx', 'new-page.aspx'). Uses delegated auth.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
    sourceFileName: z.string().describe("Source page file name (e.g. 'template.aspx')"),
    targetFileName: z.string().describe("Target page file name (e.g. 'new-page.aspx')"),
  },
  async ({ siteUrl, sourceFileName, targetFileName }) => {
    const result = await copyPage(siteUrl, sourceFileName, targetFileName);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  "list_site_pages",
  "List pages via REST API. Returns NUMERIC item IDs needed for canvas tools (get/set_page_canvas_content). Also returns title and file name. Uses delegated auth. NOTE: This gives numeric IDs for canvas operations. For page GUIDs (needed by publish_page, delete_page), use list_pages instead.",
  {
    siteUrl: z.string().describe("Full SharePoint site URL"),
  },
  async ({ siteUrl }) => {
    const result = await listSitePages(siteUrl);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  }
);

// ============ START SERVER ============

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("365center-mcp server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
