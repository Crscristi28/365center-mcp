# 365center-mcp

MCP server for Microsoft 365 and SharePoint. Gives Claude (and any MCP client) full read-write access to SharePoint sites, documents, pages, metadata, navigation, and permissions via Microsoft Graph API and SharePoint REST API.

Built for manufacturing companies that manage factory documentation in SharePoint.

## Features

**32 tools** across 7 categories:

### Sites
- `list_sites` — List all SharePoint sites in the tenant
- `get_site` — Get site by URL
- `get_site_by_id` — Get site by ID

### Documents
- `list_document_libraries` — List document libraries (drives)
- `list_documents` — List documents with both driveItemId and listItemId
- `upload_document` — Upload files to SharePoint
- `search_documents` — Search across documents
- `delete_document` — Delete a document
- `create_folder` — Create folders
- `get_document_versions` — Version history (audit trail)

### Metadata
- `list_columns` — List custom metadata columns
- `create_choice_column` — Create choice/dropdown columns
- `create_text_column` — Create text columns
- `get_document_metadata` — Read document metadata
- `set_document_metadata` — Set metadata on documents

### Pages
- `list_pages` — List all pages
- `create_page` — Create empty page
- `create_page_with_content` — Create page with sections and HTML content
- `add_quick_links` — Add Quick Links web part
- `publish_page` — Publish a draft page
- `delete_page` — Delete a page

### Pages (REST API)
- `list_site_pages` — List pages with numeric IDs
- `get_page_canvas_content` — Read raw page content (CanvasContent1)
- `set_page_canvas_content` — Write raw page content (supports Highlighted Content and any web part)
- `copy_page` — Copy a page as template

### Navigation
- `get_navigation` — Read top navigation menu
- `add_navigation_link` — Add link to navigation
- `delete_navigation_link` — Remove link from navigation

### Permissions
- `get_permissions` — List SharePoint groups (Visitors, Members, Owners)
- `get_group_members` — List members of a group
- `add_user_to_group` — Add user to a group
- `remove_user_from_group` — Remove user from a group

## Authentication

365center-mcp uses two authentication methods:

- **App-only (Client Credentials)** — for Graph API operations (sites, documents, pages, metadata). Works automatically with Azure App Registration credentials.
- **Delegated (Device Code Flow)** — for SharePoint REST API operations (navigation, permissions, Highlighted Content). On first use, you'll be prompted to sign in via https://login.microsoft.com/device with a one-time code. Token is cached and refreshed automatically.

## Prerequisites

- Microsoft 365 tenant with SharePoint
- Azure App Registration with:
  - **Application permissions:** Sites.ReadWrite.All, Sites.FullControl.All, Files.ReadWrite.All, Sites.Manage.All
  - **Delegated permissions:** Sites.ReadWrite.All, Sites.FullControl.All, offline_access
  - **SharePoint permissions:** Sites.FullControl.All
  - **"Allow public client flows" enabled** (for device code auth)

## Installation

### Docker (recommended)

```bash
docker pull 365center-mcp:latest

# Claude Desktop config (claude_desktop_config.json):
{
  "mcpServers": {
    "365center-mcp": {
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-e", "AZURE_TENANT_ID=your-tenant-id",
        "-e", "AZURE_CLIENT_ID=your-client-id",
        "-e", "AZURE_CLIENT_SECRET=your-client-secret",
        "-e", "SHAREPOINT_DOMAIN=contoso.sharepoint.com",
        "-v", "~/.365center-mcp:/home/mcp/.365center-mcp",
        "365center-mcp:latest"
      ]
    }
  }
}
```

### Node.js

```bash
git clone https://github.com/Crscristi28/365center-mcp.git
cd 365center-mcp/mcp-server
npm install
npm run build

# Create .env file:
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
SHAREPOINT_DOMAIN=contoso.sharepoint.com

# Run:
node dist/index.js
```

### Claude Desktop config (Node.js)

```json
{
  "mcpServers": {
    "365center-mcp": {
      "command": "node",
      "args": ["/path/to/365center-mcp/mcp-server/dist/index.js"],
      "env": {
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "SHAREPOINT_DOMAIN": "contoso.sharepoint.com"
      }
    }
  }
}
```

## Page Layouts

When using `create_page_with_content`, available section layouts:

| Layout | Columns | Widths |
|--------|---------|--------|
| `oneColumn` | 1 | 12 |
| `twoColumns` | 2 | 6 + 6 |
| `threeColumns` | 3 | 4 + 4 + 4 |
| `oneThirdLeftColumn` | 2 | 4 + 8 |
| `oneThirdRightColumn` | 2 | 8 + 4 |
| `fullWidth` | 1 | 12 |

## Supported Web Parts

When creating pages via Graph API, these standard web parts are supported:

| Web Part | Type ID |
|----------|---------|
| Bing Maps | `e377ea37-9047-43b9-8cdb-a761be2f8e09` |
| Button | `0f087d7f-520e-42b7-89c0-496aaf979d58` |
| Call To Action | `df8e44e7-edd5-46d5-90da-aca1539313b8` |
| Divider | `2161a1c6-db61-4731-b97c-3cdb303f7cbb` |
| Document Embed | `b7dd04e1-19ce-4b24-9132-b60a1c2b910d` |
| Image | `d1d91016-032f-456d-98a4-721247c305e8` |
| Image Gallery | `af8be689-990e-492a-81f7-ba3e4cd3ed9c` |
| Link Preview | `6410b3b6-d440-4663-8744-378976dc041e` |
| Org Chart | `e84a8ca2-f63c-4fb9-bc0b-d8eef5ccb22b` |
| People | `7f718435-ee4d-431c-bdbf-9c4ff326f46e` |
| Quick Links | `c70391ea-0b10-4ee9-b2b4-006d3fcad0cd` |
| Spacer | `8654b779-4886-46d4-8ffb-b5ed960ee986` |
| YouTube Embed | `544dd15b-cf3c-441b-96da-004d5a8cea1d` |
| Title Area | `cbe7b0a9-3504-44dd-a3a3-0e5cacd07788` |

For **Highlighted Content** and other unsupported web parts, use the REST API tools (`get_page_canvas_content` / `set_page_canvas_content`).

## Security

365center-mcp is designed for enterprise environments:

- **No data leaves your tenant** — all API calls go directly from the MCP server to Microsoft Graph API and SharePoint REST API. No third-party servers, no telemetry, no analytics.
- **Azure AD authentication** — uses your organization's existing Azure App Registration with OAuth 2.0. Credentials are never stored in the codebase.
- **Principle of least privilege** — app-only auth for read/write operations, delegated auth only when required (navigation, permissions). You control exactly which permissions are granted.
- **Device Code Flow** — delegated auth uses Microsoft's standard device code flow (same as Azure CLI, GitHub CLI). No localhost servers, no open ports, no redirect URIs needed.
- **Token security** — refresh tokens are stored locally in `~/.365center-mcp/token-cache.json` with filesystem permissions. Tokens are never transmitted to third parties.
- **Docker isolation** — runs as non-root user (`mcp`) inside the container. Token cache is mounted as a volume, not baked into the image.
- **No secrets in Docker image** — credentials are passed as environment variables at runtime, never included in the build.
- **MCP stdio transport** — communicates via stdin/stdout only. No HTTP server, no exposed ports, no network attack surface.
- **BSL license** — source code is fully auditable. Your security team can review every line before deployment.

### Recommended deployment

For production environments:

1. Create a dedicated Azure App Registration for 365center-mcp
2. Grant only the permissions your workflows need
3. Use Docker with volume mount for token persistence
4. Store credentials in your organization's secret manager (Azure Key Vault, HashiCorp Vault, etc.)
5. Restrict App Registration to specific SharePoint sites if possible

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID |
| `AZURE_CLIENT_ID` | Yes | App registration client ID |
| `AZURE_CLIENT_SECRET` | Yes | App registration client secret |
| `SHAREPOINT_DOMAIN` | Yes | SharePoint domain (e.g. `contoso.sharepoint.com`) |

## Architecture

```
Claude Desktop / Claude Code / API
        │
        │ stdio (stdin/stdout)
        │
  365center-mcp (MCP Server)
        │
        ├── Microsoft Graph API (v1.0)
        │   └── Sites, Documents, Pages, Metadata
        │
        └── SharePoint REST API
            └── Navigation, Permissions, CanvasContent1
```

- **Graph API** uses app-only auth (Client Credentials) — no user interaction needed
- **REST API** uses delegated auth (Device Code Flow) — one-time login, then automatic token refresh

## Docker Details

The Docker image runs as non-root user `mcp` and communicates only via stdio.

```bash
# Build
docker build -t 365center-mcp:latest ./mcp-server

# Run standalone (for testing)
docker run -i --rm \
  -e AZURE_TENANT_ID=your-tenant-id \
  -e AZURE_CLIENT_ID=your-client-id \
  -e AZURE_CLIENT_SECRET=your-client-secret \
  -e SHAREPOINT_DOMAIN=contoso.sharepoint.com \
  -v ~/.365center-mcp:/home/mcp/.365center-mcp \
  365center-mcp:latest
```

The `-v` flag mounts the token cache directory so delegated auth tokens persist across container restarts. Without it, you'd need to re-authenticate every time the container starts.

## Token Storage

Delegated auth tokens are stored in `~/.365center-mcp/token-cache.json`. This file contains:
- Access token (expires in ~1 hour, refreshed automatically)
- Refresh token (long-lived, used to get new access tokens)

For Docker, mount `~/.365center-mcp` as a volume. The token file has the same security sensitivity as your Azure credentials — protect it accordingly.

## License

[Business Source License 1.1](LICENSE) — Free for internal use, testing, and development. Commercial use that competes with 365center-mcp requires written permission. Converts to MIT on 2030-04-08.

## Author

Cristian Bucioacă — [cristianb.cz](https://cristianb.cz) — [info@cristianb.cz](mailto:info@cristianb.cz)
