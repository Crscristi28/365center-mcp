# 365center-mcp

**MCP server for Microsoft 365 / SharePoint — 35 tools for full read/write access**

Available on [GitHub](https://github.com/Crscristi28/365center-mcp) · [npm](https://www.npmjs.com/package/365center-mcp) · [Docker Hub](https://hub.docker.com/r/crscristi28/365center-mcp) · [cristianb.cz](https://cristianb.cz)

[![npm version](https://img.shields.io/npm/v/365center-mcp.svg)](https://www.npmjs.com/package/365center-mcp)
[![Docker Pulls](https://img.shields.io/docker/pulls/crscristi28/365center-mcp.svg)](https://hub.docker.com/r/crscristi28/365center-mcp)
[![License](https://img.shields.io/badge/license-BUSL--1.1-blue.svg)](https://github.com/Crscristi28/365center-mcp/blob/main/LICENSE)
[![Node](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)

Full visual walkthrough: **[Setup Guide PDF](https://github.com/Crscristi28/365center-mcp/blob/main/SETUP-GUIDE.pdf)** — screenshots for every Azure setup step and all 3 installation methods.

---

## Table of Contents

- [What is this?](#what-is-this)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Azure Setup (step-by-step)](#azure-setup-step-by-step)
- [Installation](#installation)
  - [Option 1: Docker (recommended)](#option-1-docker-recommended)
  - [Option 2: npx (easiest)](#option-2-npx-easiest)
  - [Option 3: Node.js from source](#option-3-nodejs-from-source)
  - [Using Claude Code instead of Claude Desktop](#using-claude-code-instead-of-claude-desktop)
- [First-time login (device code flow)](#first-time-login-device-code-flow)
- [Configuration](#configuration)
- [Usage examples](#usage-examples)
- [Architecture](#architecture)
- [Security](#security)
- [Supported Page Layouts](#supported-page-layouts)
- [Supported Web Parts](#supported-web-parts)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Author](#author)

---

## What is this?

`365center-mcp` is a Model Context Protocol (MCP) server that gives Claude — and any other MCP-compatible AI client — full read/write access to Microsoft 365 SharePoint sites.

It exposes **35 tools** covering SharePoint sites, document libraries, documents, pages, metadata columns, navigation, and permissions. Claude can list sites, upload and download files, tag documents, create and publish pages, build navigation menus, manage permissions, and more — all through a single MCP connection.

Built for manufacturing companies managing factory documentation in SharePoint, but works with any Microsoft 365 tenant.

**Typical use cases:**
- Automate SharePoint document workflows with Claude
- Build and maintain intranet sites conversationally
- Manage metadata, tagging, and permissions at scale
- Replace manual document copying across folders with metadata-driven pages

---

## Features

**35 tools** across 7 categories. All tools use Microsoft Graph API or SharePoint REST API — no middlemen.

### Sites (4 tools)
- `list_sites` — List all SharePoint sites in the tenant
- `get_site` — Get site by URL
- `get_site_by_id` — Get site by ID
- `create_site` — Create a new Communication or Team site

### Documents (9 tools)
- `list_document_libraries` — List document libraries (drives)
- `list_documents` — List documents with both driveItemId and listItemId
- `upload_document` — Upload a file to SharePoint (auto session upload for files over 4 MB)
- `upload_documents` — Upload multiple files with optional metadata in one call (max 30 per call)
- `download_document` — Download files from SharePoint to a local path
- `search_documents` — Search across documents
- `delete_document` — Delete a document
- `create_folder` — Create folders
- `get_document_versions` — Version history (audit trail)

### Metadata (5 tools)
- `list_columns` — List custom metadata columns
- `create_choice_column` — Create choice/dropdown columns (single or multi-select)
- `create_text_column` — Create text columns
- `get_document_metadata` — Read document metadata
- `set_document_metadata` — Set metadata on documents

### Pages — Graph API (6 tools)
- `list_pages` — List all pages
- `create_page` — Create empty page
- `create_page_with_content` — Create page with sections and HTML content
- `add_quick_links` — Add Quick Links web part
- `publish_page` — Publish a draft page
- `delete_page` — Delete a page

### Pages — SharePoint REST API (4 tools)
- `list_site_pages` — List pages with numeric IDs
- `get_page_canvas_content` — Read raw page content (CanvasContent1)
- `set_page_canvas_content` — Write raw page content (supports Highlighted Content and any web part)
- `copy_page` — Copy a page as template

### Navigation (3 tools)
- `get_navigation` — Read top navigation menu
- `add_navigation_link` — Add link to navigation
- `delete_navigation_link` — Remove link from navigation

### Permissions (4 tools)
- `get_permissions` — List SharePoint groups (Visitors, Members, Owners)
- `get_group_members` — List members of a group
- `add_user_to_group` — Add user to a group
- `remove_user_from_group` — Remove user from a group

---

## Prerequisites

Before installing `365center-mcp`, you need the following.

### 1. An MCP client

You need a client that can connect to MCP servers. Any of these works:

- **Claude Desktop** (Mac / Windows) — download from https://claude.ai/download
- **Claude Code** (CLI) — install from https://docs.claude.com/claude-code
- Any other MCP-compatible client

### 2. One of these runtimes (choose based on your install method)

Pick ONE installation method below, and install the matching runtime:

| If you want to use... | You need to install |
|---|---|
| **Docker** (Option 1, recommended) | [Docker Desktop](https://www.docker.com/products/docker-desktop) |
| **npx** (Option 2, easiest) | [Node.js 18 or newer](https://nodejs.org) |
| **Node.js from source** (Option 3, for developers) | [Node.js 18 or newer](https://nodejs.org) + [Git](https://git-scm.com/downloads) |

### 3. A Microsoft 365 tenant with admin access

You need a Microsoft 365 Business tenant — this is the account your company uses for Outlook, Teams, and SharePoint. `365center-mcp` does not work with personal Microsoft accounts (outlook.com, hotmail.com) — only work/school accounts on a real M365 Business tenant.

**Getting a tenant:**
- **Most likely case:** Your company already has one. Ask your IT admin.
- **If you don't have one:** Sign up for any Microsoft 365 Business plan that includes SharePoint (Business Basic, Business Standard, or Business Premium) at https://www.microsoft.com/microsoft-365/business/compare-all-plans. Microsoft offers a 1-month free trial on these plans, but a credit card is required and the trial automatically converts to a paid subscription if not cancelled.
- **Microsoft 365 Developer Program** (free tenant, 25 licenses) is an option only if you have an active Visual Studio Professional or Enterprise subscription — as of 2025, it is no longer open to personal accounts.

Your M365 tenant automatically includes **Microsoft Entra ID** (formerly Azure Active Directory). You do NOT need a separate Azure subscription — App Registrations are free and included with every M365 tenant.

**You need all of these before continuing:**

- Microsoft 365 tenant with SharePoint Online included in the plan
- **Global Administrator** or **Privileged Role Administrator** role on the tenant (required to grant admin consent in Azure)
- At least one SharePoint site you want to manage
- Access to https://portal.azure.com using the same M365 credentials

---

## Azure Setup

`365center-mcp` authenticates to Microsoft using an Azure App Registration in your tenant's Microsoft Entra ID. You create one app, grant it permissions, generate a client secret, and collect 4 values for the config. About 10 minutes the first time.

**Full step-by-step instructions with screenshots are in the [Setup Guide PDF](SETUP-GUIDE.pdf)** — recommended for first-time users.

**TL;DR for experienced Azure admins:**

1. [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration** → single tenant, no redirect URI
2. **API permissions** — add and grant admin consent for:
   - **Microsoft Graph (Application):** `Sites.ReadWrite.All`, `Sites.FullControl.All`, `Sites.Manage.All`, `Files.ReadWrite.All`
   - **Microsoft Graph (Delegated):** `Sites.ReadWrite.All`, `Sites.FullControl.All`, `offline_access`
   - **SharePoint (Delegated):** `AllSites.FullControl`
3. **Authentication** → **Allow public client flows: Yes** (required for device code flow)
4. **Certificates & secrets** → **New client secret** → copy the **Value** column immediately

You will end up with 4 values you need for the config in the next step:

| Variable | Where to find it |
|---|---|
| `AZURE_TENANT_ID` | App Registration → Overview → Directory (tenant) ID |
| `AZURE_CLIENT_ID` | App Registration → Overview → Application (client) ID |
| `AZURE_CLIENT_SECRET` | The **Value** from the client secret you created |
| `SHAREPOINT_DOMAIN` | Your SharePoint domain without `https://`, e.g. `contoso.sharepoint.com` |

> **Need a Global Administrator** (or Privileged Role Administrator) to grant admin consent. Ask your IT admin if you don't have this role.

---

## Installation

Pick the option that matches your setup.

### Option 1: Docker (recommended)

**Requires:** Docker Desktop running, Claude Desktop installed.

**1. Quit Claude Desktop completely** (quit from the menu bar, not just close the window). Editing the config while Claude Desktop is running means your changes won't be picked up until a full restart.

**2. Pull the image:**

```bash
docker pull crscristi28/365center-mcp:latest
```

Or use the Docker Desktop GUI: **Docker Hub** tab → search `365center-mcp` → click the result by `crscristi28` → **Pull**. No terminal needed. See the [Setup Guide PDF](SETUP-GUIDE.pdf) for full steps.

**3. Add to your Claude Desktop config.**

Open `claude_desktop_config.json`:
- **Mac:** `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Add this (replace `<your-server-name>` with any name you like, e.g. `sharepoint`):

```json
{
  "mcpServers": {
    "<your-server-name>": {
      "command": "docker",
      "args": [
        "run", "-i", "--rm",
        "-e", "AZURE_TENANT_ID=your-tenant-id",
        "-e", "AZURE_CLIENT_ID=your-client-id",
        "-e", "AZURE_CLIENT_SECRET=your-client-secret",
        "-e", "SHAREPOINT_DOMAIN=your-domain.sharepoint.com",
        "-v", "/Users/YOUR_USERNAME/.365center-mcp:/home/mcp/.365center-mcp",
        "crscristi28/365center-mcp:latest"
      ]
    }
  }
}
```

Replace the four `your-*` values with what you collected in Step 8 of the Azure setup. Replace `YOUR_USERNAME` with your actual macOS/Windows username.

On Windows, use `C:\\Users\\YOUR_USERNAME\\.365center-mcp:/home/mcp/.365center-mcp` as the volume mount.

**4. Open Claude Desktop.** The server will appear in the MCP menu with all 35 tools loaded.

> **Why the volume mount?** The server caches delegated auth tokens in `~/.365center-mcp/token-cache.json`. Without the volume, you would need to re-authenticate every time Docker restarts.

### Option 2: npx (easiest)

**Requires:** Node.js 18+, Claude Desktop installed.

This is the simplest method — you don't download anything yourself. Just edit your Claude Desktop config, and Claude Desktop will call `npx` automatically, which downloads `365center-mcp` from npm the first time it runs. Add this to your config:

```json
{
  "mcpServers": {
    "<your-server-name>": {
      "command": "npx",
      "args": ["-y", "365center-mcp"],
      "env": {
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "SHAREPOINT_DOMAIN": "your-domain.sharepoint.com"
      }
    }
  }
}
```

Restart Claude Desktop. The first start takes 15–30 seconds longer than usual — that is `npx` downloading `365center-mcp` from npm. Subsequent starts are instant because the package is cached locally.

> **Note:** The npx config has a different structure from the Docker config — env vars go in a separate `env` object, not as `-e` flags in `args`. If you are switching from Docker to npx (or vice versa), replace the entire entry, don't just change the `command` field.

### Option 3: Node.js from source

**Requires:** Node.js 18+, Git, Claude Desktop installed. For developers who want to modify the code.

**1. Clone and build:**

```bash
git clone https://github.com/Crscristi28/365center-mcp.git
cd 365center-mcp/mcp-server
npm install
npm run build
```

This installs all runtime dependencies automatically:
- `@modelcontextprotocol/sdk` — official MCP SDK
- `@azure/identity` — Azure authentication
- `@microsoft/microsoft-graph-client` — Microsoft Graph API client
- `dotenv` — environment variable loading

**2. Create a `.env` file** in `mcp-server/` with your Azure credentials:

```
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
SHAREPOINT_DOMAIN=your-domain.sharepoint.com
```

**3. Test it runs:**

```bash
node dist/index.js
```

The server starts and waits for MCP messages on stdin. Press Ctrl+C to stop.

**4. Add to Claude Desktop config:**

```json
{
  "mcpServers": {
    "<your-server-name>": {
      "command": "node",
      "args": ["/absolute/path/to/mcp-server/dist/index.js"],
      "env": {
        "AZURE_TENANT_ID": "your-tenant-id",
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_CLIENT_SECRET": "your-client-secret",
        "SHAREPOINT_DOMAIN": "your-domain.sharepoint.com"
      }
    }
  }
}
```

Replace `/absolute/path/to/mcp-server/dist/index.js` with the full path on your machine. Restart Claude Desktop.

### Using Claude Code instead of Claude Desktop

Claude Code is Anthropic's CLI client. Instead of editing a JSON config, MCP servers are added with `claude mcp add`.

**Easiest way — let Claude Code do it.** Once Claude Code is installed, paste this into your Claude Code session (fill in your Azure values from Azure Setup Step 8):

```
Please install the 365center-mcp MCP server for me.
You can find it on npm as "365center-mcp" or at
github.com/Crscristi28/365center-mcp. Use the npx method.

My Azure credentials:
AZURE_TENANT_ID=...
AZURE_CLIENT_ID=...
AZURE_CLIENT_SECRET=...
SHAREPOINT_DOMAIN=...
```

Claude Code will run the right `claude mcp add` command for you. For the manual command and Docker / source variants, see the [Setup Guide PDF](SETUP-GUIDE.pdf).

---

## First-time login (device code flow)

Most tools work immediately with app-only authentication. But some features (navigation, permissions, Highlighted Content web part) require **delegated authentication** — meaning you need to sign in as a real user.

The first time you call one of these tools, Claude will show you a message like:

```
LOGIN REQUIRED: To sign in, use a web browser to open the page
https://microsoft.com/devicelogin and enter the code ABC123XYZ to authenticate.
```

**What to do:**

1. Open https://microsoft.com/devicelogin in any browser
2. Enter the code shown in the message
3. Sign in with your Microsoft 365 account
4. Close the browser tab
5. Ask Claude to try the same action again — it will succeed

The login is cached in `~/.365center-mcp/token-cache.json` with a refresh token, so you only have to do this once. The server automatically refreshes tokens in the background.

---

## Configuration

| Variable | Required | Description |
|---|---|---|
| `AZURE_TENANT_ID` | Yes | Azure AD tenant ID (from App Registration Overview) |
| `AZURE_CLIENT_ID` | Yes | Application (client) ID (from App Registration Overview) |
| `AZURE_CLIENT_SECRET` | Yes | Client secret **Value** (from Certificates & secrets) |
| `SHAREPOINT_DOMAIN` | Yes | Your SharePoint domain, e.g. `contoso.sharepoint.com` (no `https://`) |

---

## Usage examples

Once connected, you can ask Claude things like:

- *"List all SharePoint sites in my tenant"*
- *"Upload this file to the Documents library on the DocCenter site, and tag it with Oblast=Production and WS=WS1"*
- *"Create a new page called 'Production Line 1' with three sections and add Quick Links to the related work stations"*
- *"Who has access to the Finance site? Add alice@contoso.com as a Member"*
- *"Search all documents containing 'safety procedure' and list their version history"*
- *"Add a Highlighted Content web part to the Home page showing documents tagged with WS1"*

Claude will call the appropriate tools automatically. You don't need to know the tool names.

> **Note on file uploads:** `upload_document` reads files from the local filesystem where the MCP server runs. It works in Claude Desktop and Claude Code (server runs on your machine), but not in the claude.ai web app (no local file access).

---

## Architecture

```
Claude Desktop / Claude Code / any MCP client
        │
        │  stdio (stdin/stdout)
        │
  365center-mcp (MCP Server)
        │
        ├── Microsoft Graph API (v1.0)
        │     Sites, Documents, Pages, Metadata
        │     Auth: App-only (Client Credentials)
        │
        └── SharePoint REST API
              Navigation, Permissions, CanvasContent1
              Auth: Delegated (Device Code Flow)
```

- **Graph API** uses app-only auth — no user interaction, works in headless environments
- **REST API** uses delegated auth — one-time device code login, then automatic token refresh
- Both auth flows share the same Azure App Registration

---

## Security

`365center-mcp` is built for enterprise environments where security matters.

- **No data leaves your tenant** — all API calls go directly from the server to Microsoft. No third-party servers, no telemetry, no analytics.
- **Azure AD authentication** — uses your own App Registration with OAuth 2.0. Credentials never stored in the codebase.
- **Principle of least privilege** — app-only auth for most operations, delegated auth only where required.
- **Device Code Flow** — delegated auth uses Microsoft's standard device code flow (same as Azure CLI and GitHub CLI). No localhost servers, no open ports, no redirect URIs.
- **Local token storage** — refresh tokens stored in `~/.365center-mcp/token-cache.json` with filesystem permissions.
- **Docker isolation** — runs as non-root user (`mcp`) inside the container.
- **No secrets in the Docker image** — credentials passed as environment variables at runtime.
- **stdio transport only** — no HTTP server, no open ports, no network attack surface.
- **Auditable source** — BUSL-1.1 license, source fully available for review.

### Recommended production deployment

1. Create a dedicated App Registration for `365center-mcp`
2. Grant only the permissions your workflows actually need
3. Use Docker with a mounted volume for token persistence
4. Store credentials in a secret manager (Azure Key Vault, HashiCorp Vault, 1Password, etc.)
5. Restrict the App Registration to specific SharePoint sites when possible (via Sites.Selected)

---

## Supported Page Layouts

When using `create_page_with_content`, these section layouts are available:

| Layout | Columns | Widths |
|---|---|---|
| `oneColumn` | 1 | 12 |
| `twoColumns` | 2 | 6 + 6 |
| `threeColumns` | 3 | 4 + 4 + 4 |
| `oneThirdLeftColumn` | 2 | 4 + 8 |
| `oneThirdRightColumn` | 2 | 8 + 4 |
| `fullWidth` | 1 | 12 |

---

## Supported Web Parts

The Graph API `create_page_with_content` tool supports these standard web parts:

| Web Part | Type ID |
|---|---|
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
| Title Area | `cbe7b0a9-3504-44dd-a3a3-0e5cacd07788` |
| YouTube Embed | `544dd15b-cf3c-441b-96da-004d5a8cea1d` |

For **Highlighted Content** and any other web part not in this list, use the REST API tools (`get_page_canvas_content` and `set_page_canvas_content`) — they can read and write any web part including Highlighted Content.

---

## Troubleshooting

### "Insufficient privileges to complete the operation"
You missed an API permission or admin consent was not granted. Go back to **Azure Setup Step 5** and make sure **Grant admin consent** shows a green checkmark next to every permission.

### "AADSTS7000215: Invalid client secret provided"
Your `AZURE_CLIENT_SECRET` is wrong or expired. Go back to **Azure Setup Step 7** and create a new secret. Make sure you copied the **Value** column, not the **Secret ID**.

### "SHAREPOINT_DOMAIN environment variable is required"
You forgot to set `SHAREPOINT_DOMAIN`. It should be your SharePoint domain only, without `https://` and without any path — e.g. `contoso.sharepoint.com`.

### "LOGIN REQUIRED: To sign in, use a web browser..."
This is expected on first use of delegated auth. See [First-time login (device code flow)](#first-time-login-device-code-flow).

### "Device code auth failed: authorization_declined"
You declined the sign-in prompt in the browser, or signed in with the wrong account. Try again with an account that has access to the target SharePoint site.

### "Token refresh failed"
Your refresh token expired or was revoked. Delete `~/.365center-mcp/token-cache.json` and trigger the device code flow again by calling any tool that needs delegated auth.

### "Allow public client flows" is not available
You are probably looking at a Personal Microsoft account, not a work/school account. Device code flow requires a work/school account (M365 tenant).

### Docker: tokens lost on every restart
You forgot the `-v` volume mount. Without it, the container has no place to persist the token cache. See **Installation → Option 1**.

### Claude Desktop doesn't see the server
- Check your JSON syntax — a missing comma or bracket silently breaks the config
- Restart Claude Desktop fully (quit from the menu, not just close the window)
- Check Claude Desktop logs: **Help → View logs** (Mac) or `%APPDATA%\Claude\logs` (Windows)

---

## Contributing

Issues and pull requests welcome at https://github.com/Crscristi28/365center-mcp.

When filing a bug, please include:
- Installation method (Docker / npx / source)
- Node.js version (if applicable)
- The exact error message
- Whether it happens on the first call or after working for a while

**Note on pull requests:** Because `365center-mcp` is released under the Business Source License 1.1, all contributions must be reviewed before merging. By opening a pull request, you agree that your contribution may be distributed under the BSL 1.1 license and the future MIT conversion (on 2030-04-08). For larger changes, please open an issue to discuss the approach first.

---

## License

[Business Source License 1.1](LICENSE) — Free for internal use, testing, development, and non-commercial purposes. Commercial use that competes with `365center-mcp` requires written permission. Automatically converts to MIT on **2030-04-08**.

---

## Links

- **GitHub:** https://github.com/Crscristi28/365center-mcp
- **npm:** https://www.npmjs.com/package/365center-mcp
- **Docker Hub:** https://hub.docker.com/r/crscristi28/365center-mcp
- **Website:** https://cristianb.cz

---

## Author

**Cristian Bucioacă** — [cristianb.cz](https://cristianb.cz) — [info@cristianb.cz](mailto:info@cristianb.cz)

Building Microsoft 365 automation and SharePoint solutions for manufacturing.
