import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";

// Load .env — try parent directory first (local dev), fallback to env vars (Docker)
const __dirname = import.meta.dirname ?? path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.resolve(__dirname, "../../.env") });

const tenantId = process.env.AZURE_TENANT_ID!;
const clientId = process.env.AZURE_CLIENT_ID!;
const clientSecret = process.env.AZURE_CLIENT_SECRET!;

// ============ APP-ONLY AUTH (Graph API) ============

export const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ["https://graph.microsoft.com/.default"],
});

export const graphClient = Client.initWithMiddleware({
  authProvider,
});

// ============ DELEGATED AUTH (SharePoint REST API) ============

// Store token in user's home directory — works for Node, Docker (with volume), and plugins
const TOKEN_DIR = path.join(process.env.HOME || process.env.USERPROFILE || "/tmp", ".365center-mcp");
if (!fs.existsSync(TOKEN_DIR)) fs.mkdirSync(TOKEN_DIR, { recursive: true });
const TOKEN_CACHE_PATH = path.join(TOKEN_DIR, "token-cache.json");
const SHAREPOINT_DOMAIN = process.env.SHAREPOINT_DOMAIN || "";
const SP_SCOPES = `offline_access https://${SHAREPOINT_DOMAIN}/AllSites.FullControl`;

interface TokenCache {
  accessToken: string;
  refreshToken: string;
  expiresAt: number; // unix timestamp ms
}

function loadTokenCache(): TokenCache | null {
  try {
    if (fs.existsSync(TOKEN_CACHE_PATH)) {
      return JSON.parse(fs.readFileSync(TOKEN_CACHE_PATH, "utf-8"));
    }
  } catch {}
  return null;
}

function saveTokenCache(cache: TokenCache) {
  fs.writeFileSync(TOKEN_CACHE_PATH, JSON.stringify(cache, null, 2));
}

async function refreshAccessToken(refreshToken: string): Promise<TokenCache> {
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: SP_SCOPES,
  });

  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );

  if (!response.ok) {
    const err = await response.text();
    throw new Error(`Token refresh failed: ${err}`);
  }

  const data = await response.json() as any;
  const cache: TokenCache = {
    accessToken: data.access_token,
    refreshToken: data.refresh_token || refreshToken,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  saveTokenCache(cache);
  return cache;
}

// Device code polling — runs in background after user gets instructions
let deviceCodePollingPromise: Promise<TokenCache> | null = null;

function pollForDeviceCodeToken(deviceCode: string, interval: number, expiresIn: number): Promise<TokenCache> {
  return new Promise(async (resolve, reject) => {
    const pollInterval = (interval || 5) * 1000;
    const deadline = Date.now() + expiresIn * 1000;

    while (Date.now() < deadline) {
      await new Promise(r => setTimeout(r, pollInterval));

      const tokenResponse = await fetch(
        `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: clientId,
            grant_type: "urn:ietf:params:oauth:grant-type:device_code",
            device_code: deviceCode,
          }),
        }
      );

      const tokenData = await tokenResponse.json() as any;

      if (tokenData.access_token) {
        const cache: TokenCache = {
          accessToken: tokenData.access_token,
          refreshToken: tokenData.refresh_token,
          expiresAt: Date.now() + tokenData.expires_in * 1000,
        };
        saveTokenCache(cache);
        deviceCodePollingPromise = null;
        resolve(cache);
        return;
      }

      if (tokenData.error === "authorization_pending") continue;
      if (tokenData.error === "slow_down") {
        await new Promise(r => setTimeout(r, 5000));
        continue;
      }

      deviceCodePollingPromise = null;
      reject(new Error(`Device code auth failed: ${tokenData.error} — ${tokenData.error_description}`));
      return;
    }

    deviceCodePollingPromise = null;
    reject(new Error("Login timeout — user did not complete authentication in time"));
  });
}

// Throws a user-facing error with login instructions, starts background polling
async function startDeviceCodeFlow(): Promise<never> {
  if (!SHAREPOINT_DOMAIN) {
    throw new Error("SHAREPOINT_DOMAIN environment variable is required for delegated auth");
  }

  const codeResponse = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/devicecode`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        scope: SP_SCOPES,
      }),
    }
  );

  if (!codeResponse.ok) {
    const err = await codeResponse.text();
    throw new Error(`Device code request failed: ${err}`);
  }

  const codeData = await codeResponse.json() as any;
  const { device_code, user_code, verification_uri, expires_in, interval, message } = codeData;

  // Start polling in background
  deviceCodePollingPromise = pollForDeviceCodeToken(device_code, interval, expires_in);

  // Throw error with login instructions — Claude sees this and tells the user
  throw new Error(
    `LOGIN REQUIRED: ${message}\n\n` +
    `Go to: ${verification_uri}\n` +
    `Enter code: ${user_code}\n\n` +
    `After logging in, try your request again.`
  );
}

export async function getDelegatedToken(): Promise<string> {
  // 1. Check cache
  const cache = loadTokenCache();
  if (cache) {
    if (cache.expiresAt > Date.now() + 300000) {
      return cache.accessToken;
    }
    try {
      const refreshed = await refreshAccessToken(cache.refreshToken);
      return refreshed.accessToken;
    } catch {
      // Refresh failed — need new login
    }
  }

  // 2. If polling is already running, wait for it
  if (deviceCodePollingPromise) {
    const result = await deviceCodePollingPromise;
    return result.accessToken;
  }

  // 3. No cache, no polling — start device code flow (throws with instructions)
  await startDeviceCodeFlow();
  throw new Error("unreachable"); // startDeviceCodeFlow always throws
}

// ============ SHAREPOINT REST API HELPERS ============

// App-only token for SharePoint REST API (limited — doesn't work for navigation/permissions)
export async function getSharePointToken(): Promise<string> {
  const domain = process.env.SHAREPOINT_DOMAIN;
  if (!domain) throw new Error("SHAREPOINT_DOMAIN environment variable is required");
  const token = await credential.getToken(`https://${domain}/.default`);
  return token.token;
}

// Delegated token for SharePoint REST API (full access — navigation, permissions, CanvasContent1)
export async function getSharePointDelegatedToken(): Promise<string> {
  // Delegated token from Graph scopes also works for SharePoint REST API
  // because Sites.FullControl.All grants access to SharePoint endpoints
  return getDelegatedToken();
}

export async function callSharePointRest(siteUrl: string, apiPath: string, method: string = "GET", body?: unknown, extraHeaders?: Record<string, string>): Promise<unknown> {
  const token = await getSharePointDelegatedToken();
  const url = `${siteUrl}${apiPath}`;

  const headers: Record<string, string> = {
    "Authorization": `Bearer ${token}`,
    "Accept": "application/json;odata=verbose",
    "Content-Type": "application/json;odata=verbose",
    ...extraHeaders,
  };

  // SharePoint REST API uses POST + X-HTTP-Method for MERGE/PUT/PATCH
  let httpMethod = method;
  if (method === "MERGE" || method === "PUT" || method === "PATCH") {
    headers["X-HTTP-Method"] = method;
    headers["IF-MATCH"] = "*";
    httpMethod = "POST";
  }

  const options: RequestInit = { method: httpMethod, headers };
  if (body) {
    options.body = JSON.stringify(body);
  }

  const response = await fetch(url, options);
  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`SharePoint REST API error ${response.status}: ${errorText}`);
  }

  const text = await response.text();
  return text ? JSON.parse(text) : null;
}
