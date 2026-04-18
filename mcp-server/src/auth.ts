import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import { Entry } from "@napi-rs/keyring";
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

// Refresh token storage: OS keyring (macOS Keychain / Windows Credential Manager / Linux libsecret)
// with file fallback (chmod 600) for headless / Docker / keyring-less environments.
// Access token is kept in-memory only — never written to disk.
const KEYRING_SERVICE = "365center-mcp";
const KEYRING_ACCOUNT = "refresh-token";
const TOKEN_DIR = path.join(process.env.HOME || process.env.USERPROFILE || "/tmp", ".365center-mcp");
const TOKEN_FILE = path.join(TOKEN_DIR, "refresh-token");
const SHAREPOINT_DOMAIN = process.env.SHAREPOINT_DOMAIN || "";
const SP_SCOPES = `offline_access https://${SHAREPOINT_DOMAIN}/AllSites.FullControl`;

let cachedAccessToken: { token: string; expiresAt: number } | null = null;

function loadRefreshToken(): string | null {
  try {
    const v = new Entry(KEYRING_SERVICE, KEYRING_ACCOUNT).getPassword();
    if (v) return v;
  } catch {}
  try {
    if (fs.existsSync(TOKEN_FILE)) {
      return fs.readFileSync(TOKEN_FILE, "utf-8").trim() || null;
    }
  } catch {}
  return null;
}

function saveRefreshToken(token: string): void {
  try {
    new Entry(KEYRING_SERVICE, KEYRING_ACCOUNT).setPassword(token);
    return;
  } catch {}
  try {
    if (!fs.existsSync(TOKEN_DIR)) {
      fs.mkdirSync(TOKEN_DIR, { recursive: true });
    }
    try { fs.chmodSync(TOKEN_DIR, 0o700); } catch {}
    fs.writeFileSync(TOKEN_FILE, token);
    try { fs.chmodSync(TOKEN_FILE, 0o600); } catch {}
  } catch {}
}

function clearRefreshToken(): void {
  try { new Entry(KEYRING_SERVICE, KEYRING_ACCOUNT).deletePassword(); } catch {}
  try { if (fs.existsSync(TOKEN_FILE)) fs.unlinkSync(TOKEN_FILE); } catch {}
  cachedAccessToken = null;
}

async function refreshAccessToken(refreshToken: string): Promise<string> {
  // Device code flow uses a public client — client_secret must NOT be sent
  // (Microsoft returns AADSTS700025 if included).
  const body = new URLSearchParams({
    client_id: clientId,
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
  const accessToken = data.access_token as string;
  const newRefreshToken = (data.refresh_token as string) || refreshToken;
  cachedAccessToken = {
    token: accessToken,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  if (newRefreshToken !== refreshToken) {
    saveRefreshToken(newRefreshToken);
  }
  return accessToken;
}

// Device code polling — runs in background after user gets instructions
let deviceCodePollingPromise: Promise<string> | null = null;

function pollForDeviceCodeToken(deviceCode: string, interval: number, expiresIn: number): Promise<string> {
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
        cachedAccessToken = {
          token: tokenData.access_token,
          expiresAt: Date.now() + tokenData.expires_in * 1000,
        };
        if (tokenData.refresh_token) {
          saveRefreshToken(tokenData.refresh_token);
        }
        deviceCodePollingPromise = null;
        resolve(tokenData.access_token);
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

  // Start polling in background. Clear any stale refresh token so a stuck/invalid
  // one from a previous run doesn't race the device code flow.
  clearRefreshToken();
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
  // 1. In-memory access token still valid?
  if (cachedAccessToken && cachedAccessToken.expiresAt > Date.now() + 300000) {
    return cachedAccessToken.token;
  }

  // 2. Try refresh token from keyring/file
  const refreshToken = loadRefreshToken();
  if (refreshToken) {
    try {
      return await refreshAccessToken(refreshToken);
    } catch {
      // Refresh failed (expired/revoked) — fall through to device code flow
    }
  }

  // 3. If polling is already running, wait for it
  if (deviceCodePollingPromise) {
    return await deviceCodePollingPromise;
  }

  // 4. No cache, no polling — start device code flow (throws with instructions)
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
