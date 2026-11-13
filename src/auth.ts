import { execFile } from "node:child_process";
import { chmod, mkdir, readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { getConfig } from "./config.js";

/**
 * Azure AD v1.0 authorization-code flow with the urn:ietf:wg:oauth:2.0:oob redirect, driven through macOS
 * ASWebAuthenticationSession (native helper in helper/main.swift) so Enterprise SSO / conditional access is
 * satisfied. Permissions come from the app registration (v1 is resource-based, not scope-based).
 */
const cachePath = process.env.OUTLOOK_MCP_CACHE || join(homedir(), ".config", "outlook-mcp", "token-cache.json");
const helperPath = join(dirname(fileURLToPath(import.meta.url)), "..", "bin", "auth-helper");
const tokenUrl = () => `https://login.microsoftonline.com/${getConfig().tenantId}/oauth2/token`;

interface Cache {
  refreshToken: string;
  accessToken?: string;
  expiresAt?: number;
}

async function readCache(): Promise<Cache | null> {
  try {
    return JSON.parse(await readFile(cachePath, "utf8"));
  } catch {
    return null;
  }
}

async function writeCache(c: Cache) {
  await mkdir(dirname(cachePath), { recursive: true, mode: 0o700 });
  await writeFile(cachePath, JSON.stringify(c), { mode: 0o600 });
  await chmod(cachePath, 0o600);
}

async function tokenRequest(params: Record<string, string>): Promise<Cache> {
  const { clientId, resource } = getConfig();
  const res = await fetch(tokenUrl(), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: clientId, resource, ...params }),
  });
  const body = (await res.json()) as Record<string, string>;
  if (!res.ok || body.error) {
    throw new Error(`${body.error ?? res.status}: ${(body.error_description ?? "").split("\r\n")[0]}`);
  }
  return {
    refreshToken: body.refresh_token,
    accessToken: body.access_token,
    expiresAt: Date.now() + Number(body.expires_in ?? 3600) * 1000 - 60_000,
  };
}

function runHelper(url: string, callbackScheme: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile(helperPath, [url, callbackScheme], { timeout: 310_000 }, (err, stdout, stderr) => {
      if (err) reject(new Error(stderr.trim() || err.message));
      else resolve(stdout.trim());
    });
  });
}

export async function login(): Promise<void> {
  const { clientId, tenantId, resource, redirectUri } = getConfig();
  const authUrl =
    `https://login.microsoftonline.com/${tenantId}/oauth2/authorize?` +
    new URLSearchParams({ client_id: clientId, response_type: "code", redirect_uri: redirectUri, resource }).toString();
  console.error("Opening macOS authentication session...");
  const callback = await runHelper(authUrl, redirectUri.split(":")[0]);
  const code = /[?&]code=([^&]+)/.exec(callback)?.[1];
  if (!code) {
    const err = /[?&]error_description=([^&]+)/.exec(callback)?.[1];
    throw new Error(`No authorization code in callback${err ? `: ${decodeURIComponent(err.replace(/\+/g, " "))}` : ""}`);
  }
  const cache = await tokenRequest({
    grant_type: "authorization_code",
    code: decodeURIComponent(code),
    redirect_uri: redirectUri,
  });
  await writeCache(cache);
  console.error(`Logged in. Token cache: ${cachePath}`);
}

let inFlight: Promise<string> | null = null;

export function getAccessToken(): Promise<string> {
  return (inFlight ??= (async () => {
    try {
      const c = await readCache();
      if (!c) throw new Error("Not logged in. Run `npm run login` once in a terminal.");
      if (c.accessToken && c.expiresAt && c.expiresAt > Date.now()) return c.accessToken;
      try {
        const fresh = await tokenRequest({ grant_type: "refresh_token", refresh_token: c.refreshToken });
        await writeCache({ ...fresh, refreshToken: fresh.refreshToken ?? c.refreshToken });
        return fresh.accessToken!;
      } catch (e) {
        throw new Error(`Token refresh failed (${(e as Error).message}). Run \`npm run login\` again.`);
      }
    } finally {
      inFlight = null;
    }
  })());
}
