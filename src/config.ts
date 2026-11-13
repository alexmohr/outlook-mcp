import { readFileSync } from "node:fs";
import { chmod, mkdir, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";

export interface Config {
  /** Azure AD application (client) ID of a public client allowed in your tenant */
  clientId: string;
  /** Tenant GUID or verified domain */
  tenantId: string;
  /** Redirect URI; only custom-scheme URIs such as urn:ietf:wg:oauth:2.0:oob are supported */
  redirectUri: string;
  /** v1 resource the token is issued for */
  resource: string;
  ewsUrl: string;
  /** Your mailbox address; derived from the token when omitted */
  email?: string;
  /** IANA time zone for working hours; defaults to this machine's zone */
  timeZone?: string;
}

export const configPath = process.env.OUTLOOK_MCP_CONFIG || join(homedir(), ".config", "outlook-mcp", "config.json");

export const DEFAULTS = {
  clientId: "d3590ed6-52b3-4102-aeff-aad2292ab01c", // Microsoft Office public client
  redirectUri: "urn:ietf:wg:oauth:2.0:oob",
  resource: "https://outlook.office365.com",
  ewsUrl: "https://outlook.office365.com/EWS/Exchange.asmx",
};

let cached: Config | undefined;

/** Precedence: OUTLOOK_MCP_* env vars > config file > defaults. */
export function getConfig(): Config {
  if (cached) return cached;
  let file: Partial<Config> = {};
  try {
    file = JSON.parse(readFileSync(configPath, "utf8"));
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== "ENOENT") throw new Error(`Cannot read ${configPath}: ${(e as Error).message}`);
  }
  const env = process.env;
  const tenantId = env.OUTLOOK_MCP_TENANT_ID || file.tenantId;
  if (!tenantId) {
    throw new Error(`No tenantId configured. Run \`node dist/index.js init --tenant <guid-or-domain>\` or edit ${configPath}.`);
  }
  const redirectUri = env.OUTLOOK_MCP_REDIRECT_URI || file.redirectUri || DEFAULTS.redirectUri;
  if (/^https?:/i.test(redirectUri)) {
    throw new Error(`redirectUri ${redirectUri} is not supported; use urn:ietf:wg:oauth:2.0:oob`);
  }
  return (cached = {
    clientId: env.OUTLOOK_MCP_CLIENT_ID || file.clientId || DEFAULTS.clientId,
    tenantId,
    redirectUri,
    resource: env.OUTLOOK_MCP_RESOURCE || file.resource || DEFAULTS.resource,
    ewsUrl: env.OUTLOOK_MCP_EWS_URL || file.ewsUrl || DEFAULTS.ewsUrl,
    email: env.OUTLOOK_MCP_EMAIL || file.email,
    timeZone: env.OUTLOOK_MCP_TIMEZONE || file.timeZone,
  });
}

/** Write a new config file (0600). Refuses to overwrite unless force. */
export async function writeConfig(cfg: Partial<Config>, force: boolean): Promise<void> {
  await mkdir(dirname(configPath), { recursive: true, mode: 0o700 });
  await writeFile(configPath, JSON.stringify(cfg, null, 2) + "\n", { mode: 0o600, flag: force ? "w" : "wx" });
  await chmod(configPath, 0o600);
}
