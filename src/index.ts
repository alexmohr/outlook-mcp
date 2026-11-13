#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { login } from "./auth.js";
import { configPath, getConfig, writeConfig, DEFAULTS } from "./config.js";
import { getBusy, getEvents, getMyEmail, listMail, MAIL_FOLDERS, markMailRead, readMail, sendMail } from "./ews.js";
import { findFreeRanges, type WorkingHours } from "./slots.js";

if (process.argv[2] === "init") {
  const argv = process.argv.slice(3);
  const opt = (n: string) => {
    const i = argv.indexOf(`--${n}`);
    return i >= 0 ? argv[i + 1] : undefined;
  };
  try {
    const cfg = {
      clientId: opt("client-id") ?? DEFAULTS.clientId,
      tenantId: opt("tenant"),
      redirectUri: opt("redirect-uri") ?? DEFAULTS.redirectUri,
    };
    if (!cfg.tenantId) throw new Error("Pass --tenant <guid-or-domain>");
    await writeConfig(cfg, argv.includes("--force"));
    console.error(`Wrote ${configPath}. Next: npm run login`);
    process.exit(0);
  } catch (e) {
    console.error(`init failed: ${(e as NodeJS.ErrnoException).code === "EEXIST" ? `${configPath} exists (use --force)` : (e as Error).message}`);
    process.exit(1);
  }
}

if (process.argv[2] === "login") {
  try {
    await login();
    process.exit(0);
  } catch (e) {
    console.error(`Login failed: ${(e as Error).message}`);
    process.exit(1);
  }
}

const server = new McpServer({ name: "outlook-mcp", version: "0.1.0" });

const text = (obj: unknown) => ({ content: [{ type: "text" as const, text: JSON.stringify(obj, null, 2) }] });
const fail = (e: unknown) => ({ isError: true, content: [{ type: "text" as const, text: (e as Error).message }] });

/** Format an instant in a time zone as "YYYY-MM-DD HH:mm". */
function fmt(instant: number, timeZone: string): string {
  const p = new Intl.DateTimeFormat("sv-SE", {
    timeZone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hourCycle: "h23",
  }).format(new Date(instant));
  return p;
}

server.registerTool(
  "get_events",
  {
    description: "List events in the user's own Outlook calendar between two ISO datetimes (read-only).",
    inputSchema: {
      from: z.string().describe("ISO 8601 start, e.g. 2026-11-16T00:00:00Z"),
      to: z.string().describe("ISO 8601 end"),
    },
    annotations: { readOnlyHint: true, openWorldHint: true },
  },
  async ({ from, to }) => {
    try {
      return text(await getEvents(Date.parse(from), Date.parse(to)));
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "find_free_slots",
  {
    description:
      "Find free meeting slots in the user's Outlook calendar, optionally intersected with other people's availability " +
      "(free/busy only, via Exchange Web Services; requires that the org shares free/busy). " +
      "Returns free ranges inside working hours that fit the requested duration.",
    inputSchema: {
      durationMinutes: z.number().int().min(5).max(600).describe("Required meeting length"),
      from: z.string().describe("ISO 8601 range start (UTC or with offset)"),
      to: z.string().describe("ISO 8601 range end"),
      attendees: z.array(z.string().email()).optional().describe("Other people's email addresses"),
      timeZone: z.string().optional().describe("IANA zone for working hours/output; defaults to this machine's zone"),
      workingHoursStart: z.string().regex(/^\d{2}:\d{2}$/).optional().describe("HH:MM, default 09:00"),
      workingHoursEnd: z.string().regex(/^\d{2}:\d{2}$/).optional().describe("HH:MM, default 17:00"),
      bufferMinutes: z.number().int().min(0).max(120).optional().describe("Gap kept around existing meetings"),
      includeTentative: z.boolean().optional().describe("Treat tentative as busy (default true)"),
      maxResults: z.number().int().min(1).max(100).optional(),
    },
    annotations: { readOnlyHint: true, openWorldHint: true },
  },
  async (a) => {
    try {
      const from = Date.parse(a.from), to = Date.parse(a.to);
      if (!(from < to)) throw new Error("`from` must be before `to`");
      if (to - from > 62 * 86_400_000) throw new Error("Range too large (max 62 days)");

      const timeZone = a.timeZone ?? getConfig().timeZone ?? Intl.DateTimeFormat().resolvedOptions().timeZone;
      new Intl.DateTimeFormat("en-US", { timeZone }); // throws on invalid IANA zone
      const wh: WorkingHours = {
        start: a.workingHoursStart ?? "09:00",
        end: a.workingHoursEnd ?? "17:00",
        days: [1, 2, 3, 4, 5],
        timeZone,
      };

      const me = await getMyEmail();
      const emails = [me, ...(a.attendees ?? []).filter((e) => e.toLowerCase() !== me.toLowerCase())];
      const people = await getBusy(emails, from, to, a.includeTentative ?? true);

      const ranges = findFreeRanges({
        from, to, durationMin: a.durationMinutes, workingHours: wh,
        busy: people.flatMap((p) => p.busy),
        bufferMin: a.bufferMinutes, stepMin: 15, maxResults: a.maxResults ?? 20,
      });

      return text({
        timeZone,
        workingHours: `${wh.start}-${wh.end}`,
        warnings: people.filter((p) => p.error).map((p) => `${p.email}: ${p.error} (treated as free!)`),
        freeRanges: ranges.map((r) => ({
          start: fmt(r.start, timeZone),
          end: fmt(r.end, timeZone),
          startUtc: new Date(r.start).toISOString(),
          endUtc: new Date(r.end).toISOString(),
          minutes: Math.round((r.end - r.start) / 60_000),
        })),
      });
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "list_mail",
  {
    description:
      "Overview of mail in a folder, newest first: id, sender, subject, received time, read state and a short preview. " +
      "Use unreadOnly for an unread overview. Read-only: does NOT mark anything as read. Use read_mail with an id for the full body.",
    inputSchema: {
      folder: z.enum(MAIL_FOLDERS).optional().describe("Default inbox"),
      unreadOnly: z.boolean().optional().describe("Only unread messages (default false)"),
      since: z.string().optional().describe("ISO 8601; only mail received at or after this instant"),
      limit: z.number().int().min(1).max(100).optional().describe("Default 25"),
      offset: z.number().int().min(0).optional().describe("Skip this many (paging)"),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
  },
  async (a) => {
    try {
      const since = a.since === undefined ? undefined : Date.parse(a.since);
      if (since !== undefined && Number.isNaN(since)) throw new Error("`since` is not a valid ISO 8601 datetime");
      return text(
        await listMail({
          folder: a.folder ?? "inbox",
          unreadOnly: a.unreadOnly ?? false,
          since,
          limit: a.limit ?? 25,
          offset: a.offset ?? 0,
        }),
      );
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "search_mail",
  {
    description:
      "Search mail with Exchange's server-side query syntax (AQS), newest first. Bare words match anywhere; " +
      "keywords narrow it: from:alice, to:bob, subject:report, hasattachments:true, isread:false, " +
      "received:>=2026-11-01, and AND/OR/NOT, e.g. 'from:alice subject:(budget OR forecast) hasattachments:true'. " +
      "Searches all mail folders except Deleted Items and Junk by default. Returns the same fields as list_mail " +
      "(use read_mail with an id for the body). Read-only: does NOT mark anything as read.",
    inputSchema: {
      query: z.string().min(1).max(500).describe("AQS query"),
      folder: z.enum([...MAIL_FOLDERS, "all"]).optional().describe("One folder, or all (default)"),
      unreadOnly: z.boolean().optional().describe("Only unread messages (default false)"),
      since: z.string().optional().describe("ISO 8601; only mail received on or after this date (day granularity)"),
      limit: z.number().int().min(1).max(100).optional().describe("Default 25"),
      offset: z.number().int().min(0).optional().describe("Skip this many (paging)"),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
  },
  async (a) => {
    try {
      const since = a.since === undefined ? undefined : Date.parse(a.since);
      if (since !== undefined && Number.isNaN(since)) throw new Error("`since` is not a valid ISO 8601 datetime");
      return text(
        await listMail({
          query: a.query,
          folder: a.folder ?? "all",
          unreadOnly: a.unreadOnly ?? false,
          since,
          limit: a.limit ?? 25,
          offset: a.offset ?? 0,
        }),
      );
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "read_mail",
  {
    description:
      "Read one message (headers + plain-text body) by id from list_mail. Read-only: does NOT mark it as read; " +
      "call mark_mail_read for that.",
    inputSchema: {
      id: z.string().min(1).describe("Message id from list_mail"),
      maxChars: z.number().int().min(100).max(200_000).optional().describe("Body length cap (default 20000)"),
    },
    annotations: { readOnlyHint: true, openWorldHint: false },
  },
  async ({ id, maxChars }) => {
    try {
      return text(await readMail(id, maxChars ?? 20_000));
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "mark_mail_read",
  {
    description: "Mark messages as read (or unread with isRead=false). Sends no read receipts.",
    inputSchema: {
      ids: z.array(z.string().min(1)).min(1).max(50).describe("Message ids from list_mail"),
      isRead: z.boolean().optional().describe("Default true"),
    },
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  async ({ ids, isRead }) => {
    try {
      return text(await markMailRead(ids, isRead ?? true));
    } catch (e) {
      return fail(e);
    }
  },
);

const mailShape = {
  to: z.array(z.string().email()).min(1).max(50).describe("Recipient email addresses"),
  cc: z.array(z.string().email()).max(50).optional(),
  bcc: z.array(z.string().email()).max(50).optional(),
  subject: z.string().max(255),
  body: z.string().max(200_000).describe("Message body (plain text unless isHtml)"),
  isHtml: z.boolean().optional().describe("Treat body as HTML (default false)"),
};

server.registerTool(
  "draft_mail",
  {
    description: "Save an email to the Drafts folder of the user's Outlook mailbox. Nothing is sent.",
    inputSchema: mailShape,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
  },
  async (a) => {
    try {
      return text(await sendMail({ ...a, saveAsDraft: true }));
    } catch (e) {
      return fail(e);
    }
  },
);

server.registerTool(
  "send_mail",
  {
    description:
      "Send an email from the user's own Outlook mailbox. Sending is irreversible: only call it with recipients, " +
      "subject and body the user has explicitly approved (prefer draft_mail first).",
    inputSchema: mailShape,
    annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
  },
  async (a) => {
    try {
      return text(await sendMail(a));
    } catch (e) {
      return fail(e);
    }
  },
);

await server.connect(new StdioServerTransport());
