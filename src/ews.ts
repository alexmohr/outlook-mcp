import { getAccessToken } from "./auth.js";
import { getConfig } from "./config.js";
import type { Interval } from "./slots.js";

const esc = (s: string) => s.replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" })[c]!);
const utc = (s: string) => Date.parse(/(Z|[+-]\d\d:\d\d)$/.test(s) ? s : `${s}Z`);
const isoNoMs = (t: number) => new Date(t).toISOString().replace(/\.\d{3}Z$/, "Z");
const tag = (xml: string, name: string) =>
  new RegExp(`<(?:\\w+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)</(?:\\w+:)?${name}>`).exec(xml)?.[1];
const blocks = (xml: string, name: string) =>
  [...xml.matchAll(new RegExp(`<(?:\\w+:)?${name}(?:\\s[^>]*)?>([\\s\\S]*?)</(?:\\w+:)?${name}>`, "g"))].map((m) => m[1]);
const unesc = (s: string) => s.replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d))).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");

const NS =
  'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
  'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
  'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
  'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"';

async function ews(body: string): Promise<string> {
  const token = await getAccessToken();
  const res = await fetch(getConfig().ewsUrl, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "text/xml; charset=utf-8" },
    body:
      `<?xml version="1.0" encoding="utf-8"?><soap:Envelope ${NS}><soap:Header>` +
      `<t:RequestServerVersion Version="Exchange2013_SP1"/></soap:Header><soap:Body>${body}</soap:Body></soap:Envelope>`,
  });
  const xml = await res.text();
  if (!res.ok) throw new Error(`EWS ${res.status}: ${xml.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").slice(0, 300)}`);
  return xml;
}

/** Own address from the access token claims (EWS has no /me). Override with OUTLOOK_MCP_EMAIL. */
export async function getMyEmail(): Promise<string> {
  if (getConfig().email) return getConfig().email!;
  const claims = JSON.parse(Buffer.from((await getAccessToken()).split(".")[1], "base64url").toString());
  const email = claims.upn || claims.unique_name || claims.preferred_username || claims.email;
  if (!email) throw new Error("Could not determine your email from the token; set email in the config file.");
  return email;
}

export interface PersonBusy {
  email: string;
  busy: Interval[];
  error?: string;
}

/** Free/busy for several mailboxes via EWS GetUserAvailability (UTC). */
export async function getBusy(emails: string[], from: number, to: number, includeTentative: boolean): Promise<PersonBusy[]> {
  const zero = (m: number) =>
    `<t:Bias>0</t:Bias><t:Time>00:00:00</t:Time><t:DayOrder>1</t:DayOrder><t:Month>${m}</t:Month><t:DayOfWeek>Sunday</t:DayOfWeek>`;
  const xml = await ews(
    `<m:GetUserAvailabilityRequest>` +
      `<t:TimeZone><t:Bias>0</t:Bias><t:StandardTime>${zero(1)}</t:StandardTime><t:DaylightTime>${zero(7)}</t:DaylightTime></t:TimeZone>` +
      `<m:MailboxDataArray>` +
      emails
        .map((e) => `<t:MailboxData><t:Email><t:Address>${esc(e)}</t:Address></t:Email><t:AttendeeType>Required</t:AttendeeType><t:ExcludeConflicts>false</t:ExcludeConflicts></t:MailboxData>`)
        .join("") +
      `</m:MailboxDataArray>` +
      `<t:FreeBusyViewOptions><t:TimeWindow><t:StartTime>${isoNoMs(from)}</t:StartTime><t:EndTime>${isoNoMs(to)}</t:EndTime></t:TimeWindow>` +
      `<t:MergedFreeBusyIntervalInMinutes>15</t:MergedFreeBusyIntervalInMinutes><t:RequestedView>FreeBusy</t:RequestedView></t:FreeBusyViewOptions>` +
      `</m:GetUserAvailabilityRequest>`,
  );

  const blocking = new Set(["Busy", "OOF", ...(includeTentative ? ["Tentative"] : [])]);
  const responses = blocks(xml, "FreeBusyResponse");
  return emails.map((email, i) => {
    const r = responses[i] ?? "";
    if (/ResponseClass="Error"/.test(r) || !r) {
      return { email, busy: [], error: unesc(tag(r, "MessageText") ?? "no response") };
    }
    const busy = blocks(r, "CalendarEvent")
      .filter((ev) => blocking.has(tag(ev, "BusyType") ?? ""))
      .map((ev) => ({ start: utc(tag(ev, "StartTime")!), end: utc(tag(ev, "EndTime")!) }));
    return { email, busy };
  });
}

export interface EventSummary {
  subject: string;
  start: string;
  end: string;
  isAllDay: boolean;
  location?: string;
  organizer?: string;
}

/** Own calendar events (read-only) via FindItem + CalendarView (expands recurrences). */
export async function getEvents(from: number, to: number): Promise<EventSummary[]> {
  const xml = await ews(
    `<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>Default</t:BaseShape><t:AdditionalProperties>` +
      ["item:Subject", "calendar:Start", "calendar:End", "calendar:IsAllDayEvent", "calendar:Location", "calendar:Organizer"]
        .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
        .join("") +
      `</t:AdditionalProperties></m:ItemShape>` +
      `<m:CalendarView MaxEntriesReturned="200" StartDate="${isoNoMs(from)}" EndDate="${isoNoMs(to)}"/>` +
      `<m:ParentFolderIds><t:DistinguishedFolderId Id="calendar"/></m:ParentFolderIds></m:FindItem>`,
  );
  if (/ResponseClass="Error"/.test(xml)) throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
  return blocks(xml, "CalendarItem")
    .map((it) => ({
      subject: unesc(tag(it, "Subject") ?? "(No subject)"),
      start: new Date(utc(tag(it, "Start")!)).toISOString(),
      end: new Date(utc(tag(it, "End")!)).toISOString(),
      isAllDay: tag(it, "IsAllDayEvent") === "true",
      location: unesc(tag(it, "Location") ?? "") || undefined,
      organizer: unesc(tag(tag(it, "Organizer") ?? "", "Name") ?? "") || undefined,
    }))
    .sort((a, b) => a.start.localeCompare(b.start));
}

export interface MailInput {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  isHtml?: boolean;
  /** Save to Drafts instead of sending. */
  saveAsDraft?: boolean;
}

/** Send (or save as draft) a message from the user's own mailbox via EWS CreateItem. */
export async function sendMail(m: MailInput): Promise<{ status: "sent" | "draft_saved"; to: string[]; cc: string[]; bcc: string[]; subject: string }> {
  const rcpts = (name: string, list?: string[]) =>
    list?.length
      ? `<t:${name}>${list.map((e) => `<t:Mailbox><t:EmailAddress>${esc(e)}</t:EmailAddress></t:Mailbox>`).join("")}</t:${name}>`
      : "";
  const draft = !!m.saveAsDraft;
  const xml = await ews(
    `<m:CreateItem MessageDisposition="${draft ? "SaveOnly" : "SendAndSaveCopy"}">` +
      `<m:SavedItemFolderId><t:DistinguishedFolderId Id="${draft ? "drafts" : "sentitems"}"/></m:SavedItemFolderId>` +
      `<m:Items><t:Message><t:Subject>${esc(m.subject)}</t:Subject>` +
      `<t:Body BodyType="${m.isHtml ? "HTML" : "Text"}">${esc(m.body)}</t:Body>` +
      rcpts("ToRecipients", m.to) + rcpts("CcRecipients", m.cc) + rcpts("BccRecipients", m.bcc) +
      `</t:Message></m:Items></m:CreateItem>`,
  );
  if (!/ResponseClass="Success"/.test(xml) || /ResponseClass="Error"/.test(xml)) {
    throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
  }
  return { status: draft ? "draft_saved" : "sent", to: m.to, cc: m.cc ?? [], bcc: m.bcc ?? [], subject: m.subject };
}

const attr = (xml: string, el: string, name: string) =>
  new RegExp(`<(?:\\w+:)?${el}\\s[^>]*\\b${name}="([^"]*)"`).exec(xml)?.[1];

export const MAIL_FOLDERS = ["inbox", "sentitems", "drafts", "junkemail", "deleteditems", "archive"] as const;
export type MailFolder = (typeof MAIL_FOLDERS)[number];

const mailbox = (xml: string | undefined) =>
  xml === undefined
    ? undefined
    : { name: unesc(tag(xml, "Name") ?? "") || undefined, email: unesc(tag(xml, "EmailAddress") ?? "") || undefined };
const mailboxes = (xml: string | undefined) => blocks(xml ?? "", "Mailbox").map((m) => mailbox(m)!);

export interface MailSummary {
  id: string;
  subject: string;
  from?: { name?: string; email?: string };
  received: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance?: string;
  preview?: string;
}

/** Overview of a folder via FindItem. Does NOT mark anything as read. */
export async function listMail(o: {
  folder: MailFolder;
  unreadOnly: boolean;
  since?: number;
  limit: number;
  offset: number;
}): Promise<{ totalMatching: number; returned: number; hasMore: boolean; mails: MailSummary[] }> {
  const conds: string[] = [];
  if (o.unreadOnly) {
    conds.push(`<t:IsEqualTo><t:FieldURI FieldURI="message:IsRead"/><t:FieldURIOrConstant><t:Constant Value="false"/></t:FieldURIOrConstant></t:IsEqualTo>`);
  }
  if (o.since !== undefined) {
    conds.push(
      `<t:IsGreaterThanOrEqualTo><t:FieldURI FieldURI="item:DateTimeReceived"/><t:FieldURIOrConstant><t:Constant Value="${isoNoMs(o.since)}"/></t:FieldURIOrConstant></t:IsGreaterThanOrEqualTo>`,
    );
  }
  const restriction = conds.length
    ? `<m:Restriction>${conds.length > 1 ? `<t:And>${conds.join("")}</t:And>` : conds[0]}</m:Restriction>`
    : "";
  const xml = await ews(
    `<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>` +
      ["item:Subject", "message:From", "item:DateTimeReceived", "message:IsRead", "item:HasAttachments", "item:Importance", "item:Preview"]
        .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
        .join("") +
      `</t:AdditionalProperties></m:ItemShape>` +
      `<m:IndexedPageItemView MaxEntriesReturned="${o.limit}" Offset="${o.offset}" BasePoint="Beginning"/>` +
      restriction +
      `<m:SortOrder><t:FieldOrder Order="Descending"><t:FieldURI FieldURI="item:DateTimeReceived"/></t:FieldOrder></m:SortOrder>` +
      `<m:ParentFolderIds><t:DistinguishedFolderId Id="${o.folder}"/></m:ParentFolderIds></m:FindItem>`,
  );
  if (/ResponseClass="Error"/.test(xml)) throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
  const mails = blocks(xml, "Message").map((it) => ({
    id: attr(it, "ItemId", "Id")!,
    subject: unesc(tag(it, "Subject") ?? "(No subject)"),
    from: mailbox(tag(tag(it, "From") ?? "", "Mailbox")),
    received: new Date(utc(tag(it, "DateTimeReceived")!)).toISOString(),
    isRead: tag(it, "IsRead") === "true",
    hasAttachments: tag(it, "HasAttachments") === "true",
    importance: tag(it, "Importance"),
    preview: unesc(tag(it, "Preview") ?? "") || undefined,
  }));
  const total = Number(attr(xml, "RootFolder", "TotalItemsInView") ?? mails.length);
  return { totalMatching: total, returned: mails.length, hasMore: o.offset + mails.length < total, mails };
}

export interface MailFull extends MailSummary {
  to: { name?: string; email?: string }[];
  cc: { name?: string; email?: string }[];
  body: string;
  bodyTruncated: boolean;
}

/** Full message via GetItem (plain-text body). Does NOT mark the message as read. */
export async function readMail(id: string, maxChars: number): Promise<MailFull> {
  const xml = await ews(
    `<m:GetItem><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:BodyType>Text</t:BodyType><t:AdditionalProperties>` +
      ["item:Subject", "message:From", "message:ToRecipients", "message:CcRecipients", "item:DateTimeReceived", "message:IsRead",
       "item:HasAttachments", "item:Importance", "item:Body"]
        .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
        .join("") +
      `</t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id="${esc(id)}"/></m:ItemIds></m:GetItem>`,
  );
  if (/ResponseClass="Error"/.test(xml)) throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
  const it = blocks(xml, "Message")[0];
  if (!it) throw new Error("Message not found");
  const body = unesc(tag(it, "Body") ?? "").replace(/\r\n?/g, "\n");
  return {
    id,
    subject: unesc(tag(it, "Subject") ?? "(No subject)"),
    from: mailbox(tag(tag(it, "From") ?? "", "Mailbox")),
    to: mailboxes(tag(it, "ToRecipients")),
    cc: mailboxes(tag(it, "CcRecipients")),
    received: new Date(utc(tag(it, "DateTimeReceived")!)).toISOString(),
    isRead: tag(it, "IsRead") === "true",
    hasAttachments: tag(it, "HasAttachments") === "true",
    importance: tag(it, "Importance"),
    body: body.slice(0, maxChars),
    bodyTruncated: body.length > maxChars,
  };
}

/** Set the read flag on one or more messages via UpdateItem (no read receipts are sent). */
export async function markMailRead(ids: string[], isRead: boolean): Promise<{ id: string; ok: boolean; error?: string }[]> {
  const xml = await ews(
    `<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite"><m:ItemChanges>` +
      ids
        .map(
          (id) =>
            `<t:ItemChange><t:ItemId Id="${esc(id)}"/><t:Updates><t:SetItemField><t:FieldURI FieldURI="message:IsRead"/>` +
            `<t:Message><t:IsRead>${isRead}</t:IsRead></t:Message></t:SetItemField></t:Updates></t:ItemChange>`,
        )
        .join("") +
      `</m:ItemChanges></m:UpdateItem>`,
  );
  const rs = [...xml.matchAll(/<(?:\w+:)?UpdateItemResponseMessage\s+ResponseClass="(\w+)"[^>]*>([\s\S]*?)<\/(?:\w+:)?UpdateItemResponseMessage>/g)];
  return ids.map((id, i) => {
    const r = rs[i];
    if (r?.[1] === "Success") return { id, ok: true };
    return { id, ok: false, error: unesc(tag(r?.[2] ?? "", "MessageText") ?? "no response") };
  });
}
