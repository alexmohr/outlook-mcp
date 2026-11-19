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
export async function getEvents(from: number, to: number): Promise<{ events: EventSummary[]; truncated: boolean }> {
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
  const events = blocks(xml, "CalendarItem")
    .map((it) => ({
      subject: unesc(tag(it, "Subject") ?? "(No subject)"),
      start: new Date(utc(tag(it, "Start")!)).toISOString(),
      end: new Date(utc(tag(it, "End")!)).toISOString(),
      isAllDay: tag(it, "IsAllDayEvent") === "true",
      location: unesc(tag(it, "Location") ?? "") || undefined,
      organizer: unesc(tag(tag(it, "Organizer") ?? "", "Name") ?? "") || undefined,
    }))
    .sort((a, b) => a.start.localeCompare(b.start));
  const total = Number(attr(xml, "RootFolder", "TotalItemsInView") ?? events.length);
  return { events, truncated: events.length < total };
}

export const SHOW_AS = ["Busy", "Tentative", "Free", "OOF"] as const;
export type ShowAs = (typeof SHOW_AS)[number];

export const RECURRENCE_PATTERNS = ["daily", "weekly", "monthly", "yearly"] as const;
export const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] as const;

export interface EventRecurrence {
  pattern: (typeof RECURRENCE_PATTERNS)[number];
  /** Repeat every `interval` days/weeks/months/years. */
  interval?: number;
  /** Weekly only; defaults to the weekday of `start`. */
  daysOfWeek?: string[];
  /** Monthly/yearly only; defaults to the day of month of `start` (yearly then keeps the start month). */
  dayOfMonth?: number;
  /** Stop after this many occurrences (alternative to `ends`). */
  occurrences?: number;
  /** Stop repeating after this instant (alternative to `occurrences`). */
  ends?: number;
}

export interface MeetingInput {
  subject: string;
  start: number;
  end: number;
  location?: string;
  body?: string;
  isHtml?: boolean;
  attendees?: string[];
  optionalAttendees?: string[];
  reminderMinutes?: number;
  showAs?: ShowAs;
  recurrence?: EventRecurrence;
  /** Save to the calendar (appointment, or draft request with attendees) instead of sending invitations. */
  saveAsDraft: boolean;
}

/** Calendar timezone used to derive recurrence day-of-week/day-of-month/month from the start time. */
const eventTimeZone = () => getConfig().timeZone ?? Intl.DateTimeFormat().resolvedOptions().timeZone;

/** Date (and parts) of an instant in the calendar's display time zone. */
function eventDateParts(t: number) {
  const opts: Intl.DateTimeFormatOptions = { timeZone: eventTimeZone(), weekday: "long", day: "numeric", year: "numeric" };
  const name = new Intl.DateTimeFormat("en-US", { ...opts, month: "long" }).formatToParts(new Date(t));
  const num = new Intl.DateTimeFormat("en-US", { ...opts, month: "numeric" }).formatToParts(new Date(t));
  const g = (parts: Intl.DateTimeFormatPart[], type: string) => parts.find((x) => x.type === type)!.value;
  const pad = (s: string) => s.padStart(2, "0");
  return {
    weekday: g(name, "weekday"),
    monthName: g(name, "month"),
    dayOfMonth: Number(g(num, "day")),
    isoDate: `${g(num, "year")}-${pad(g(num, "month"))}-${pad(g(num, "day"))}`,
  };
}

/**
 * Exchange Online ignores `RecurrenceEnd`, so a date-bounded series is converted to an occurrence count by
 * simulating the pattern day by day (absolute monthly/yearly patterns skip months that lack the day, like Exchange).
 */
function countOccurrences(parts: ReturnType<typeof eventDateParts>, endIsoDate: string, r: EventRecurrence): number {
  const start = Date.parse(`${parts.isoDate}T00:00:00Z`);
  const end = Date.parse(`${endIsoDate}T00:00:00Z`);
  if (end - start > 20 * 365 * 86_400_000) throw new Error("Recurrence `ends` must be within 20 years of `start`");
  const interval = r.interval ?? 1;
  const dayOfMonth = r.dayOfMonth ?? parts.dayOfMonth;
  const days = new Set(r.pattern === "weekly" ? (r.daysOfWeek ?? [parts.weekday]) : []);
  const sd = new Date(start);
  let n = 0;
  for (let t = start; t <= end; t += 86_400_000) {
    const d = new Date(t);
    const dayDiff = Math.round((t - start) / 86_400_000);
    const monthDiff = (d.getUTCFullYear() - sd.getUTCFullYear()) * 12 + (d.getUTCMonth() - sd.getUTCMonth());
    if (r.pattern === "daily" && dayDiff % interval === 0) n++;
    else if (r.pattern === "weekly" && days.has(WEEKDAYS[d.getUTCDay()]) && Math.floor(dayDiff / 7) % interval === 0) n++;
    else if (r.pattern === "monthly" && d.getUTCDate() === dayOfMonth && monthDiff % interval === 0) n++;
    else if (r.pattern === "yearly" && d.getUTCDate() === dayOfMonth && monthDiff % (interval * 12) === 0) n++;
  }
  if (n < 1) throw new Error("Recurrence `ends` before the first occurrence");
  return n;
}

function recurrenceXml(start: number, r: EventRecurrence): string {
  if (r.daysOfWeek && r.pattern !== "weekly") throw new Error("`daysOfWeek` only applies to weekly recurrence");
  if (r.dayOfMonth && r.pattern !== "monthly" && r.pattern !== "yearly") throw new Error("`dayOfMonth` only applies to monthly/yearly recurrence");
  if (r.occurrences !== undefined && r.ends !== undefined) throw new Error("recurrence: use either `occurrences` or `ends`, not both");
  const interval = r.interval ?? 1;
  const parts = eventDateParts(start);
  const occurrences = r.occurrences ?? (r.ends !== undefined ? countOccurrences(parts, eventDateParts(r.ends).isoDate, r) : undefined);
  const end = occurrences !== undefined
    ? `<t:NumberedRecurrence><t:StartDate>${parts.isoDate}</t:StartDate><t:NumberOfOccurrences>${occurrences}</t:NumberOfOccurrences></t:NumberedRecurrence>`
    : "";
  switch (r.pattern) {
    case "daily":
      return `<t:Recurrence><t:DailyRecurrence><t:Interval>${interval}</t:Interval></t:DailyRecurrence>${end}</t:Recurrence>`;
    case "weekly": {
      const days = (r.daysOfWeek ?? [parts.weekday]).filter((d) => (WEEKDAYS as readonly string[]).includes(d)).join(" ");
      if (!days) throw new Error("`daysOfWeek` must use the English weekday names");
      return `<t:Recurrence><t:WeeklyRecurrence><t:Interval>${interval}</t:Interval><t:DaysOfWeek>${days}</t:DaysOfWeek></t:WeeklyRecurrence>${end}</t:Recurrence>`;
    }
    case "monthly":
      return `<t:Recurrence><t:AbsoluteMonthlyRecurrence><t:Interval>${interval}</t:Interval><t:DayOfMonth>${r.dayOfMonth ?? parts.dayOfMonth}</t:DayOfMonth></t:AbsoluteMonthlyRecurrence>${end}</t:Recurrence>`;
    case "yearly":
      return `<t:Recurrence><t:AbsoluteYearlyRecurrence><t:Interval>${interval}</t:Interval><t:DayOfMonth>${r.dayOfMonth ?? parts.dayOfMonth}</t:DayOfMonth><t:Month>${parts.monthName}</t:Month></t:AbsoluteYearlyRecurrence>${end}</t:Recurrence>`;
  }
}

/**
 * Create a calendar event via EWS CreateItem: either save it as an appointment / draft meeting request
 * (saveAsDraft, nothing is sent) or send invitations to the attendees.
 */
export async function createMeeting(m: MeetingInput): Promise<{
  status: "invitation_sent" | "draft_saved";
  id?: string;
  subject: string;
  start: string;
  end: string;
  attendees: string[];
  optionalAttendees: string[];
}> {
  if (m.recurrence?.ends !== undefined && m.recurrence.ends <= m.start) throw new Error("Recurrence must end after the event starts");
  const rcpts = (name: string, list?: string[]) =>
    list?.length
      ? `<t:${name}>${list.map((e) => `<t:Attendee><t:Mailbox><t:EmailAddress>${esc(e)}</t:EmailAddress></t:Mailbox></t:Attendee>`).join("")}</t:${name}>`
      : "";
  const xml = await ews(
    (m.saveAsDraft
      ? `<m:CreateItem MessageDisposition="SaveOnly" SendMeetingInvitations="SendNone">`
      : `<m:CreateItem SendMeetingInvitations="SendToAllAndSaveCopy">`) +
      `<m:SavedItemFolderId><t:DistinguishedFolderId Id="calendar"/></m:SavedItemFolderId>` +
      `<m:Items><t:CalendarItem>` +
      `<t:Subject>${esc(m.subject)}</t:Subject>` +
      (m.body !== undefined ? `<t:Body BodyType="${m.isHtml ? "HTML" : "Text"}">${esc(m.body)}</t:Body>` : "") +
      (m.reminderMinutes !== undefined
        ? `<t:ReminderMinutesBeforeStart>${m.reminderMinutes}</t:ReminderMinutesBeforeStart><t:IsReminderOn>true</t:IsReminderOn>`
        : "") +
      `<t:Start>${isoNoMs(m.start)}</t:Start><t:End>${isoNoMs(m.end)}</t:End>` +
      (m.location !== undefined ? `<t:Location>${esc(m.location)}</t:Location>` : "") +
      rcpts("RequiredAttendees", m.attendees) +
      rcpts("OptionalAttendees", m.optionalAttendees) +
      (m.showAs ? `<t:ShowAs>${m.showAs}</t:ShowAs>` : "") +
      (m.recurrence ? recurrenceXml(m.start, m.recurrence) : "") +
      `</t:CalendarItem></m:Items></m:CreateItem>`,
  );
  if (!/ResponseClass="Success"/.test(xml) || /ResponseClass="Error"/.test(xml)) throw ewsError(xml);
  return {
    status: m.saveAsDraft ? "draft_saved" : "invitation_sent",
    id: attr(xml, "ItemId", "Id"),
    subject: m.subject,
    start: new Date(m.start).toISOString(),
    end: new Date(m.end).toISOString(),
    attendees: m.attendees ?? [],
    optionalAttendees: m.optionalAttendees ?? [],
  };
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
/** Well-known folders addressable by name; "msgfolderroot" is the top of the mail folder tree. */
const WELL_KNOWN = new Set<string>([...MAIL_FOLDERS, "msgfolderroot"]);

/** A well-known folder name (case-insensitive) or a folder id from listFolders, as an EWS folder id element. */
const folderRef = (f: string) =>
  WELL_KNOWN.has(f.toLowerCase()) ? `<t:DistinguishedFolderId Id="${f.toLowerCase()}"/>` : `<t:FolderId Id="${esc(f)}"/>`;

const ewsError = (xml: string) => new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);

const mailbox = (xml: string | undefined) =>
  xml === undefined
    ? undefined
    : { name: unesc(tag(xml, "Name") ?? "") || undefined, email: unesc(tag(xml, "EmailAddress") ?? "") || undefined };
const mailboxes = (xml: string | undefined) => blocks(xml ?? "", "Mailbox").map((m) => mailbox(m)!);

/** Element names of mail items: meeting requests/responses/cancellations are mail too but have their own types. */
const MAIL_ITEM = "(?:Message|MeetingRequest|MeetingResponse|MeetingCancellation)";

const flagStatus = (it: string) => {
  const f = tag(tag(it, "Flag") ?? "", "FlagStatus");
  return f && f !== "NotFlagged" ? f : undefined;
};

export interface MailSummary {
  id: string;
  subject: string;
  from?: { name?: string; email?: string };
  received: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance?: string;
  preview?: string;
  /** Follow-up flag: "Flagged" or "Complete"; omitted when not flagged. */
  flag?: string;
}


/** Like Promise.all(items.map(fn)) but with at most `limit` calls in flight (EWS throttles concurrent connections). */
async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) {
        const i = next++;
        out[i] = await fn(items[i]);
      }
    }),
  );
  return out;
}

/** Ids of all mail folders (recursive), excluding Deleted Items and Junk Email. */
async function searchableMailFolderIds(): Promise<string[]> {
  const [all, skip] = await Promise.all([
    ews(
      `<m:FindFolder Traversal="Deep"><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape>` +
        `<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI="folder:FolderClass"/><t:FieldURIOrConstant><t:Constant Value="IPF.Note"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>` +
        `<m:ParentFolderIds><t:DistinguishedFolderId Id="msgfolderroot"/></m:ParentFolderIds></m:FindFolder>`,
    ),
    ews(
      `<m:GetFolder><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape></m:FolderShape><m:FolderIds>` +
        `<t:DistinguishedFolderId Id="deleteditems"/><t:DistinguishedFolderId Id="junkemail"/></m:FolderIds></m:GetFolder>`,
    ),
  ]);
  if (/ResponseClass="Error"/.test(all)) throw new Error(`EWS error: ${unesc(tag(all, "MessageText") ?? "unknown")}`);
  const excluded = new Set([...skip.matchAll(/<(?:\w+:)?FolderId\s[^>]*\bId="([^"]*)"/g)].map((m) => m[1]));
  return [...all.matchAll(/<(?:\w+:)?FolderId\s[^>]*\bId="([^"]*)"/g)].map((m) => m[1]).filter((id) => !excluded.has(id));
}

/**
 * Overview of one folder (or, with folder "all", every mail folder except Deleted/Junk) via FindItem, newest first.
 * An optional Advanced Query Syntax `query` is evaluated server-side. Does NOT mark anything as read.
 */
export async function listMail(o: {
  /** Well-known folder name, folder id, or "all" (every mail folder except Deleted/Junk). */
  folder: string;
  unreadOnly: boolean;
  since?: number;
  limit: number;
  offset: number;
  /** Advanced Query Syntax (e.g. `from:alice subject:report`). */
  query?: string;
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
  // EWS does not allow Restriction and QueryString together, so with a query the filters go into the query itself.
  const aqs = o.query
    ? [`(${o.query})`, ...(o.unreadOnly ? ["isread:false"] : []), ...(o.since !== undefined ? [`received:>=${isoNoMs(o.since).slice(0, 10)}`] : [])].join(" ")
    : undefined;
  const restriction = aqs || !conds.length
    ? ""
    : `<m:Restriction>${conds.length > 1 ? `<t:And>${conds.join("")}</t:And>` : conds[0]}</m:Restriction>`;

  const findIn = async (parent: string, max: number, offset: number) => {
    const xml = await ews(
      `<m:FindItem Traversal="Shallow"><m:ItemShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>` +
        ["item:Subject", "message:From", "item:DateTimeReceived", "message:IsRead", "item:HasAttachments", "item:Importance", "item:Preview", "item:Flag"]
          .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
          .join("") +
        `</t:AdditionalProperties></m:ItemShape>` +
        `<m:IndexedPageItemView MaxEntriesReturned="${max}" Offset="${offset}" BasePoint="Beginning"/>` +
        restriction +
        `<m:SortOrder><t:FieldOrder Order="Descending"><t:FieldURI FieldURI="item:DateTimeReceived"/></t:FieldOrder></m:SortOrder>` +
        `<m:ParentFolderIds>${parent}</m:ParentFolderIds>` +
        (aqs ? `<m:QueryString>${esc(aqs)}</m:QueryString>` : "") +
        `</m:FindItem>`,
    );
    if (/ResponseClass="Error"/.test(xml)) throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
    const mails: MailSummary[] = blocks(xml, MAIL_ITEM).map((it) => ({
      id: attr(it, "ItemId", "Id")!,
      subject: unesc(tag(it, "Subject") ?? "(No subject)"),
      from: mailbox(tag(tag(it, "From") ?? "", "Mailbox")),
      received: new Date(utc(tag(it, "DateTimeReceived")!)).toISOString(),
      isRead: tag(it, "IsRead") === "true",
      hasAttachments: tag(it, "HasAttachments") === "true",
      importance: tag(it, "Importance"),
      preview: unesc(tag(it, "Preview") ?? "") || undefined,
      flag: flagStatus(it),
    }));
    return { mails, total: Number(attr(xml, "RootFolder", "TotalItemsInView") ?? mails.length) };
  };

  if (o.folder !== "all") {
    const r = await findIn(folderRef(o.folder), o.limit, o.offset);
    return { totalMatching: r.total, returned: r.mails.length, hasMore: o.offset + r.mails.length < r.total, mails: r.mails };
  }

  // All folders: ask each for its newest (offset + limit) matches, merge by date, then page.
  const ids = await searchableMailFolderIds();
  const parts = await mapLimit(ids, 3, (id) => findIn(`<t:FolderId Id="${esc(id)}"/>`, o.offset + o.limit, 0));
  const total = parts.reduce((n, p) => n + p.total, 0);
  const mails = parts
    .flatMap((p) => p.mails)
    .sort((a, b) => b.received.localeCompare(a.received))
    .slice(o.offset, o.offset + o.limit);
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
       "item:HasAttachments", "item:Importance", "item:Flag", "item:Body"]
        .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
        .join("") +
      `</t:AdditionalProperties></m:ItemShape><m:ItemIds><t:ItemId Id="${esc(id)}"/></m:ItemIds></m:GetItem>`,
  );
  if (/ResponseClass="Error"/.test(xml)) throw new Error(`EWS error: ${unesc(tag(xml, "MessageText") ?? "unknown")}`);
  const it = blocks(xml, MAIL_ITEM)[0];
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
    flag: flagStatus(it),
    body: body.slice(0, maxChars),
    bodyTruncated: body.length > maxChars,
  };
}

type ItemResult = { id: string; ok: boolean; error?: string };

/** Apply the same field update to several items via UpdateItem, reporting success per item. */
async function updateItems(ids: string[], update: string): Promise<ItemResult[]> {
  const xml = await ews(
    `<m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AlwaysOverwrite"><m:ItemChanges>` +
      ids.map((id) => `<t:ItemChange><t:ItemId Id="${esc(id)}"/><t:Updates>${update}</t:Updates></t:ItemChange>`).join("") +
      `</m:ItemChanges></m:UpdateItem>`,
  );
  const rs = [...xml.matchAll(/<(?:\w+:)?UpdateItemResponseMessage\s+ResponseClass="(\w+)"[^>]*>([\s\S]*?)<\/(?:\w+:)?UpdateItemResponseMessage>/g)];
  if (!rs.length && /ResponseClass="Error"/.test(xml)) throw ewsError(xml);
  return ids.map((id, i) => {
    const r = rs[i];
    if (r?.[1] === "Success") return { id, ok: true };
    return { id, ok: false, error: unesc(tag(r?.[2] ?? "", "MessageText") ?? "no response") };
  });
}

/** Set the read flag on one or more messages via UpdateItem (no read receipts are sent). */
export async function markMailRead(ids: string[], isRead: boolean): Promise<ItemResult[]> {
  return updateItems(
    ids,
    `<t:SetItemField><t:FieldURI FieldURI="message:IsRead"/><t:Message><t:IsRead>${isRead}</t:IsRead></t:Message></t:SetItemField>`,
  );
}

export const FLAG_STATES = ["flagged", "complete", "cleared"] as const;
export type FlagState = (typeof FLAG_STATES)[number];

/**
 * Set the follow-up flag on messages: flagged (optionally with start and due date), complete (marked done now),
 * or cleared.
 */
export async function flagMail(ids: string[], state: FlagState, dates?: { start: number; due: number }): Promise<ItemResult[]> {
  const inner =
    state === "cleared"
      ? `<t:FlagStatus>NotFlagged</t:FlagStatus>`
      : state === "complete"
        ? `<t:FlagStatus>Complete</t:FlagStatus><t:CompleteDate>${isoNoMs(Date.now())}</t:CompleteDate>`
        : `<t:FlagStatus>Flagged</t:FlagStatus>` +
          (dates ? `<t:StartDate>${isoNoMs(dates.start)}</t:StartDate><t:DueDate>${isoNoMs(dates.due)}</t:DueDate>` : "");
  return updateItems(ids, `<t:SetItemField><t:FieldURI FieldURI="item:Flag"/><t:Message><t:Flag>${inner}</t:Flag></t:Message></t:SetItemField>`);
}

export interface FolderInfo {
  id: string;
  name: string;
  /** Slash-separated display path from the top of the mailbox, e.g. "Inbox/Projects". */
  path: string;
  parentId?: string;
  totalCount: number;
  unreadCount: number;
}

/** All mail folders (recursive) with their paths and counts, via FindFolder. */
export async function listFolders(): Promise<FolderInfo[]> {
  const xml = await ews(
    `<m:FindFolder Traversal="Deep"><m:FolderShape><t:BaseShape>IdOnly</t:BaseShape><t:AdditionalProperties>` +
      ["folder:DisplayName", "folder:ParentFolderId", "folder:TotalCount", "folder:UnreadCount"]
        .map((f) => `<t:FieldURI FieldURI="${f}"/>`)
        .join("") +
      `</t:AdditionalProperties></m:FolderShape>` +
      `<m:Restriction><t:IsEqualTo><t:FieldURI FieldURI="folder:FolderClass"/><t:FieldURIOrConstant><t:Constant Value="IPF.Note"/></t:FieldURIOrConstant></t:IsEqualTo></m:Restriction>` +
      `<m:ParentFolderIds><t:DistinguishedFolderId Id="msgfolderroot"/></m:ParentFolderIds></m:FindFolder>`,
  );
  if (/ResponseClass="Error"/.test(xml)) throw ewsError(xml);
  const raw = blocks(xml, "Folder").map((f) => ({
    id: attr(f, "FolderId", "Id")!,
    name: unesc(tag(f, "DisplayName") ?? ""),
    parentId: attr(f, "ParentFolderId", "Id"),
    totalCount: Number(tag(f, "TotalCount") ?? 0),
    unreadCount: Number(tag(f, "UnreadCount") ?? 0),
  }));
  const byId = new Map(raw.map((f) => [f.id, f]));
  const path = (f: (typeof raw)[number], depth = 0): string => {
    const parent = f.parentId ? byId.get(f.parentId) : undefined;
    return parent && depth < 50 ? `${path(parent, depth + 1)}/${f.name}` : f.name;
  };
  return raw
    .map((f) => ({ ...f, path: path(f), parentId: f.parentId && byId.has(f.parentId) ? f.parentId : undefined }))
    .sort((a, b) => a.path.localeCompare(b.path));
}

/** Create a mail folder under `parent` (well-known name or folder id; "msgfolderroot" = top level) via CreateFolder. */
export async function createFolder(name: string, parent: string): Promise<{ id: string; name: string }> {
  const xml = await ews(
    `<m:CreateFolder><m:ParentFolderId>${folderRef(parent)}</m:ParentFolderId><m:Folders><t:Folder>` +
      `<t:FolderClass>IPF.Note</t:FolderClass><t:DisplayName>${esc(name)}</t:DisplayName>` +
      `</t:Folder></m:Folders></m:CreateFolder>`,
  );
  const id = attr(xml, "FolderId", "Id");
  if (/ResponseClass="Error"/.test(xml) || !id) throw ewsError(xml);
  return { id, name };
}

/** Move messages to a folder (well-known name or folder id) via MoveItem. Moved messages get new ids. */
export async function moveMail(ids: string[], destination: string): Promise<{ id: string; ok: boolean; newId?: string; error?: string }[]> {
  const xml = await ews(
    `<m:MoveItem><m:ToFolderId>${folderRef(destination)}</m:ToFolderId><m:ItemIds>` +
      ids.map((id) => `<t:ItemId Id="${esc(id)}"/>`).join("") +
      `</m:ItemIds></m:MoveItem>`,
  );
  const rs = [...xml.matchAll(/<(?:\w+:)?MoveItemResponseMessage\s+ResponseClass="(\w+)"[^>]*>([\s\S]*?)<\/(?:\w+:)?MoveItemResponseMessage>/g)];
  if (!rs.length && /ResponseClass="Error"/.test(xml)) throw ewsError(xml);
  return ids.map((id, i) => {
    const r = rs[i];
    if (r?.[1] === "Success") return { id, ok: true, newId: attr(r[2], "ItemId", "Id") };
    return { id, ok: false, error: unesc(tag(r?.[2] ?? "", "MessageText") ?? "no response") };
  });
}
