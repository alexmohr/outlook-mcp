/** Pure scheduling logic: busy intervals -> free slots. All instants are epoch ms (UTC). */

export interface Interval {
  start: number;
  end: number;
}

export interface WorkingHours {
  /** "HH:MM" local time in `timeZone` */
  start: string;
  end: string;
  /** 0 = Sunday ... 6 = Saturday */
  days: number[];
  timeZone: string;
}

/** Offset (ms) of `timeZone` from UTC at the given instant. */
function tzOffsetMs(instant: number, timeZone: string): number {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).formatToParts(new Date(instant));
  const g = (t: string) => Number(parts.find((p) => p.type === t)!.value);
  const asUtc = Date.UTC(g("year"), g("month") - 1, g("day"), g("hour"), g("minute"), g("second"));
  return asUtc - Math.floor(instant / 1000) * 1000;
}

/** Convert a wall-clock time in `timeZone` to a UTC instant (DST-safe via two-pass). */
export function zonedToUtc(y: number, m: number, d: number, hh: number, mm: number, timeZone: string): number {
  const guess = Date.UTC(y, m - 1, d, hh, mm);
  const first = guess - tzOffsetMs(guess, timeZone);
  return guess - tzOffsetMs(first, timeZone);
}

/** Local calendar date (y, m, d) and weekday of an instant in `timeZone`. */
function localDate(instant: number, timeZone: string) {
  const shifted = new Date(instant + tzOffsetMs(instant, timeZone));
  return { y: shifted.getUTCFullYear(), m: shifted.getUTCMonth() + 1, d: shifted.getUTCDate() };
}

/** Working-hour windows (UTC instants) overlapping [from, to). */
export function workingWindows(from: number, to: number, wh: WorkingHours): Interval[] {
  const [sh, sm] = wh.start.split(":").map(Number);
  const [eh, em] = wh.end.split(":").map(Number);
  const out: Interval[] = [];
  const first = localDate(from, wh.timeZone);
  // Iterate calendar days by noon-UTC anchor to avoid DST edge issues.
  let anchor = Date.UTC(first.y, first.m - 1, first.d, 12);
  for (let i = 0; i < 400; i++, anchor += 86_400_000) {
    const day = new Date(anchor);
    const y = day.getUTCFullYear(), m = day.getUTCMonth() + 1, d = day.getUTCDate();
    const s = zonedToUtc(y, m, d, sh, sm, wh.timeZone);
    if (s >= to) break;
    if (!wh.days.includes(day.getUTCDay())) continue;
    const e = zonedToUtc(y, m, d, eh, em, wh.timeZone);
    const start = Math.max(s, from), end = Math.min(e, to);
    if (end > start) out.push({ start, end });
  }
  return out;
}

export function mergeIntervals(list: Interval[]): Interval[] {
  const sorted = [...list].filter((i) => i.end > i.start).sort((a, b) => a.start - b.start);
  const out: Interval[] = [];
  for (const i of sorted) {
    const last = out[out.length - 1];
    if (last && i.start <= last.end) last.end = Math.max(last.end, i.end);
    else out.push({ ...i });
  }
  return out;
}

export interface FindOptions {
  from: number;
  to: number;
  busy: Interval[];
  durationMin: number;
  workingHours: WorkingHours;
  /** buffer added around each busy block, minutes */
  bufferMin?: number;
  /** slot start times are rounded up to this grid, minutes */
  stepMin?: number;
  /** earliest allowed start (e.g. now + notice) */
  notBefore?: number;
  maxResults?: number;
}

/** Free slots of at least `durationMin`, returned as maximal free ranges clipped to working hours. */
export function findFreeRanges(o: FindOptions): Interval[] {
  const buf = (o.bufferMin ?? 0) * 60_000;
  const step = (o.stepMin ?? 15) * 60_000;
  const need = o.durationMin * 60_000;
  const busy = mergeIntervals(o.busy.map((b) => ({ start: b.start - buf, end: b.end + buf })));
  const result: Interval[] = [];

  for (const w of workingWindows(Math.max(o.from, o.notBefore ?? o.from), o.to, o.workingHours)) {
    let cursor = w.start;
    const gaps: Interval[] = [];
    for (const b of busy) {
      if (b.end <= cursor) continue;
      if (b.start >= w.end) break;
      if (b.start > cursor) gaps.push({ start: cursor, end: b.start });
      cursor = Math.max(cursor, b.end);
    }
    if (cursor < w.end) gaps.push({ start: cursor, end: w.end });

    for (const g of gaps) {
      const start = Math.ceil(g.start / step) * step;
      if (g.end - start >= need) result.push({ start, end: g.end });
    }
    if (o.maxResults && result.length >= o.maxResults) break;
  }
  return o.maxResults ? result.slice(0, o.maxResults) : result;
}
