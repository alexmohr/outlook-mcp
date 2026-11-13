import { test } from "node:test";
import assert from "node:assert/strict";
import { findFreeRanges, zonedToUtc, type WorkingHours } from "./slots.js";

const tz = "Europe/Berlin";
const wh: WorkingHours = { start: "09:00", end: "17:00", days: [1, 2, 3, 4, 5], timeZone: tz };

test("zonedToUtc handles CET and CEST", () => {
  assert.equal(new Date(zonedToUtc(2026, 1, 15, 9, 0, tz)).toISOString(), "2026-01-15T08:00:00.000Z");
  assert.equal(new Date(zonedToUtc(2026, 7, 15, 9, 0, tz)).toISOString(), "2026-07-15T07:00:00.000Z");
});

test("free ranges skip busy blocks and weekends", () => {
  // Fri 2026-11-13 .. Mon 2026-11-16 (CET = UTC+1)
  const from = zonedToUtc(2026, 11, 13, 0, 0, tz);
  const to = zonedToUtc(2026, 11, 17, 0, 0, tz);
  const busy = [{ start: zonedToUtc(2026, 11, 13, 10, 0, tz), end: zonedToUtc(2026, 11, 13, 12, 0, tz) }];
  const r = findFreeRanges({ from, to, busy, durationMin: 60, workingHours: wh });
  const fmt = (i: { start: number; end: number }) => `${new Date(i.start).toISOString()}|${new Date(i.end).toISOString()}`;
  assert.deepEqual(r.map(fmt), [
    "2026-11-13T08:00:00.000Z|2026-11-13T09:00:00.000Z", // 09-10
    "2026-11-13T11:00:00.000Z|2026-11-13T16:00:00.000Z", // 12-17
    "2026-11-16T08:00:00.000Z|2026-11-16T16:00:00.000Z", // Monday
  ]);
});

test("duration filter drops short gaps; buffer widens busy", () => {
  const from = zonedToUtc(2026, 11, 16, 0, 0, tz);
  const to = zonedToUtc(2026, 11, 17, 0, 0, tz);
  const busy = [{ start: zonedToUtc(2026, 11, 16, 9, 30, tz), end: zonedToUtc(2026, 11, 16, 10, 0, tz) }];
  assert.equal(findFreeRanges({ from, to, busy, durationMin: 45, workingHours: wh }).length, 1);
  const r = findFreeRanges({ from, to, busy, durationMin: 30, workingHours: wh, bufferMin: 15 });
  assert.equal(new Date(r[0].start).toISOString(), "2026-11-16T09:15:00.000Z");
});
