const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const qs = require("../src/lib/delay.js");

const TZ = "Europe/London";

function london(year, month, day, hour, minute) {
  return qs.zonedLocalToDate(year, month, day, hour, minute, TZ);
}

function partsOf(date) {
  return qs.zonedParts(date, TZ);
}

describe("offset delays from Send time", () => {
  it("adds whole hours to the provided instant", () => {
    const sendAt = new Date("2026-10-07T16:00:00.000Z");
    const one = qs.applyPresetToDate(
      { label: "1h", kind: "offset", hours: 1, minutes: 0 },
      sendAt,
      TZ
    );
    const two = qs.applyPresetToDate(
      { label: "2h", kind: "offset", hours: 2, minutes: 0 },
      sendAt,
      TZ
    );
    assert.equal(one.toISOString(), "2026-10-07T17:00:00.000Z");
    assert.equal(two.toISOString(), "2026-10-07T18:00:00.000Z");
  });

  it("does not use the button-click time — only the now passed in (Send)", () => {
    const clickedAt = new Date("2026-10-07T09:00:00.000Z");
    const sentAt = new Date("2026-10-07T16:00:00.000Z");
    const unusedClick = qs.addOffset(clickedAt, 2, 0);
    const applied = qs.applyPresetToDate(
      { label: "2h", kind: "offset", hours: 2, minutes: 0 },
      sentAt,
      TZ
    );
    assert.equal(unusedClick.toISOString(), "2026-10-07T11:00:00.000Z");
    assert.equal(applied.toISOString(), "2026-10-07T18:00:00.000Z");
  });

  it("supports hours plus minutes", () => {
    const sendAt = new Date("2026-10-07T10:00:00.000Z");
    const when = qs.addOffset(sendAt, 1, 30);
    assert.equal(when.toISOString(), "2026-10-07T11:30:00.000Z");
  });
});

describe("tomorrow 08:00 across midnight", () => {
  it("rolls to the next London calendar day after 23:30", () => {
    const now = london(2026, 10, 7, 23, 30);
    const when = qs.tomorrowAt(now, 8, 0, TZ);
    const parts = partsOf(when);
    assert.equal(parts.year, 2026);
    assert.equal(parts.month, 10);
    assert.equal(parts.day, 8);
    assert.equal(parts.hour, 8);
    assert.equal(parts.minute, 0);
  });

  it("still lands on 08:00 when Send is just after midnight", () => {
    const now = london(2026, 10, 8, 0, 15);
    const when = qs.applyPresetToDate(
      { label: "Tomorrow 8am", kind: "tomorrow", hour: 8, minute: 0 },
      now,
      TZ
    );
    const parts = partsOf(when);
    assert.equal(parts.day, 9);
    assert.equal(parts.hour, 8);
  });
});

describe("UK DST — BST / GMT change", () => {
  it("autumn fallback: Saturday 24 Oct 2026 23:30 BST → Sunday 25 Oct 08:00 GMT", () => {
    const now = london(2026, 10, 24, 23, 30);
    const saturday = partsOf(now);
    assert.equal(saturday.day, 24);
    assert.equal(saturday.hour, 23);
    // 23:30 BST is 22:30 UTC
    assert.equal(now.getUTCHours(), 22);

    const when = qs.tomorrowAt(now, 8, 0, TZ);
    const parts = partsOf(when);
    assert.equal(parts.year, 2026);
    assert.equal(parts.month, 10);
    assert.equal(parts.day, 25);
    assert.equal(parts.hour, 8);
    assert.equal(parts.minute, 0);
    // After 02:00 BST → 01:00 GMT, 08:00 London is GMT (UTC+0)
    assert.equal(when.getUTCHours(), 8);
    assert.equal(when.toISOString(), "2026-10-25T08:00:00.000Z");
  });

  it("spring forward: Saturday 28 Mar 2026 23:30 GMT → Sunday 29 Mar 08:00 BST", () => {
    const now = london(2026, 3, 28, 23, 30);
    assert.equal(now.getUTCHours(), 23);

    const when = qs.tomorrowAt(now, 8, 0, TZ);
    const parts = partsOf(when);
    assert.equal(parts.year, 2026);
    assert.equal(parts.month, 3);
    assert.equal(parts.day, 29);
    assert.equal(parts.hour, 8);
    assert.equal(parts.minute, 0);
    // After 01:00 GMT → 02:00 BST, 08:00 London is BST (UTC+1)
    assert.equal(when.getUTCHours(), 7);
    assert.equal(when.toISOString(), "2026-03-29T07:00:00.000Z");
  });
});

describe("copy and validation", () => {
  it("describes offset and tomorrow presets", () => {
    assert.equal(
      qs.describeSchedule({ label: "2h", kind: "offset", hours: 2, minutes: 0 }).message,
      "Will send 2 hours after you press Send"
    );
    assert.equal(
      qs.describeSchedule({ label: "Tomorrow 8am", kind: "tomorrow", hour: 8, minute: 0 }).message,
      "Will send tomorrow at 08:00 after you press Send"
    );
  });

  it("rejects a zero-length offset", () => {
    assert.equal(qs.normalizePreset({ kind: "offset", hours: 0, minutes: 0, label: "0" }), null);
    assert.throws(() => qs.addOffset(new Date(), 0, 0));
  });
});
