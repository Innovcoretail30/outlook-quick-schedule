const { describe, it } = require("node:test");
const assert = require("node:assert/strict");
const qs = require("../src/lib/delay.js");

const twoHours = { label: "2h", kind: "offset", hours: 2, minutes: 0 };

describe("item isolation — no writes on select/read", () => {
  it("selecting another item plans zero writes", () => {
    const stored = qs.serializeSchedule(twoHours, {
      itemId: "compose-1",
      conversationId: "conv-1"
    });
    const plan = qs.planWrites({
      type: "select",
      item: { surface: "read", itemId: "inbox-9", conversationId: "conv-9" },
      sessionData: stored,
      customProperties: stored
    });
    assert.deepEqual(plan.writes, []);
    assert.equal(plan.applyDelay, false);
  });

  it("reading another item plans zero writes", () => {
    const stored = qs.serializeSchedule(twoHours, { itemId: "compose-1" });
    const plan = qs.planWrites({
      type: "read",
      item: { surface: "read", itemId: "inbox-2" },
      sessionData: null,
      customProperties: stored
    });
    assert.deepEqual(plan.writes, []);
  });

  it("item-changed (switching the reading pane) plans zero writes", () => {
    const plan = qs.planWrites({
      type: "itemChanged",
      item: { surface: "read", itemId: "inbox-3" },
      sessionData: qs.serializeSchedule(twoHours, { itemId: "compose-1" }),
      customProperties: qs.serializeSchedule(twoHours, { itemId: "compose-1" })
    });
    assert.deepEqual(plan.writes, []);
    assert.equal(plan.applyDelay, false);
  });
});

describe("item isolation — replies and forwards are not scheduled", () => {
  it("ignores a parent schedule left in custom properties on a reply", () => {
    const parentSchedule = qs.serializeSchedule(twoHours, {
      itemId: "orig-1",
      conversationId: "thread-1"
    });
    const reply = {
      surface: "compose",
      itemId: "reply-1",
      conversationId: "thread-1",
      sessionData: true
    };
    const fromCustom = qs.parseSchedule(parentSchedule, reply);
    assert.equal(fromCustom, null);

    const send = qs.planWrites({
      type: "send",
      item: reply,
      sessionData: null,
      customProperties: parentSchedule
    });
    assert.equal(send.applyDelay, false);
    assert.deepEqual(send.writes, []);
    assert.equal(send.allowEvent, true);
  });

  it("ignores a stored schedule whose itemId is a different message", () => {
    const stored = qs.parseStoredSchedule(
      qs.serializeSchedule(twoHours, { itemId: "orig-1", conversationId: "thread-1" })
    );
    const applied = qs.scheduleAppliesToItem(stored, {
      surface: "compose",
      itemId: "reply-1",
      conversationId: "thread-1",
      sessionData: true
    });
    assert.equal(applied, null);
  });
});

describe("item isolation — send with no preset", () => {
  it("allows send immediately and plans no writes", () => {
    const plan = qs.planWrites({
      type: "send",
      item: { surface: "compose", itemId: "new-1", sessionData: true },
      sessionData: null,
      customProperties: null
    });
    assert.equal(plan.allowEvent, true);
    assert.equal(plan.applyDelay, false);
    assert.deepEqual(plan.writes, []);
    assert.equal(plan.preset, null);
  });

  it("still does nothing if leftover custom properties exist but sessionData does not", () => {
    const leaked = qs.serializeSchedule(twoHours, { itemId: "old-1" });
    const plan = qs.planWrites({
      type: "send",
      item: { surface: "compose", itemId: "new-2", sessionData: true },
      sessionData: "",
      customProperties: leaked
    });
    assert.equal(plan.applyDelay, false);
    assert.deepEqual(plan.writes, []);
  });
});

describe("item isolation — explicit compose click still writes the current item only", () => {
  it("plans sessionData + notification for a compose time click", () => {
    const plan = qs.planWrites({
      type: "composeClick",
      action: "choose",
      item: { surface: "compose", itemId: "c1", sessionData: true },
      preset: twoHours
    });
    assert.deepEqual(plan.writes, ["sessionData.set", "notification"]);
    assert.equal(plan.applyDelay, false);
  });

  it("plans no item writes for a time click on the read surface", () => {
    const plan = qs.planWrites({
      type: "composeClick",
      action: "choose",
      item: { surface: "read", itemId: "inbox-1" },
      preset: twoHours
    });
    assert.deepEqual(plan.writes, []);
  });
});
