// Unit tests for the activity-based turn timeout (daemon/runtime/lifecycle.mjs).
//
// Regression guard for the 2026-09-20 incident: a long read-only analysis turn
// on a real contract was killed by the fixed 180s per-turn wall even though the
// model and Office tools were making steady progress. The lifecycle must now
// (a) time out only after a period of INACTIVITY, (b) keep a configurable hard
// ceiling, and (c) say whether Office edits were in flight when it gives up.

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { createLifecycle } from "../daemon/runtime/lifecycle.mjs";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Execute the actual lifecycle with a deterministic scheduler. Real 30/60ms
// timers can lose the keepalive to CPU contention in the full parallel suite;
// that measures host scheduling rather than the activity/ceiling contract.
const lifecycleSource = await readFile(
  new URL("../daemon/runtime/lifecycle.mjs", import.meta.url),
  "utf8",
);
function lifecycleClock() {
  const timers = new Map();
  let now = 0;
  const context = {
    process,
    setTimeout(fn, ms) {
      const timer = { at: now + ms, fn, unref() {} };
      timers.set(timer, timer);
      return timer;
    },
    clearTimeout(timer) {
      timers.delete(timer);
    },
  };
  const create = vm.runInNewContext(
    lifecycleSource.replace(/^export /gm, "") + "\ncreateLifecycle",
    context,
  );
  return {
    create,
    advance(ms) {
      const end = now + ms;
      while (true) {
        const timer = [...timers.values()]
          .filter((t) => t.at <= end)
          .sort((a, b) => a.at - b.at)[0];
        if (!timer) break;
        now = timer.at;
        timers.delete(timer);
        timer.fn();
      }
      now = end;
    },
  };
}

test("idle timeout fires only after inactivity; activity() resets it", () => {
  const events = [];
  const clock = lifecycleClock();
  const s = clock.create({
    emit: (e) => events.push(e),
    cleanup() {},
    idleMs: 80,
    timeoutMs: 5000,
  });
  s.begin("t1");
  for (let i = 0; i < 5; i++) {
    clock.advance(40);
    s.activity();
  }
  assert.equal(events.length, 0, "steady activity must keep the turn alive past idleMs");
  clock.advance(79);
  assert.equal(events.length, 0, "idle deadline has not elapsed");
  clock.advance(1);
  assert.equal(events.length, 1);
  assert.equal(events[0].subtype, "timeout");
  assert.match(events[0].error, /no activity/i);
  assert.equal(s.closed, true);
});

test("hard ceiling fires even with continuous activity", () => {
  const events = [];
  const clock = lifecycleClock();
  const s = clock.create({
    emit: (e) => events.push(e),
    cleanup() {},
    idleMs: 60,
    timeoutMs: 150,
  });
  s.begin("t2");
  for (let i = 0; i < 4; i++) {
    clock.advance(30);
    s.activity();
  }
  assert.equal(events.length, 0);
  clock.advance(29);
  assert.equal(events.length, 0, "hard deadline has not elapsed");
  s.activity();
  clock.advance(1);
  assert.equal(events.length, 1);
  assert.equal(events[0].subtype, "timeout");
  assert.match(events[0].error, /limit/i);
  assert.equal(s.closed, true);
});

test("timeout copy says 'unknown outcome' only when Office writes were pending", async () => {
  const withWrites = [];
  const a = createLifecycle({
    emit: (e) => withWrites.push(e),
    cleanup() {},
    idleMs: 30,
    writesPending: () => true,
  });
  a.begin("w");
  await delay(80);
  assert.match(withWrites[0].error, /unknown outcome/);

  const readOnly = [];
  const b = createLifecycle({
    emit: (e) => readOnly.push(e),
    cleanup() {},
    idleMs: 30,
    writesPending: () => false,
  });
  b.begin("r");
  await delay(80);
  assert.match(readOnly[0].error, /No Office edit call is pending/);
  assert.doesNotMatch(readOnly[0].error, /document is unchanged/);
  assert.doesNotMatch(readOnly[0].error, /unknown outcome/);
});

test("writesPending defaults to the conservative unknown-outcome copy", async () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {}, idleMs: 30 });
  s.begin("default");
  await delay(80);
  assert.match(events[0].error, /unknown outcome/);
});

test("activity() after finish is a no-op", async () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {}, idleMs: 40 });
  s.begin("done");
  s.finish({ subtype: "success" });
  s.activity();
  await delay(90);
  assert.equal(events.length, 1);
  assert.equal(events[0].subtype, "success");
});

test("back-compat: bare timeoutMs still acts as the turn ceiling", async () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {}, timeoutMs: 20 });
  s.begin("hung");
  await delay(50);
  assert.equal(events.length, 1);
  assert.equal(events[0].subtype, "timeout");
  assert.equal(s.closed, true);
});

test("env DRAFTSPECT_TURN_IDLE_MS configures the idle default", async () => {
  process.env.DRAFTSPECT_TURN_IDLE_MS = "40";
  try {
    const events = [];
    const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {} });
    s.begin("env");
    await delay(90);
    assert.equal(events.length, 1);
    assert.match(events[0].error, /no activity/i);
  } finally {
    delete process.env.DRAFTSPECT_TURN_IDLE_MS;
  }
});

test("fail() forwards extra terminal fields such as the auth flag", () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {} });
  s.begin("t-auth");
  s.fail("401 Unauthorized", "runtime_error", { auth: true });
  assert.equal(events.length, 1);
  assert.equal(events[0].event, "turn_complete");
  assert.equal(events[0].auth, true);
  assert.equal(events[0].error, "401 Unauthorized");
  assert.equal(s.closed, true);
});
