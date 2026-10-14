import { test } from "node:test";
import assert from "node:assert/strict";
import { createHealthWatch } from "../app/health-watch.mjs";

test("tray watchdog detects silence without restarting and rejects stale/future replies", () => {
  let clock = 0;
  const w = createHealthWatch({ now: () => clock, timeoutMs: 30 });
  const id = w.request();
  clock = 31;
  assert.equal(w.unresponsive(), true);
  assert.equal(w.reply(id + 1), false);
  assert.equal(w.reply(id), true);
  clock = 62;
  assert.equal(w.reply(id), false);
  assert.equal(w.unresponsive(), true);
});
test("tray watchdog allows sleep/wake grace without accepting pre-sleep replies", () => {
  let clock = 0;
  const w = createHealthWatch({ now: () => clock, timeoutMs: 30 });
  const old = w.request();
  w.suspend();
  clock = 1000;
  assert.equal(w.unresponsive(), false);
  w.resume();
  assert.equal(w.reply(old), false);
  assert.equal(w.unresponsive(), false);
  clock += 31;
  assert.equal(w.unresponsive(), true);
});
