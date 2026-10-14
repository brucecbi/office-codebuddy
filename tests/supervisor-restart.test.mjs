import { test } from "node:test";
import assert from "node:assert/strict";
import { trayHarness, tick } from "./helpers/tray-supervisor.mjs";

const pendingRestarts = (h) => h.timers.filter((t) => !t.cancelled && t.ms <= 8000);

test("manual Restart after handled SIGTERM exit(0,null) creates exactly one replacement", async () => {
  const h = await trayHarness();
  await h.start();
  h.restart();
  h.restart(); // a second click must not register another exit-driven start
  h.children[0].emit("exit", 0, null);
  await tick();
  assert.equal(h.children.length, 2);
  assert.equal(pendingRestarts(h).length, 0, "intentional restart also queued a crash restart");
  assert.equal(h.owner(), h.children[1]);
});

test("Stop supersedes a pending intentional Restart and prevents its replacement", async () => {
  const h = await trayHarness();
  await h.start();
  h.restart();
  h.stop();
  h.children[0].emit("exit", 0, null);
  await tick();
  assert.equal(h.children.length, 1);
  assert.equal(h.owner(), null);
  assert.equal(h.status(), "stopped");
  assert.equal(pendingRestarts(h).length, 0);
});

test("manual start cancels crash backoff; an already queued stale callback cannot start again", async () => {
  const h = await trayHarness();
  await h.start();
  h.children[0].emit("exit", 1, null);
  const [timer] = pendingRestarts(h);
  assert.equal(timer.ms, 1000);
  await h.start();
  assert.equal(timer.cancelled, true);
  timer.fn(); // simulate callback already placed in the event queue
  await tick();
  assert.equal(h.children.length, 2);
  assert.equal(h.owner(), h.children[1]);
});

test("Stop cancels crash backoff even while no child exists", async () => {
  const h = await trayHarness();
  await h.start();
  h.children[0].emit("exit", 1, null);
  const [timer] = pendingRestarts(h);
  h.stop();
  timer.fn();
  await tick();
  assert.equal(h.children.length, 1);
  assert.equal(h.status(), "stopped");
});

test("async workspace discovery allows only one concurrent start and Stop fences it", async () => {
  let release;
  const workspace = new Promise((resolve) => {
    release = resolve;
  });
  const h = await trayHarness({ findWorkspace: () => workspace });
  const first = h.start();
  const second = h.start();
  h.stop();
  release(null);
  await Promise.all([first, second]);
  assert.equal(h.children.length, 0, "a cancelled async start launched after Stop");
  await Promise.all([h.start(), h.start()]);
  assert.equal(h.children.length, 1);
});

test("stale child exit cannot clear the current owner or its health monitor", async () => {
  const h = await trayHarness();
  await h.start();
  const old = h.children[0];
  old.emit("exit", 1, null);
  await h.start();
  const current = h.owner();
  old.emit("exit", 0, null); // defensive stale callback replay
  assert.equal(h.owner(), current);
  assert.equal(pendingRestarts(h).length, 0);
});

test("genuine crash restarts retain the bounded 1s/2s/4s budget", async () => {
  const h = await trayHarness();
  await h.start();
  for (const ms of [1000, 2000, 4000]) {
    h.owner().emit("exit", 1, null);
    const timer = h.timers.at(-1);
    assert.equal(timer.ms, ms);
    timer.fn();
    await tick();
  }
  h.owner().emit("exit", 1, null);
  assert.equal(h.children.length, 4);
  assert.match(h.logs.at(-1), /giving up after 3 restart attempts/);
});
