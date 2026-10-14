// Independent final integration review: actual launch function and actual
// persistence module. Synthetic dependencies only; no provider or daemon boot.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { createSessionPersistenceRegistry } from "../daemon/session-persistence.mjs";
import { createHealthWatch } from "../app/health-watch.mjs";
import { EventEmitter } from "node:events";

const identity = {
  host: "word",
  cwd: "/synthetic",
  provider: "codebuddy",
  session_id: "newest-id",
};
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((a, b) => {
    resolve = a;
    reject = b;
  });
  return { promise, resolve, reject };
};
const tick = () => new Promise((r) => setImmediate(r));

async function nextTurnHarness(save) {
  const raw = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const start = raw.indexOf("async function ensureLoopForMessage(");
  const end = raw.indexOf("\nconst bridge =", start);
  assert.ok(start >= 0 && end > start, "actual admission function boundary changed");
  const calls = [];
  const registry = createSessionPersistenceRegistry({ save });
  const workspaces = new Map([["pane", identity.cwd]]);
  let active = "new-turn";
  const switchStart = raw.indexOf("async function switchFolder(");
  const switchEnd = raw.indexOf("\n// Re-launch one", switchStart);
  const context = vm.createContext({
    sessionPersistence: registry,
    runtime: { id: "codebuddy" },
    bridge: { activeTurn: () => active, clearUserMessages() {} },
    cwdForKey: () => workspaces.get("pane"),
    handoffInProgress: new Set(),
    getHandoff: async () => null,
    sessionFor: () => null,
    getSessionId: async () => "older-disk-id",
    resolveReplaySessionId: async () =>
      registry.status("pane", identity)?.session_id ?? "older-disk-id",
    diag() {},
    scheduleSessionStart: (...args) => calls.push(args),
    resolve: (value) => value,
    assertAcceptanceWorkspace() {},
    stat: async () => ({ isDirectory: () => true }),
    workspaceByKey: workspaces,
    startQueue: new Map(),
    sendTranscriptReplayTo: async () => {},
  });
  vm.runInContext(
    raw.slice(start, end) +
      "\n" +
      raw.slice(switchStart, switchEnd) +
      "\nglobalThis.ensure=ensureLoopForMessage;globalThis.change= switchFolder;",
    context,
  );
  return {
    registry,
    calls,
    ensure: () => context.ensure("pane", "word"),
    switchWorkspace: (cwd) => context.change(cwd, "pane", "word"),
    workspace: () => workspaces.get("pane"),
    busy: (value) => {
      active = value;
    },
  };
}

test("RED next request after Stop cannot resume an older disk ID while latest persistence failed", async () => {
  const h = await nextTurnHarness(async () => {
    throw Object.assign(new Error("synthetic save denied"), { code: "EACCES" });
  });
  await assert.rejects(h.registry.persist("pane", identity), { code: "EACCES" });
  await assert.rejects(h.ensure(), /save denied|persistence|persist/i);
  assert.deepEqual(h.calls, [], "failed latest identity must not start an older conversation");
});

test("RED next request waits pending newest ID and never races ahead using stale disk state", async () => {
  const save = deferred();
  const h = await nextTurnHarness(() => save.promise);
  const pending = h.registry.persist("pane", identity);
  const starting = h.ensure();
  await tick();
  assert.deepEqual(h.calls, [], "startup ran before newest persistence settled");
  save.resolve();
  await Promise.all([pending, starting]);
  assert.equal(h.calls.length, 1);
  assert.equal(h.calls[0][1], "newest-id");
});

test("RED actual workspace switch cannot overwrite an unsaved failed identity with a new scope", async () => {
  const h = await nextTurnHarness(async () => {
    throw new Error("save failed");
  });
  h.busy(null);
  await assert.rejects(h.registry.persist("pane", identity), /save failed/);
  await assert.rejects(h.switchWorkspace("/other"), /save failed/);
  assert.equal(h.workspace(), identity.cwd);
  assert.equal(h.registry.status("pane", identity).session_id, "newest-id");
  assert.equal(h.registry.status("pane", identity).status, "failed");
});

test("RED actual workspace switch waits for old-scope persistence before changing identity", async () => {
  const save = deferred();
  const h = await nextTurnHarness(() => save.promise);
  h.busy(null);
  const saving = h.registry.persist("pane", identity);
  const switching = h.switchWorkspace("/other");
  await tick();
  assert.equal(h.workspace(), identity.cwd);
  save.resolve();
  await Promise.all([saving, switching]);
  assert.equal(h.workspace(), "/other");
  assert.equal(h.registry.status("pane", identity).status, "saved");
});

test("RED actual tray watchdog reports disconnected or asynchronously failed IPC without restarting the live daemon", async () => {
  const raw = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const report = raw.slice(
    raw.indexOf("function reportHealthIncident("),
    raw.indexOf("\nif (ownsDesktopInstance)", raw.indexOf("function reportHealthIncident(")),
  );
  const start = raw.slice(
    raw.indexOf("async function startDaemon("),
    raw.indexOf("\nfunction stopDaemon("),
  );
  const readyStart = raw.indexOf("function markDaemonReady(");
  const readyEnd = raw.indexOf(
    "\n// --------------------------------------------------------------------------\n// IPC:",
    readyStart,
  );
  assert.ok(readyStart >= 0 && readyEnd > readyStart, "actual readiness boundary changed");
  const ready = raw.slice(readyStart, readyEnd);
  for (const mode of ["disconnected", "async-error"]) {
    const child = new EventEmitter();
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();
    child.connected = true;
    let clock = 0,
      interval,
      spawns = 0;
    const sent = [],
      notifications = [];
    child.send = (message, callback) => {
      sent.push(message);
      setImmediate(() =>
        callback?.(
          Object.assign(new Error("IPC channel closed"), { code: "ERR_IPC_CHANNEL_CLOSED" }),
        ),
      );
    };
    class Notification {
      static isSupported() {
        return true;
      }
      constructor(options) {
        this.options = options;
      }
      show() {
        notifications.push(this.options);
      }
    }
    const context = vm.createContext({
      findInitialWorkspace: async () => null,
      spawn: () => {
        spawns++;
        return child;
      },
      process: { execPath: "/synthetic/node", env: {} },
      PROJECT_ROOT: "/synthetic",
      DAEMON_ENTRY: "/synthetic/daemon.mjs",
      daemonProcess: null,
      pendingDaemonStart: null,
      restartTimer: null,
      daemonExitIntent: new WeakMap(),
      daemonStatus: "starting",
      currentWorkspace: null,
      restartAttempts: 0,
      MAX_RESTART: 3,
      STABLE_AFTER_MS: 45000,
      stableTimer: null,
      daemonStartupTimer: null,
      daemonHealthTimer: null,
      daemonWatch: null,
      healthIncident: null,
      desktopStartupError: null,
      desktopLifecycle: { record() {} },
      publishDesktopPhase() {},
      desktopIncident() {},
      appIsQuitting: false,
      createHealthWatch: () => createHealthWatch({ now: () => clock, timeoutMs: 35000 }),
      setInterval: (fn) => {
        interval = fn;
        return 1;
      },
      clearInterval() {},
      setTimeout: () => 2,
      clearTimeout() {},
      logStream: { write() {} },
      updateTray() {},
      async handleDaemonMessage() {},
      Notification,
    });
    vm.runInContext(report + start + ready + "\nglobalThis.start=startDaemon;", context);
    await context.start();
    child.connected = mode !== "disconnected";
    clock = mode === "disconnected" ? 40000 : 10000;
    interval();
    await tick();
    assert.equal(spawns, 1, "watchdog must never restart a live daemon automatically");
    assert.equal(
      sent.length,
      mode === "disconnected" ? 0 : 1,
      "cannot send IPC request over a disconnected channel",
    );
    assert.ok(notifications.length > 0, "IPC disconnect silently bypassed the watchdog");
    assert.ok(context.healthIncident, "tray status still implies healthy daemon");
  }
});
