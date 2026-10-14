import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { trayHarness, tick } from "./helpers/tray-supervisor.mjs";

const pending = (h, ms) => h.timers.find((timer) => !timer.cancelled && timer.ms === ms);
test("daemon created without readiness times out, stops only captured child and never retries", async () => {
  const phases = [];
  const h = await trayHarness({
    overrides: {
      desktopLifecycle: {
        publish: (phase, fields) => phases.push({ phase, ...fields }),
        record() {},
      },
    },
  });
  await h.start();
  const child = h.owner();
  assert.equal(h.status(), "starting");
  assert.equal(pending(h, 45_000), undefined, "unready lifetime cannot refill restart budget");
  pending(h, 20_000).fn();
  await tick();
  assert.equal(h.status(), "blocked");
  assert.deepEqual(child.signals, ["SIGTERM"]);
  assert.equal(h.owner(), child, "timeout does not pretend the owned PID has exited");
  child.emit("message", { type: "daemon_ready" });
  assert.equal(h.status(), "blocked", "late handshake cannot revive a revoked startup");
  child.emit("exit", 0, null);
  assert.equal(h.status(), "blocked");
  assert.equal(h.children.length, 1);
  assert.equal(h.timers.filter((timer) => !timer.cancelled).length, 0);
  assert.ok(
    phases.some(
      (value) => value.phase === "startup-failed" && value.code === "DESKTOP_DAEMON_START_TIMEOUT",
    ),
  );
});

test("daemon handshake cancels timeout and only ready/stable lifetime refills bounded crash budget", async () => {
  const h = await trayHarness();
  await h.start();
  h.owner().emit("exit", 1, null);
  pending(h, 1000).fn();
  await tick();
  const timeout = pending(h, 20_000);
  h.owner().emit("message", { type: "daemon_ready" });
  assert.equal(timeout.cancelled, true);
  assert.equal(h.status(), "running");
  pending(h, 45_000).fn();
  h.owner().emit("exit", 1, null);
  assert.ok(pending(h, 1000), "successful stable handshake refills first restart budget");
  h.stop();
});

test("external SIGTERM is not fabricated as user Stop; exact Stop intent suppresses recovery", async () => {
  const h = await trayHarness();
  await h.start();
  h.owner().emit("exit", null, "SIGTERM");
  assert.equal(h.status(), "crashed");
  assert.ok(pending(h, 1000));
  h.stop();
  await h.start();
  const child = h.owner();
  h.stop();
  child.emit("exit", 0, null);
  assert.equal(h.status(), "stopped");
  assert.equal(h.timers.filter((timer) => !timer.cancelled && timer.ms <= 8000).length, 0);
});

test("the private launcher channel is never inherited by daemon or SDK children", async () => {
  let launched;
  const h = await trayHarness({
    spawnChild(_entry, _args, options) {
      launched = options.env;
      return Object.assign(new EventEmitter(), {
        stdout: new EventEmitter(),
        stderr: new EventEmitter(),
        connected: true,
        send() {},
        kill() {},
      });
    },
    overrides: {
      process: {
        execPath: "/synthetic/node",
        env: {
          DRAFTSPECT_LAUNCH_STATUS_PATH: "/synthetic/private-status.json",
          DRAFTSPECT_LAUNCH_TOKEN: "SECRET_TOKEN",
          DRAFTSPECT_LAUNCH_PARENT_PID: "42",
          DRAFTSPECT_RUNTIME: "codebuddy",
        },
      },
    },
  });
  await h.start();
  assert.equal(launched.ELECTRON_RUN_AS_NODE, "1");
  assert.equal(launched.DRAFTSPECT_RUNTIME, "codebuddy");
  assert.equal(launched.DRAFTSPECT_LAUNCH_STATUS_PATH, undefined);
  assert.equal(launched.DRAFTSPECT_LAUNCH_TOKEN, undefined);
  assert.equal(launched.DRAFTSPECT_LAUNCH_PARENT_PID, undefined);
  h.stop();
});

test("late native-picker result from an exited daemon never crosses into its replacement", async () => {
  let finish;
  const sent = [];
  const h = await trayHarness({
    spawnChild() {
      return Object.assign(new EventEmitter(), {
        stdout: new EventEmitter(),
        stderr: new EventEmitter(),
        connected: true,
        send: (value) => sent.push(value),
        kill() {},
      });
    },
    overrides: {
      dialog: {
        showOpenDialog: () =>
          new Promise((resolve) => {
            finish = resolve;
          }),
      },
    },
  });
  await h.start();
  h.owner().emit("message", { type: "pick_path", id: "old-picker" });
  await tick();
  h.owner().emit("exit", 1, null);
  await h.start();
  finish({ canceled: true });
  await tick();
  assert.deepEqual(sent, [], "the reply belongs to the old captured IPC owner");
  h.stop();
});

test("actual main exception/rejection and utility faults are distinct sanitized events with no replay", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const first = source.indexOf(
    "// JavaScript exceptions occur after native Electron initialization.",
  );
  const last = source.indexOf("async function handleBeforeQuit(", first);
  assert.ok(first >= 0 && last > first);
  for (const event of ["uncaughtException", "unhandledRejection"]) {
    const app = new EventEmitter(),
      processEvents = new EventEmitter(),
      recorded = [],
      phases = [],
      exits = [];
    let stops = 0,
      cancels = 0;
    app.exit = (code) => exits.push(code);
    const context = vm.createContext({
      ownsDesktopInstance: true,
      fatalMainError: false,
      appIsQuitting: false,
      app,
      process: processEvents,
      daemonProcess: { owned: true },
      desktopLifecycle: { record: (_event, value) => recorded.push(value), close() {} },
      publishDesktopPhase: (phase, value) => phases.push({ phase, ...value }),
      desktopIncident: (value) => recorded.push(value),
      cancelScheduledRestart: () => cancels++,
      stopDaemon: () => stops++,
      waitForOwnedExit: async (_child, stop) => stop(),
    });
    vm.runInContext(source.slice(first, last), context);
    app.emit(
      "child-process-gone",
      {},
      { type: "GPU", reason: "crashed", exitCode: 9, name: "SECRET" },
    );
    assert.equal(recorded[0].code, "DESKTOP_CHILD_GONE");
    assert.equal(recorded[0].processType, "GPU");
    processEvents.emit(event, new Error("SECRET_MAIN_ERROR"));
    processEvents.emit(event, new Error("SECOND_SECRET"));
    await tick();
    assert.equal(stops, 1);
    assert.equal(cancels, 1);
    assert.deepEqual(exits, [1]);
    assert.equal(phases[0].phase, "startup-failed");
    assert.equal(
      phases[0].code,
      event === "uncaughtException" ? "DESKTOP_MAIN_EXCEPTION" : "DESKTOP_MAIN_REJECTION",
    );
    assert.doesNotMatch(JSON.stringify(recorded) + JSON.stringify(phases), /SECRET/);
    const count = recorded.length;
    app.emit("child-process-gone", {}, { type: "GPU", reason: "killed" });
    assert.equal(recorded.length, count, "manual/fatal quit suppresses child recovery noise");
  }
});

test("diagnostic failure is a visible warning and never invents a local-service startup failure", async () => {
  const h = await trayHarness({
    overrides: { desktopLifecycle: { diagnosticAvailable: false, publish() {}, record() {} } },
  });
  const snapshot = h.getDesktopWindow().options.readStatus();
  assert.match(snapshot.detail, /DESKTOP_DIAGNOSTIC_UNAVAILABLE/);
  assert.equal(snapshot.error, "");
  assert.equal(h.status(), "starting");
  await h.start();
  h.owner().emit("message", { type: "daemon_ready" });
  assert.equal(h.status(), "running");
  assert.equal(h.getDesktopWindow().options.readStatus().error, "");
  h.stop();
});

test("manual Stop remains stopping until its exact captured process exit is confirmed", async () => {
  const h = await trayHarness();
  await h.start();
  const owned = h.owner();
  owned.emit("message", { type: "daemon_ready" });
  h.stop();
  assert.equal(h.status(), "stopping");
  assert.equal(h.owner(), owned);
  owned.emit("exit", 0, null);
  assert.equal(h.status(), "stopped");
  assert.equal(h.owner(), null);
});
