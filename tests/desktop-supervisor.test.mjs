// Exercise the actual tray start/stop functions with the desktop preflight
// branch enabled, including cancellation while the local check is in flight.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { EventEmitter } from "node:events";
import vm from "node:vm";
import { trayHarness, tick } from "./helpers/tray-supervisor.mjs";
import { waitForOwnedExit } from "../app/quit-barrier.mjs";

const valid = () => ({
  cliPath: "/synthetic/codebuddy",
  nodePath: "/synthetic/node",
  warnings: [],
});
const desktopProcess = () => ({
  execPath: "/synthetic/electron",
  env: { DRAFTSPECT_DESKTOP_LAUNCH: "1", DRAFTSPECT_RUNTIME: "codebuddy" },
});

test("desktop preflight denial is visible and starts no daemon or retry timer", async () => {
  const dialogs = [];
  const h = await trayHarness({
    overrides: {
      process: desktopProcess(),
      desktopPreflight: async () => {
        throw Object.assign(new Error("HTTPS required"), { code: "DESKTOP_TLS_REQUIRED" });
      },
      dialog: {
        showMessageBox: async (options) => {
          dialogs.push(options);
          return { response: 0 };
        },
      },
    },
  });
  await h.start();
  await tick();
  assert.equal(h.status(), "blocked");
  assert.equal(h.children.length, 0);
  assert.equal(h.timers.length, 0);
  assert.equal(dialogs.length, 0);
  assert.equal(h.desktopWindows.length, 1);
  assert.match(h.desktopWindows[0].shown[0].error, /DESKTOP_TLS_REQUIRED/);
  assert.match(h.desktopWindows[0].shown[0].detail, /不会自动重发/);
});

test("Stop during preflight fences both a late success and a late error", async () => {
  for (const fail of [false, true]) {
    let resolve, reject;
    let dialogs = 0;
    const envProcess = desktopProcess();
    const pending = new Promise((yes, no) => {
      resolve = yes;
      reject = no;
    });
    const h = await trayHarness({
      overrides: {
        process: envProcess,
        desktopPreflight: () => pending,
        dialog: {
          showMessageBox: async () => {
            dialogs++;
            return { response: 0 };
          },
        },
      },
    });
    const starting = h.start();
    h.stop();
    if (fail) reject(new Error("late failure"));
    else resolve(valid());
    await starting;
    assert.equal(h.status(), "stopped");
    assert.equal(h.children.length, 0);
    assert.equal(dialogs, 0);
    assert.equal(envProcess.env.CODEBUDDY_CODE_PATH, undefined);
  }
});

test("concurrent desktop starts share one check and only full IPC readiness marks ready", async () => {
  let resolve,
    checks = 0;
  const pending = new Promise((yes) => {
    resolve = yes;
  });
  const envProcess = desktopProcess();
  const h = await trayHarness({
    overrides: {
      process: envProcess,
      desktopPreflight: () => {
        checks++;
        return pending;
      },
    },
  });
  const a = h.start(),
    b = h.start();
  resolve(valid());
  await Promise.all([a, b]);
  assert.equal(checks, 1);
  assert.equal(h.children.length, 1);
  assert.equal(envProcess.env.CODEBUDDY_CODE_PATH, "/synthetic/codebuddy");
  assert.equal(envProcess.env.DRAFTSPECT_NODE_PATH, "/synthetic/node");
  h.owner().stdout.emit("data", "[daemon] HTTPS server listening on synthetic");
  assert.equal(h.status(), "starting", "one listening server is not full readiness");
  h.owner().emit("message", { type: "daemon_ready" });
  assert.equal(h.status(), "running");
  h.stop();
});

test("synchronous spawn failure is visible without leaking an arbitrary error message", async () => {
  const dialogs = [];
  const h = await trayHarness({
    spawnChild() {
      throw new Error("SECRET_TEST_VALUE");
    },
    overrides: {
      process: desktopProcess(),
      desktopPreflight: async () => valid(),
      dialog: {
        showMessageBox: async (options) => {
          dialogs.push(options);
          return { response: 0 };
        },
      },
    },
  });
  await h.start();
  await tick();
  assert.equal(h.status(), "blocked");
  assert.equal(h.owner(), null);
  assert.equal(dialogs.length, 0);
  assert.equal(h.desktopWindows.length, 1);
  assert.match(h.desktopWindows[0].shown[0].error, /DESKTOP_START_FAILED/);
  assert.doesNotMatch(h.desktopWindows[0].shown[0].error, /SECRET_TEST_VALUE/);
  assert.doesNotMatch(h.logs.join("\n"), /SECRET_TEST_VALUE/);
});

test("asynchronous spawn failure does not crash the parent or schedule restart", async () => {
  const h = await trayHarness({
    overrides: { process: desktopProcess(), desktopPreflight: async () => valid() },
  });
  await h.start();
  const child = h.owner();
  child.emit("error", new Error("spawn failure"));
  child.emit("exit", -2, null);
  await tick();
  assert.equal(h.owner(), null);
  assert.equal(h.status(), "blocked");
  assert.equal(h.children.length, 1);
  assert.equal(h.timers.filter((t) => !t.cancelled).length, 0);
});

test("actual log stream reports async write failure, cancels owned startup, and never logs raw error", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const stream = new EventEmitter();
  stream.write = () => {};
  let stops = 0,
    options;
  const errors = [];
  const context = vm.createContext({
    fs: {
      mkdirSync() {},
      createWriteStream(_path, opts) {
        options = opts;
        return stream;
      },
    },
    dirname: () => "/synthetic",
    LOG_FILE: "/synthetic/daemon.log",
    logStream: null,
    appIsQuitting: false,
    stopDaemon: () => stops++,
    reportStartupFailure: (error) => errors.push(error),
  });
  vm.runInContext(
    source.slice(
      source.indexOf("function openLogStream("),
      source.indexOf("async function startDaemon("),
    ) + "\nopenLogStream();",
    context,
  );
  stream.emit("error", new Error("SECRET_TEST_VALUE"));
  assert.equal(context.logStream, null);
  assert.equal(options.mode, 0o600);
  assert.equal(stops, 1);
  assert.equal(errors[0].code, "DESKTOP_LOG_UNAVAILABLE");
  assert.doesNotMatch(errors[0].message, /SECRET_TEST_VALUE/);
  stream.emit("error", new Error("late failure"));
  assert.equal(stops, 1);
});

test("actual Quit handler holds ownership until the child exits and coalesces repeated Quit", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const child = Object.assign(new EventEmitter(), { pid: 101, exitCode: null, signalCode: null });
  let quits = 0,
    stops = 0,
    prevented = 0,
    disposed = 0;
  const context = vm.createContext({
    desktopQuitReady: false,
    publishDesktopPhase() {},
    desktopQuitPending: false,
    appIsQuitting: false,
    daemonProcess: child,
    desktopWindow: {
      dispose() {
        disposed++;
      },
    },
    logStream: { write() {}, end() {} },
    app: {
      quit() {
        assert.equal(
          disposed,
          1,
          "the close-to-hide window must be disposed before final app.quit",
        );
        quits++;
      },
    },
    stopDaemon() {
      stops++;
    },
    waitForOwnedExit,
    reportStartupFailure: (error) => {
      throw error;
    },
  });
  vm.runInContext(
    source.slice(
      source.indexOf("async function handleBeforeQuit("),
      source.lastIndexOf("if (ownsDesktopInstance)"),
    ) + "\nglobalThis.quitHandler = handleBeforeQuit;",
    context,
  );
  const event = {
    preventDefault() {
      prevented++;
    },
  };
  const first = context.quitHandler(event);
  await context.quitHandler(event);
  assert.equal(quits, 0);
  assert.equal(disposed, 0, "window stays available while awaiting actual owned-child exit");
  assert.equal(stops, 1);
  assert.equal(prevented, 2);
  assert.equal(context.appIsQuitting, true);
  child.emit("exit", 0, null);
  await first;
  assert.equal(quits, 1);
  assert.equal(disposed, 1);
  assert.equal(context.desktopQuitReady, true);
  await context.quitHandler(event);
  assert.equal(prevented, 2, "final quit must proceed without another preventDefault");
});

test("actual Quit timeout keeps the app alive and shows a recoverable error", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const child = Object.assign(new EventEmitter(), { pid: 101, exitCode: null, signalCode: null });
  let quits = 0;
  let disposed = 0;
  const errors = [];
  const context = vm.createContext({
    desktopQuitReady: false,
    publishDesktopPhase() {},
    desktopQuitPending: false,
    appIsQuitting: false,
    daemonProcess: child,
    desktopWindow: {
      dispose() {
        disposed++;
      },
    },
    logStream: { write() {}, end() {} },
    app: {
      quit() {
        quits++;
      },
    },
    stopDaemon() {},
    waitForOwnedExit: (target, stop) => waitForOwnedExit(target, stop, { timeoutMs: 10 }),
    reportStartupFailure: async (error) => errors.push(error),
  });
  vm.runInContext(
    source.slice(
      source.indexOf("async function handleBeforeQuit("),
      source.lastIndexOf("if (ownsDesktopInstance)"),
    ) + "\nglobalThis.quitHandler = handleBeforeQuit;",
    context,
  );
  await context.quitHandler({ preventDefault() {} });
  assert.equal(quits, 0);
  assert.equal(disposed, 0, "timeout must leave controls available");
  assert.equal(errors[0].code, "DESKTOP_STOP_TIMEOUT");
  assert.equal(context.desktopQuitReady, false);
  assert.equal(context.desktopQuitPending, false);
  assert.equal(context.appIsQuitting, false);
});

test("desktop HTML load rejection falls back to a sanitized native dialog", async () => {
  const dialogs = [];
  let disposed = 0;
  const h = await trayHarness({
    overrides: {
      process: desktopProcess(),
      createDesktopStatusWindow: () => ({
        show: async () => {
          throw new Error("SECRET_TEST_LOAD_ERROR");
        },
        dispose: () => {
          disposed++;
        },
      }),
      dialog: {
        showMessageBox: async (options) => {
          dialogs.push(options);
          return { response: 0 };
        },
      },
    },
  });
  await h.showDesktopStatus();
  assert.equal(disposed, 1);
  assert.equal(dialogs.length, 1);
  assert.match(dialogs[0].detail, /DESKTOP_STATUS_LOAD_FAILED/);
  assert.doesNotMatch(JSON.stringify(dialogs) + h.logs.join("\n"), /SECRET_TEST_LOAD_ERROR/);
});

test("desktop stopped restart returns its startup promise and contains workspace rejections", async () => {
  const h = await trayHarness({
    findWorkspace: async () => {
      throw new Error("SECRET_WORKSPACE_ERROR");
    },
    overrides: { process: desktopProcess(), desktopPreflight: async () => valid() },
  });
  const result = h.restart();
  assert.equal(typeof result.then, "function");
  await result;
  assert.equal(h.status(), "blocked");
  assert.equal(h.children.length, 0);
  assert.match(h.desktopWindows[0].shown[0].error, /DESKTOP_START_FAILED/);
  assert.doesNotMatch(JSON.stringify(h.desktopWindows[0].shown), /SECRET_WORKSPACE_ERROR/);
});

test("actual desktop action reports failed log opening and each host initialization", async () => {
  const opened = [];
  const h = await trayHarness({
    overrides: {
      process: desktopProcess(),
      shell: {
        openPath: async (path) => {
          opened.push(path);
          return "Synthetic open failure";
        },
      },
      runInitializeWordAddin: async () => false,
      runInitializeExcelAddin: async () => false,
    },
  });
  h.getDesktopWindow();
  const action = h.desktopWindows[0].options.onAction;
  await assert.rejects(action("open-logs"), /无法打开日志/);
  assert.deepEqual(opened, ["/synthetic/daemon.log"]);
  await assert.rejects(action("initialize-word"), /Word 初始化未完成/);
  await assert.rejects(action("initialize-excel"), /Excel 初始化未完成/);
  await assert.rejects(action("arbitrary"), /不支持的操作/);
});
