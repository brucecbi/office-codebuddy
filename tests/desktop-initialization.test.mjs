// Execute the actual initialization and registration code, without running its
// ready callbacks, Electron, providers, filesystem writes or production daemon.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import vm from "node:vm";
import { stateDirectory, stateFile } from "../daemon/state-paths.mjs";

const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
const start = source.indexOf("let SESSIONS_FILE;");
assert.ok(start >= 0, "Desktop initialization boundary changed");
const initialization = source.slice(start);

function initialize({
  env = { DRAFTSPECT_STATE_DIR: "/synthetic/state" },
  duplicate = false,
  lockError,
} = {}) {
  const observed = { dialogs: [], quits: 0, locks: [], readyCallbacks: [], listeners: [] };
  const app = {
    quit() {
      observed.quits++;
    },
    whenReady() {
      return {
        then(callback) {
          observed.readyCallbacks.push(callback);
          return { catch() {} };
        },
      };
    },
    on(event, callback) {
      observed.listeners.push({ event, callback });
    },
  };
  const context = vm.createContext({
    app,
    setInterval: () => ({ unref() {} }),
    clearInterval() {},
    dialog: { showErrorBox: (...args) => observed.dialogs.push(args) },
    createDesktopLifecycle: () => ({ publish() {}, record() {}, close() {} }),
    process: { on() {} },
    stateDirectory: () => stateDirectory(env),
    stateFile: (name) => stateFile(name, env),
    acquireDesktopInstance: (options) => {
      observed.locks.push(options);
      if (lockError) throw lockError;
      if (duplicate) app.quit();
      return !duplicate;
    },
  });
  vm.runInContext(initialization, context);
  observed.state = vm.runInContext("({ SESSIONS_FILE, LOG_FILE, ownsDesktopInstance })", context);
  return observed;
}

test("forbidden legacy state fails inside controlled initialization before lock or startup", () => {
  const observed = initialize({
    env: { DRAFTSPECT_STATE_DIR: join(homedir(), ".claude", "office-addins") },
  });
  assert.equal(observed.dialogs.length, 1);
  assert.match(observed.dialogs[0][0], /Draftspect/);
  assert.match(observed.dialogs[0][1], /状态目录或应用锁不可用/);
  assert.equal(observed.quits, 1);
  assert.equal(observed.locks.length, 0);
  assert.equal(observed.state.ownsDesktopInstance, false);
  assert.equal(observed.readyCallbacks.length, 0);
  assert.equal(observed.listeners.length, 0);
});

test("lock failure is controlled and sanitized with no startup or lifecycle registration", () => {
  const observed = initialize({ lockError: new Error("SECRET_TEST_VALUE") });
  assert.equal(observed.dialogs.length, 1);
  assert.doesNotMatch(JSON.stringify(observed.dialogs), /SECRET_TEST_VALUE/);
  assert.equal(observed.quits, 1);
  assert.equal(observed.locks.length, 1);
  assert.equal(observed.state.ownsDesktopInstance, false);
  assert.equal(observed.readyCallbacks.length, 0);
  assert.equal(observed.listeners.length, 0);
});

test("duplicate instance exits without readiness callbacks, lifecycle handlers or dialogs", () => {
  const observed = initialize({ duplicate: true });
  assert.equal(observed.quits, 1);
  assert.equal(observed.locks.length, 1);
  assert.equal(observed.state.ownsDesktopInstance, false);
  assert.equal(observed.readyCallbacks.length, 0);
  assert.equal(observed.listeners.length, 0);
  assert.equal(observed.dialogs.length, 0);
});

test("owner resolves effective state and registers only the expected guarded startup paths", () => {
  const observed = initialize();
  assert.equal(observed.state.ownsDesktopInstance, true);
  assert.equal(observed.state.SESSIONS_FILE, "/synthetic/state/sessions.json");
  assert.equal(observed.state.LOG_FILE, "/synthetic/state/daemon.log");
  assert.equal(observed.locks[0].stateDir, "/synthetic/state");
  assert.equal(typeof observed.locks[0].onSecondInstance, "function");
  assert.equal(observed.readyCallbacks.length, 2);
  assert.deepEqual(
    observed.listeners.map(({ event }) => event),
    ["window-all-closed", "child-process-gone", "will-quit", "before-quit"],
  );
  assert.equal(observed.quits, 0);
  assert.equal(observed.dialogs.length, 0);
});
