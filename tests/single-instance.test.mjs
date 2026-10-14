import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { spawnSync } from "node:child_process";
import {
  existsSync,
  mkdtempSync,
  mkdirSync,
  realpathSync,
  rmSync,
  statSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { acquireDesktopInstance } from "../app/single-instance.mjs";

function fixture(t) {
  const root = mkdtempSync(join(tmpdir(), "draftspect-desktop-lock-"));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  return realpathSync(root);
}

class AppStub extends EventEmitter {
  calls = [];
  lockResult = true;

  setName(name) {
    this.calls.push(["setName", name]);
  }
  setPath(name, value) {
    this.calls.push(["setPath", name, value]);
    this.userData = value;
  }
  requestSingleInstanceLock() {
    this.calls.push(["lock"]);
    return this.lockResult;
  }
  on(name, callback) {
    this.calls.push(["on", name]);
    return super.on(name, callback);
  }
  releaseSingleInstanceLock() {
    this.calls.push(["release"]);
  }
  quit() {
    this.calls.push(["quit"]);
  }
  exit(code) {
    this.calls.push(["exit", code]);
  }
}

test("desktop lock configures a stable private profile before acquiring ownership", (t) => {
  const root = fixture(t);
  const app = new AppStub();
  const stateDir = join(root, "state");
  assert.equal(acquireDesktopInstance({ app, stateDir }), true);
  assert.deepEqual(app.calls, [
    ["setName", "Draftspect"],
    ["setPath", "userData", join(stateDir, "desktop")],
    ["lock"],
    ["on", "second-instance"],
  ]);
  assert.equal(statSync(app.userData).mode & 0o777, 0o700);
});

test("duplicate instance quits without registering listeners or invoking the callback", (t) => {
  const app = new AppStub();
  app.lockResult = false;
  let callbacks = 0;
  assert.equal(
    acquireDesktopInstance({
      app,
      stateDir: fixture(t),
      onSecondInstance: () => callbacks++,
    }),
    false,
  );
  assert.equal(callbacks, 0);
  assert.equal(app.listenerCount("second-instance"), 0);
  assert.equal(
    app.calls.some(([name]) => name === "on" || name === "release"),
    false,
  );
  assert.deepEqual(app.calls.at(-1), ["quit"]);
});

test("only a successful owner receives the second-instance event and arguments", (t) => {
  const app = new AppStub();
  const received = [];
  assert.equal(
    acquireDesktopInstance({
      app,
      stateDir: fixture(t),
      onSecondInstance: (...args) => received.push(args),
    }),
    true,
  );
  assert.deepEqual(received, []);
  const args = [{}, ["launcher"], "/synthetic", { source: "desktop" }];
  app.emit("second-instance", ...args);
  assert.deepEqual(received, [args]);
});

test("explicit state directories isolate profiles and equivalent root aliases share one", (t) => {
  const root = fixture(t);
  const first = new AppStub();
  const second = new AppStub();
  const alias = new AppStub();
  const firstState = join(root, "first");
  const secondState = join(root, "second");
  assert.equal(acquireDesktopInstance({ app: first, stateDir: firstState }), true);
  assert.equal(acquireDesktopInstance({ app: second, stateDir: secondState }), true);
  const linkedState = join(root, "linked-first");
  symlinkSync(firstState, linkedState, "dir");
  assert.equal(acquireDesktopInstance({ app: alias, stateDir: linkedState }), true);
  assert.notEqual(first.userData, second.userData);
  assert.equal(first.userData, alias.userData);
});

test("omitted stateDir uses effective Draftspect state, independent of checkout identity", (t) => {
  const root = fixture(t);
  const previous = process.env.DRAFTSPECT_STATE_DIR;
  process.env.DRAFTSPECT_STATE_DIR = root;
  t.after(() => {
    if (previous === undefined) delete process.env.DRAFTSPECT_STATE_DIR;
    else process.env.DRAFTSPECT_STATE_DIR = previous;
  });
  const first = new AppStub();
  const second = new AppStub();
  assert.equal(acquireDesktopInstance({ app: first }), true);
  assert.equal(acquireDesktopInstance({ app: second }), true);
  assert.equal(first.userData, join(root, "desktop"));
  assert.equal(first.userData, second.userData);
});

test("existing desktop profile permissions are restricted without changing the parent", (t) => {
  const root = fixture(t);
  const stateDir = join(root, "state");
  mkdirSync(stateDir, { mode: 0o755 });
  mkdirSync(join(stateDir, "desktop"), { mode: 0o755 });
  const parentMode = statSync(stateDir).mode & 0o777;
  assert.equal(acquireDesktopInstance({ app: new AppStub(), stateDir }), true);
  assert.equal(statSync(join(stateDir, "desktop")).mode & 0o777, 0o700);
  assert.equal(statSync(stateDir).mode & 0o777, parentMode);
});

test("filesystem failure throws a controlled error before lock or listener registration", (t) => {
  const root = fixture(t);
  const stateDir = join(root, "not-a-directory");
  writeFileSync(stateDir, "synthetic");
  const app = new AppStub();
  assert.throws(() => acquireDesktopInstance({ app, stateDir }), {
    code: "DESKTOP_SINGLE_INSTANCE_FAILED",
  });
  assert.deepEqual(app.calls, []);
});

test("a symlinked desktop profile fails closed without chmod of its target", (t) => {
  const root = fixture(t);
  const target = join(root, "unrelated");
  mkdirSync(target, { mode: 0o755 });
  const originalMode = statSync(target).mode & 0o777;
  symlinkSync(target, join(root, "desktop"), "dir");
  const app = new AppStub();
  assert.throws(() => acquireDesktopInstance({ app, stateDir: root }), {
    code: "DESKTOP_SINGLE_INSTANCE_FAILED",
  });
  assert.deepEqual(app.calls, []);
  assert.equal(statSync(target).mode & 0o777, originalMode);
});

test("non-boolean success values fail closed", (t) => {
  const root = fixture(t);
  for (const lockResult of [undefined, 1, "true"]) {
    const app = new AppStub();
    app.requestSingleInstanceLock = () => {
      app.calls.push(["lock"]);
      return lockResult;
    };
    assert.equal(acquireDesktopInstance({ app, stateDir: root }), false);
    assert.equal(app.listenerCount("second-instance"), 0);
    assert.deepEqual(app.calls.at(-1), ["quit"]);
  }
});

test("lock exceptions are controlled failures and cannot register a listener", (t) => {
  const app = new AppStub();
  const original = new Error("synthetic lock failure");
  app.requestSingleInstanceLock = () => {
    throw original;
  };
  assert.throws(
    () => acquireDesktopInstance({ app, stateDir: fixture(t) }),
    (error) => {
      assert.equal(error.code, "DESKTOP_SINGLE_INSTANCE_FAILED");
      assert.equal(error.cause, original);
      return true;
    },
  );
  assert.equal(app.listenerCount("second-instance"), 0);
  assert.equal(
    app.calls.some(([name]) => name === "quit"),
    false,
  );
});

test("failed listener registration releases ownership and removes a partially added listener", (t) => {
  const app = new AppStub();
  app.on = function (name, callback) {
    EventEmitter.prototype.on.call(this, name, callback);
    throw new Error("synthetic registration failure");
  };
  assert.throws(() => acquireDesktopInstance({ app, stateDir: fixture(t) }), {
    code: "DESKTOP_SINGLE_INSTANCE_FAILED",
  });
  assert.equal(app.listenerCount("second-instance"), 0);
  assert.deepEqual(app.calls.at(-1), ["release"]);
});

test("invalid callback fails before taking ownership", (t) => {
  const app = new AppStub();
  assert.throws(
    () => acquireDesktopInstance({ app, stateDir: fixture(t), onSecondInstance: null }),
    { code: "DESKTOP_SINGLE_INSTANCE_FAILED" },
  );
  assert.deepEqual(app.calls, []);
});

test("invalid stateDir does not fall back to the working directory", () => {
  for (const stateDir of ["", "   ", null, 1]) {
    const app = new AppStub();
    assert.throws(() => acquireDesktopInstance({ app, stateDir }), {
      code: "DESKTOP_SINGLE_INSTANCE_FAILED",
    });
    assert.deepEqual(app.calls, []);
  }
});

test("a duplicate whose quit throws falls back to exit without registering listeners", (t) => {
  const app = new AppStub();
  app.lockResult = false;
  app.quit = function () {
    this.calls.push(["quit"]);
    throw new Error("synthetic quit failure");
  };
  assert.equal(acquireDesktopInstance({ app, stateDir: fixture(t) }), false);
  assert.deepEqual(app.calls.slice(-2), [["quit"], ["exit", 1]]);
  assert.equal(app.listenerCount("second-instance"), 0);
});

test("importing the helper does not initialize state or require Electron", (t) => {
  const stateDir = join(fixture(t), "not-created");
  const moduleUrl = new URL("../app/single-instance.mjs", import.meta.url).href;
  const result = spawnSync(
    process.execPath,
    ["--input-type=module", "-e", "await import(process.env.DRAFTSPECT_HELPER_MODULE)"],
    {
      env: { DRAFTSPECT_STATE_DIR: stateDir, DRAFTSPECT_HELPER_MODULE: moduleUrl },
      encoding: "utf8",
      timeout: 5000,
    },
  );
  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.stdout, "");
  assert.equal(existsSync(stateDir), false);
});
