import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { execFileSync } from "node:child_process";
import { createDesktopLifecycle } from "../app/desktop-lifecycle.mjs";

const token = "7ae08c41-c561-4da1-9f01-217fa43eb9c0";
test("heartbeat advances only private metadata, without journal growth, and stops on Quit", () => {
  const f = fixture();
  const life = f.create();
  try {
    life.publish("desktop-visible");
    const journal = fs.readFileSync(join(f.root, "desktop", "desktop-events.jsonl"), "utf8");
    life.heartbeat();
    life.heartbeat();
    assert.equal(f.read().heartbeat, 2);
    assert.equal(fs.readFileSync(join(f.root, "desktop", "desktop-events.jsonl"), "utf8"), journal);
    life.publish("quitting");
    life.heartbeat();
    assert.equal(f.read().heartbeat, 2);
  } finally {
    life.close();
    f.cleanup();
  }
});
function fixture() {
  const root = fs.realpathSync(fs.mkdtempSync(join(tmpdir(), "draftspect-desktop-lifecycle-")));
  fs.chmodSync(root, 0o700);
  fs.mkdirSync(join(root, "desktop"), { mode: 0o700 });
  const path = join(root, "status.json");
  fs.writeFileSync(
    path,
    JSON.stringify({ schemaVersion: 1, token, phase: "launcher-started", pid: 0 }),
    { mode: 0o600 },
  );
  const env = {
    DRAFTSPECT_LAUNCH_STATUS_PATH: path,
    DRAFTSPECT_LAUNCH_TOKEN: token,
    DRAFTSPECT_LAUNCH_PARENT_PID: String(process.ppid),
    SECRET: "SECRET_ENV_VALUE",
  };
  const create = (overrides = {}) => createDesktopLifecycle({ env, stateDir: root, ...overrides });
  return {
    root,
    path,
    env,
    create,
    read: () => JSON.parse(fs.readFileSync(path, "utf8")),
    cleanup: () => fs.rmSync(root, { recursive: true, force: true }),
  };
}

test("private launcher handshake separates visible/service ready and preserves failure across quit", () => {
  const f = fixture();
  let lifecycle;
  try {
    const original = fs.statSync(f.path);
    lifecycle = f.create();
    lifecycle.publish("js-started");
    assert.equal(f.read().pid, process.pid);
    lifecycle.publish("service-ready");
    assert.equal(f.read().serviceReady, true);
    assert.equal(f.read().desktopVisible, false, "service readiness cannot fake a visible window");
    lifecycle.publish("desktop-visible");
    assert.equal(f.read().desktopVisible, true);
    lifecycle.publish("startup-failed", {
      code: "DESKTOP_DAEMON_START_TIMEOUT",
      serviceReady: false,
    });
    lifecycle.publish("quitting");
    assert.equal(f.read().failureCode, "DESKTOP_DAEMON_START_TIMEOUT");
    assert.equal(f.read().serviceReady, false);
    assert.equal(f.read().token, token);
    assert.equal(
      fs.statSync(f.path).ino,
      original.ino,
      "update must retain the private fd identity",
    );
  } finally {
    lifecycle?.close();
    f.cleanup();
  }
});

test("local diagnostic journal admits only safe metadata and never emits token/env/path/raw exception", () => {
  const f = fixture();
  let lifecycle;
  try {
    lifecycle = f.create();
    lifecycle.record("SECRET_EVENT", {
      code: "SECRET_CODE",
      reason: "crashed",
      processType: "renderer",
      exitCode: 9,
      attempt: 1,
      stack: "SECRET_STACK",
      document: "SECRET_DOCUMENT",
      token,
      path: f.path,
      env: f.env,
    });
    const raw = fs.readFileSync(join(f.root, "desktop", "desktop-events.jsonl"), "utf8");
    const value = JSON.parse(raw);
    assert.equal(value.event, "incident");
    assert.equal(value.code, "DESKTOP_START_FAILED");
    assert.equal(value.reason, "crashed");
    assert.equal(value.processType, "renderer");
    assert.doesNotMatch(raw, /SECRET|7ae08c41|status\.json|stack|document|\benv\b/);
    assert.equal(fs.statSync(join(f.root, "desktop", "desktop-events.jsonl")).mode & 0o777, 0o600);
  } finally {
    lifecycle?.close();
    f.cleanup();
  }
});

test("invalid launch authority, shared modes, symlinks and token substitution are rejected without raw input", () => {
  for (const variant of [
    "wrong-parent",
    "partial",
    "token",
    "file-mode",
    "parent-mode",
    "symlink",
  ]) {
    const f = fixture();
    try {
      if (variant === "wrong-parent") f.env.DRAFTSPECT_LAUNCH_PARENT_PID = "2";
      if (variant === "partial") delete f.env.DRAFTSPECT_LAUNCH_TOKEN;
      if (variant === "token")
        f.env.DRAFTSPECT_LAUNCH_TOKEN = "a7247818-fd01-4ee5-a30c-c7316b08e1b9";
      if (variant === "file-mode") fs.chmodSync(f.path, 0o644);
      if (variant === "parent-mode") fs.chmodSync(f.root, 0o755);
      if (variant === "symlink") {
        fs.renameSync(f.path, join(f.root, "target.json"));
        fs.symlinkSync(join(f.root, "target.json"), f.path);
      }
      assert.throws(
        f.create,
        (error) =>
          error.code === "DESKTOP_LAUNCH_CHANNEL_INVALID" && !error.message.includes(f.path),
        variant,
      );
    } finally {
      f.cleanup();
    }
  }
});

test("status replacement or journal symlink cannot redirect later writes to another file", () => {
  const f = fixture();
  let lifecycle;
  try {
    const unrelated = join(f.root, "unrelated.json");
    fs.writeFileSync(unrelated, "KEEP", { mode: 0o600 });
    fs.symlinkSync(unrelated, join(f.root, "desktop", "desktop-events.jsonl"));
    lifecycle = f.create();
    assert.equal(lifecycle.diagnosticAvailable, false);
    lifecycle.publish("js-started");
    fs.renameSync(f.path, join(f.root, "owned-old.json"));
    fs.symlinkSync(unrelated, f.path);
    assert.throws(() => lifecycle.publish("desktop-visible"), {
      code: "DESKTOP_LAUNCH_CHANNEL_INVALID",
    });
    assert.equal(fs.readFileSync(unrelated, "utf8"), "KEEP");
  } finally {
    lifecycle?.close();
    f.cleanup();
  }
});

test("developer entry without launcher channel remains metadata-only and needs no global service", () => {
  const lifecycle = createDesktopLifecycle({ env: {} });
  assert.equal(lifecycle.publish("js-started").desktopVisible, false);
  lifecycle.close();
  lifecycle.close();
});

test("FIFO launch-channel and journal substitutions are rejected without blocking the main thread", () => {
  for (const journal of [false, true]) {
    const f = fixture();
    try {
      const path = journal ? join(f.root, "desktop", "desktop-events.jsonl") : f.path;
      if (!journal) fs.renameSync(f.path, join(f.root, "original-status.json"));
      execFileSync("/usr/bin/mkfifo", [path], { timeout: 1000 });
      const entry = `import {createDesktopLifecycle} from ${JSON.stringify(new URL("../app/desktop-lifecycle.mjs", import.meta.url).href)};
try { const lifecycle=createDesktopLifecycle({env:{...${JSON.stringify(f.env)},DRAFTSPECT_LAUNCH_PARENT_PID:String(process.ppid)},stateDir:${JSON.stringify(f.root)}});
console.log(JSON.stringify({diagnosticAvailable:lifecycle.diagnosticAvailable})); lifecycle.close(); }
catch(error){ console.log(JSON.stringify({code:error.code})); }`;
      const raw = execFileSync(process.execPath, ["--input-type=module", "-e", entry], {
        timeout: 1500,
        encoding: "utf8",
      });
      const result = JSON.parse(raw);
      assert.deepEqual(
        result,
        journal ? { diagnosticAvailable: false } : { code: "DESKTOP_LAUNCH_CHANNEL_INVALID" },
      );
    } finally {
      f.cleanup();
    }
  }
});
