// Opt-in native component check, intentionally outside *.test.mjs. No windows,
// production main/daemon, Word, provider, credentials or normal state are used.
// Run: node tests/single-instance-electron-e2e.mjs [--output /absolute/report.json]
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, mkdir, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { homedir, tmpdir } from "node:os";
import { dirname, isAbsolute, join } from "node:path";
import { fileURLToPath } from "node:url";
import { setTimeout as pause } from "node:timers/promises";

const args = process.argv.slice(2);
if (args.length && (args.length !== 2 || args[0] !== "--output" || !isAbsolute(args[1]))) {
  throw new Error("Use --output with one absolute report path, or no arguments");
}
const project = fileURLToPath(new URL("../", import.meta.url));
const helper = new URL("../app/single-instance.mjs", import.meta.url).href;
const electron = join(project, "node_modules/electron/dist/Electron.app/Contents/MacOS/Electron");
const root = await realpath(await mkdtemp(join(tmpdir(), "draftspect-electron-lock-")));
const entry = join(root, "lock-only.mjs");
const children = [];
const results = [];
let failure;
let cleanupComplete = false;

function launch(name, state) {
  const record = { name, events: [], stderr: "", exited: false, code: null, signal: null };
  const child = spawn(electron, [entry], {
    cwd: root,
    env: {
      HOME: homedir(),
      TMPDIR: join(root, "tmp"),
      PATH: "/usr/bin:/bin",
      LANG: "en_US.UTF-8",
      DRAFTSPECT_STATE_DIR: join(root, state),
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  record.child = child;
  children.push(record);
  let pending = "";
  child.stdout.on("data", (chunk) => {
    pending += String(chunk);
    const lines = pending.split("\n");
    pending = lines.pop();
    for (const line of lines) {
      try {
        record.events.push(JSON.parse(line));
      } catch {
        // Electron diagnostics are not part of the test event protocol.
      }
    }
  });
  child.stderr.on("data", (chunk) => {
    record.stderr = (record.stderr + String(chunk)).slice(-6000);
  });
  child.on("error", (error) => {
    record.spawnError = error.message;
  });
  child.on("exit", (code, signal) => Object.assign(record, { exited: true, code, signal }));
  return record;
}

async function until(predicate, label, record, timeout = 10000) {
  const deadline = Date.now() + timeout;
  while (!predicate()) {
    if (record?.spawnError) throw new Error(`${label}: ${record.spawnError}`);
    if (Date.now() >= deadline) throw new Error(`Timed out: ${label}`);
    await pause(25);
  }
}
const event = (record, type) => record.events.find((item) => item.event === type);

async function owner(record) {
  await until(() => event(record, "ready") || record.exited, `${record.name} ready`, record);
  assert.ok(!record.exited, `${record.name} exited before ready: ${record.stderr}`);
  assert.equal(event(record, "acquired")?.owner, true);
}

async function stop(record) {
  if (!record.exited) {
    assert.equal(record.child.kill("SIGTERM"), true);
    await until(() => record.exited, `${record.name} exit`, record, 5000);
  }
  assert.equal(record.code, 0, record.stderr);
  assert.equal(record.signal, null);
}

try {
  await mkdir(join(root, "tmp"));
  await writeFile(
    entry,
    `import { app } from "electron";
import { acquireDesktopInstance } from ${JSON.stringify(helper)};
const emit = (event, values = {}) => console.log(JSON.stringify({ event, ...values }));
app.disableHardwareAcceleration();
const owner = acquireDesktopInstance({ app,
  onSecondInstance: () => emit("second-instance") });
emit("acquired", { owner, pid: process.pid, userData: app.getPath("userData") });
if (owner) {
  process.on("SIGTERM", () => app.quit());
  app.on("quit", () => emit("quit"));
  app.whenReady().then(() => { app.dock?.hide(); emit("ready"); });
  setTimeout(() => { emit("safety-timeout"); app.exit(90); }, 30000);
}
`,
  );
  const first = launch("first-owner", "state-a");
  await owner(first);
  const duplicate = launch("duplicate", "state-a");
  await until(() => duplicate.exited, "duplicate exit", duplicate);
  assert.equal(duplicate.code, 0, duplicate.stderr);
  assert.equal(duplicate.signal, null);
  assert.equal(event(duplicate, "acquired")?.owner, false);
  assert.equal(event(duplicate, "ready"), undefined);
  await until(() => event(first, "second-instance"), "owner second-instance callback", first);
  results.push({ check: "same-state duplicate refused and owner notified", status: "PASS" });

  const isolated = launch("isolated-owner", "state-b");
  await owner(isolated);
  assert.notEqual(event(first, "acquired").userData, event(isolated, "acquired").userData);
  assert.equal(first.exited, false);
  results.push({ check: "different state directories permit isolated owners", status: "PASS" });

  await stop(first);
  const replacement = launch("replacement-owner", "state-a");
  await owner(replacement);
  assert.equal(event(first, "acquired").userData, event(replacement, "acquired").userData);
  results.push({ check: "owner exits cleanly and replacement reacquires lock", status: "PASS" });
  await stop(replacement);
  await stop(isolated);
} catch (error) {
  failure = error.message;
} finally {
  for (const record of children) {
    if (!record.exited && record.child.pid) record.child.kill("SIGTERM");
  }
  await until(
    () => children.every((record) => record.exited || !record.child.pid),
    "cleanup",
    null,
    5000,
  ).catch(async () => {
    // Only exact child handles created above, never process-name/global kill.
    for (const record of children) {
      if (!record.exited && record.child.pid) record.child.kill("SIGKILL");
    }
    await until(
      () => children.every((record) => record.exited || !record.child.pid),
      "forced cleanup",
      null,
      5000,
    );
  });
  cleanupComplete = children.every((record) => record.exited || !record.child.pid);
  if (cleanupComplete) await rm(root, { recursive: true, force: true });
}

const report = {
  timestamp: new Date().toISOString(),
  status: failure ? "FAIL" : "PASS",
  electron,
  electronVersion: (
    await readFile(join(project, "node_modules/electron/package.json"), "utf8").then(JSON.parse)
  ).version,
  scope:
    "native Electron single-instance helper only; no production main, daemon, windows, Word or provider",
  results,
  failure,
  cleanupComplete,
  children: children.map(({ child, ...record }) => ({ ...record, pid: child.pid })),
};
if (args.length) {
  await mkdir(dirname(args[1]), { recursive: true });
  await writeFile(args[1], `${JSON.stringify(report, null, 2)}\n`);
}
console.log(JSON.stringify(report, null, 2));
if (failure) process.exitCode = 1;
