// Opt-in native Electron acceptance, intentionally outside *.test.mjs.
// Actual production main/status-window code; isolated copied project + model0
// daemon/preflight/registration fixtures. No provider, Office, installed app,
// security-option changes, system trust, global configuration or model request.
// Run: node tests/desktop-fault-electron-e2e.mjs --output /absolute/report.json
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { cp, mkdir, mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { dirname, isAbsolute, join } from "node:path";
import { tmpdir } from "node:os";
import { fileURLToPath } from "node:url";
import { setTimeout as pause } from "node:timers/promises";

const args = process.argv.slice(2);
if (args.length !== 2 || args[0] !== "--output" || !isAbsolute(args[1]))
  throw new Error("Use --output with one absolute new report path");
const output = args[1],
  project = fileURLToPath(new URL("../", import.meta.url));
const electron = join(project, "node_modules/electron/dist/Electron.app/Contents/MacOS/Electron");
const root = await realpath(await mkdtemp(join(tmpdir(), "draftspect-native-desktop-fault-")));
const fixture = join(root, "project"),
  state = join(root, "state");
const children = [],
  checks = [],
  sources = {},
  pendingRequests = new Map();
const sha = (value) => createHash("sha256").update(value).digest("hex");
let sequence = 0,
  failure = null,
  cleanupComplete = false;
const safeStatus = (value) =>
  value && {
    schemaVersion: value.schemaVersion,
    pid: value.pid,
    phase: value.phase,
    desktopVisible: value.desktopVisible,
    serviceReady: value.serviceReady,
    code: value.code,
    failureCode: value.failureCode,
  };
async function status(record) {
  try {
    return safeStatus(JSON.parse(await readFile(record.statusPath, "utf8")));
  } catch {
    return null;
  }
}
async function until(predicate, label, record, timeoutMs = 15_000) {
  const deadline = Date.now() + timeoutMs;
  while (!(await predicate())) {
    if (record?.exited)
      throw new Error(
        `${label}: owned Electron exited (code=${record.code}, signal=${record.signal})`,
      );
    if (Date.now() >= deadline) throw new Error(`${label}: bounded timeout`);
    await pause(25);
  }
}
async function request(record, action, fields = {}) {
  const id = ++sequence;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingRequests.delete(id);
      reject(new Error(`Native fixture action timed out: ${action}`));
    }, 5000);
    pendingRequests.set(id, { resolve, reject, timer });
    try {
      record.child.send({ id, action, ...fields }, (error) => {
        if (error) {
          clearTimeout(timer);
          pendingRequests.delete(id);
          reject(new Error("Native fixture IPC unavailable"));
        }
      });
    } catch {
      clearTimeout(timer);
      pendingRequests.delete(id);
      reject(new Error("Native fixture IPC unavailable"));
    }
  });
}
async function launch(name, stateDir = state) {
  const attempt = join(root, name),
    token = randomUUID();
  await mkdir(attempt, { mode: 0o700 });
  const statusPath = join(attempt, "status.json");
  await writeFile(
    statusPath,
    JSON.stringify({ schemaVersion: 1, token, pid: 0, phase: "launcher-started" }),
    { flag: "wx", mode: 0o600 },
  );
  const record = {
    name,
    token,
    statusPath,
    exited: false,
    code: null,
    signal: null,
    events: [],
    stdout: "",
    stderr: "",
    child: null,
  };
  const env = {
    HOME: process.env.HOME,
    TMPDIR: root,
    PATH: "/usr/bin:/bin",
    LANG: "en_US.UTF-8",
    DRAFTSPECT_STATE_DIR: stateDir,
    DRAFTSPECT_DESKTOP_LAUNCH: "1",
    DRAFTSPECT_RUNTIME: "codebuddy",
    DRAFTSPECT_LAUNCH_STATUS_PATH: statusPath,
    DRAFTSPECT_LAUNCH_TOKEN: token,
    DRAFTSPECT_LAUNCH_PARENT_PID: String(process.pid),
  };
  record.child = spawn(electron, [join(fixture, "app", "native-fixture.mjs")], {
    cwd: fixture,
    env,
    stdio: ["ignore", "pipe", "pipe", "ipc"],
  });
  children.push(record);
  record.child.stdout.on("data", (value) => {
    record.stdout += String(value);
  });
  record.child.stderr.on("data", (value) => {
    record.stderr += String(value);
  });
  record.child.on("message", (message) => {
    if (message?.event === "response") {
      const pending = pendingRequests.get(message.id);
      if (!pending) return;
      clearTimeout(pending.timer);
      pendingRequests.delete(message.id);
      if (message.ok) pending.resolve(message.data);
      else pending.reject(new Error("Native controlled fixture action failed"));
    } else if (message?.event) record.events.push(message);
  });
  record.child.once("error", () => {
    record.spawnError = true;
  });
  record.child.once("exit", (code, signal) => {
    Object.assign(record, { exited: true, code, signal });
  });
  return record;
}
async function waitExited(record) {
  await until(() => record.exited, `${record.name} exit`, null, 6000);
}
async function stop(record) {
  if (!record.exited) await request(record, "quit");
  await waitExited(record);
  assert.equal(record.code, 0);
  assert.equal(record.signal, null);
  assert.equal((await status(record)).phase, "quitting");
}
async function snapshot(record) {
  return request(record, "snapshot");
}
async function recordCheck(name, execute) {
  const beganAt = new Date().toISOString();
  const detail = await execute();
  checks.push({ name, status: "PASS", beganAt, endedAt: new Date().toISOString(), detail });
}
async function captureJournal(record) {
  try {
    return (await readFile(join(state, "desktop", "desktop-events.jsonl"), "utf8"))
      .trim()
      .split("\n")
      .map(JSON.parse)
      .filter((value) => value.pid === record.child.pid);
  } catch {
    return [];
  }
}

try {
  await mkdir(fixture, { mode: 0o700 });
  await mkdir(state, { mode: 0o700 });
  await cp(join(project, "app"), join(fixture, "app"), { recursive: true, errorOnExist: true });
  await mkdir(join(fixture, "daemon"), { mode: 0o700 });
  for (const name of ["system-paths.mjs", "state-paths.mjs", "log.mjs", "local-transport.mjs"])
    await cp(join(project, "daemon", name), join(fixture, "daemon", name), { errorOnExist: true });
  const source = await readFile(join(project, "app", "main.mjs"), "utf8");
  for (const name of [
    "app/main.mjs",
    "app/desktop-status-window.mjs",
    "app/desktop-lifecycle.mjs",
    "app/single-instance.mjs",
    "app/quit-barrier.mjs",
    "app/restart-backoff.mjs",
  ])
    sources[name] = sha(await readFile(join(project, name)));
  await writeFile(
    join(fixture, "app", "main.mjs"),
    source +
      "\nexport { daemonProcess, daemonStatus, restartAttempts, desktopStartupError, appIsQuitting, ownsDesktopInstance, startDaemon, stopDaemon, restartDaemon };\n",
    { mode: 0o600 },
  );
  await writeFile(
    join(fixture, "app", "desktop-preflight.mjs"),
    `export async function desktopPreflight(){return {cliPath:${JSON.stringify(process.execPath)},nodePath:${JSON.stringify(process.execPath)},warnings:[]};}\n`,
    { mode: 0o600 },
  );
  await writeFile(
    join(fixture, "app", "sideload.mjs"),
    "export async function isAddinInstalled(){return true;} export async function installAddin(){throw new Error('FIXTURE_INSTALL_DENIED');} export async function uninstallAddin(){throw new Error('FIXTURE_INSTALL_DENIED');}\n",
    { mode: 0o600 },
  );
  const registration = new URL("../app/addin-registration.mjs", import.meta.url).href;
  await writeFile(
    join(fixture, "app", "addin-registration.mjs"),
    `export {officeConnectionDetail,inspectManifestRegistration} from ${JSON.stringify(registration)}; export async function readAddinRegistration(){return ['word','excel'].map(host=>({host,state:'matched',label:'synthetic fixture registration; no Office access'}));}\n`,
    { mode: 0o600 },
  );
  await writeFile(
    join(fixture, "daemon", "index.mjs"),
    `// Controlled no-model IPC daemon fixture; no listeners, SDK or provider imports.
process.on('message', msg=>{if(msg?.type==='health_check')process.send?.({type:'health_result',id:msg.id,phase:'ready',connected_panes_by_host:{word:0,excel:0},unresponsive_active_panes:0});});
process.on('SIGTERM',()=>process.exit(0)); process.on('disconnect',()=>process.exit(0));
process.send?.({type:'daemon_ready'}); setInterval(()=>{},1000);\n`,
    { mode: 0o600 },
  );
  await writeFile(
    join(fixture, "app", "native-fixture.mjs"),
    `import {app,BrowserWindow} from 'electron';
const main=await import('./main.mjs');
const send=value=>{try{process.send?.(value);}catch{}};
const win=()=>BrowserWindow.getAllWindows().find(window=>!window.isDestroyed());
process.on('message',async msg=>{
try{let data=null;
if(msg.action==='snapshot'){const window=win();data={windowCount:BrowserWindow.getAllWindows().length,windowVisible:window?.isVisible()===true,rendererPid:window?.webContents.getOSProcessId()||null,daemonPid:main.daemonProcess?.pid||null,daemonStatus:main.daemonStatus,restartAttempts:main.restartAttempts,errorCode:(main.desktopStartupError||'').match(/^DESKTOP_[A-Z_]+/)?.[0]||null};}
else if(msg.action==='crash-renderer'){win().webContents.forcefullyCrashRenderer();}
else if(msg.action==='hide'){win().close();}
else if(msg.action==='show'){app.emit('second-instance',{},[],process.cwd());}
else if(msg.action==='crash-daemon'){main.daemonProcess.kill('SIGTERM');}
else if(msg.action==='stop'){main.stopDaemon();}
else if(msg.action==='start'){await main.startDaemon();}
else if(msg.action==='quit'){send({event:'response',id:msg.id,ok:true,data:{requested:true}});app.quit();return;}
else throw new Error('FIXTURE_ACTION_DENIED');
send({event:'response',id:msg.id,ok:true,data});
}catch{send({event:'response',id:msg.id,ok:false});}
});
send({event:'fixture-loaded',owner:main.ownsDesktopInstance,pid:process.pid});\n`,
    { flag: "wx", mode: 0o600 },
  );
  const owner = await launch("owner");
  await recordCheck("A production main/window visible and structured readiness", async () => {
    await until(
      async () => {
        const value = await status(owner);
        return value?.desktopVisible === true && value.serviceReady === true;
      },
      "native visible+fixture daemon readiness",
      owner,
    );
    const value = await snapshot(owner);
    assert.equal(value.windowVisible, true);
    assert.equal(value.daemonStatus, "running");
    assert.ok(value.daemonPid > 0);
    return {
      ...value,
      handshake: await status(owner),
      serviceReadinessSource: "controlled model0 IPC fixture (not real listeners/provider)",
    };
  });
  await recordCheck(
    "F real same-state duplicate refused and forwarded to existing visible owner",
    async () => {
      const duplicate = await launch("duplicate");
      await waitExited(duplicate);
      assert.equal(duplicate.code, 0);
      assert.equal(duplicate.signal, null);
      assert.equal((await status(duplicate)).phase, "duplicate");
      const value = await snapshot(owner);
      assert.equal(value.windowCount, 1);
      assert.equal(value.windowVisible, true);
      return { duplicate: await status(duplicate), owner: value };
    },
  );
  await recordCheck(
    "D/E actual renderer death recovers three times then stops at finite budget",
    async () => {
      const results = [];
      for (const expectedAttempts of [1, 2, 3]) {
        const old = await snapshot(owner);
        await request(owner, "crash-renderer");
        await until(
          async () => {
            const value = await snapshot(owner);
            return (
              value.windowVisible &&
              value.rendererPid &&
              value.rendererPid !== old.rendererPid &&
              value.errorCode === null
            );
          },
          `renderer recovery ${expectedAttempts}`,
          owner,
          9000,
        );
        const current = await snapshot(owner);
        assert.equal(current.daemonPid, old.daemonPid);
        assert.equal(current.daemonStatus, "running");
        results.push(current);
      }
      await request(owner, "crash-renderer");
      await until(
        async () => (await snapshot(owner)).errorCode === "DESKTOP_RENDERER_RECOVERY_EXHAUSTED",
        "renderer finite cap",
        owner,
      );
      assert.equal((await snapshot(owner)).daemonStatus, "running");
      return { replacements: results.length, journal: await captureJournal(owner) };
    },
  );
  await recordCheck("explicit manual reopen can recover exhausted dead renderer", async () => {
    const old = await snapshot(owner);
    await request(owner, "show");
    await until(
      async () => {
        const value = await snapshot(owner);
        return (
          value.windowVisible &&
          value.rendererPid &&
          value.rendererPid !== old.rendererPid &&
          value.errorCode === null
        );
      },
      "manual renderer reopen",
      owner,
    );
    return snapshot(owner);
  });
  await recordCheck("hidden dead renderer stays hidden until explicit reopen", async () => {
    await request(owner, "hide");
    const old = await snapshot(owner);
    assert.equal(old.windowVisible, false);
    await request(owner, "crash-renderer");
    await pause(250);
    assert.equal((await snapshot(owner)).windowVisible, false);
    await request(owner, "show");
    await until(
      async () => {
        const value = await snapshot(owner);
        return value.windowVisible && value.rendererPid && value.rendererPid !== old.rendererPid;
      },
      "hidden renderer explicit reopen",
      owner,
    );
    return snapshot(owner);
  });
  await recordCheck(
    "F daemon crash backoff is revoked by manual Stop before its timer can restart",
    async () => {
      await request(owner, "crash-daemon");
      await until(
        async () => (await snapshot(owner)).daemonStatus === "crashed",
        "daemon crash recognition",
        owner,
      );
      await request(owner, "stop");
      await pause(1250);
      const value = await snapshot(owner);
      assert.equal(value.daemonPid, null);
      assert.equal(value.daemonStatus, "stopped");
      assert.equal((await status(owner)).serviceReady, false);
      await request(owner, "start");
      await until(
        async () => (await snapshot(owner)).daemonStatus === "running",
        "explicit daemon start",
        owner,
      );
      return { stopped: value, restarted: await snapshot(owner) };
    },
  );
  await recordCheck("A normal explicit Quit clears exact owned daemon and process", async () => {
    const value = await snapshot(owner);
    await stop(owner);
    return { before: value, final: await status(owner) };
  });
  const terminated = await launch("main-terminated");
  await until(
    async () =>
      (await status(terminated))?.desktopVisible === true &&
      (await snapshot(terminated)).daemonStatus === "running",
    "second isolated main visible",
    terminated,
  );
  await recordCheck(
    "D actual Electron main SIGTERM follows native graceful Quit without fabricated crash",
    async () => {
      const before = await snapshot(terminated);
      terminated.child.kill("SIGTERM");
      await waitExited(terminated);
      assert.equal(terminated.code, 0);
      assert.equal(terminated.signal, null);
      assert.equal((await status(terminated)).phase, "quitting");
      // Own fixture daemon exits on genuine inherited IPC disconnect.
      await until(
        () => {
          try {
            process.kill(before.daemonPid, 0);
            return false;
          } catch (error) {
            return error.code === "ESRCH";
          }
        },
        "owned fixture daemon released after main death",
        null,
        5000,
      );
      return {
        before,
        exitCode: terminated.code,
        signal: terminated.signal,
        lastHandshake: await status(terminated),
      };
    },
  );
  const killed = await launch("main-killed");
  await until(
    async () =>
      (await status(killed))?.desktopVisible === true &&
      (await snapshot(killed)).daemonStatus === "running",
    "third isolated main visible",
    killed,
  );
  await recordCheck(
    "D sudden Electron main SIGKILL is observed outside JS and releases fixture daemon through IPC",
    async () => {
      const before = await snapshot(killed);
      killed.child.kill("SIGKILL");
      await waitExited(killed);
      assert.equal(killed.signal, "SIGKILL");
      assert.notEqual((await status(killed)).phase, "quitting");
      await until(
        () => {
          try {
            process.kill(before.daemonPid, 0);
            return false;
          } catch (error) {
            return error.code === "ESRCH";
          }
        },
        "owned fixture daemon released after sudden main death",
        null,
        5000,
      );
      return {
        before,
        exitCode: killed.code,
        signal: killed.signal,
        lastHandshake: await status(killed),
        mainAutomaticRestarts: 0,
      };
    },
  );
} catch (error) {
  failure = error.message;
} finally {
  for (const record of children)
    if (!record.exited && record.child.pid) {
      try {
        await request(record, "quit");
      } catch {
        record.child.kill("SIGTERM");
      }
    }
  try {
    await until(
      () => children.every((record) => record.exited || !record.child.pid),
      "owned native fixture cleanup",
      null,
      6000,
    );
  } catch {
    for (const record of children)
      if (!record.exited && record.child.pid) record.child.kill("SIGKILL");
    await until(
      () => children.every((record) => record.exited || !record.child.pid),
      "forced exact owned native fixture cleanup",
      null,
      6000,
    );
  }
  cleanupComplete = children.every((record) => record.exited || !record.child.pid);
  for (const pending of pendingRequests.values()) {
    clearTimeout(pending.timer);
    pending.reject(new Error("Native fixture closed"));
  }
  pendingRequests.clear();
}
const childEvidence = await Promise.all(
  children.map(async (record) => ({
    name: record.name,
    pid: record.child.pid,
    exited: record.exited,
    code: record.code,
    signal: record.signal,
    events: record.events,
    lastHandshake: await status(record),
    stdout: record.stdout.replaceAll(record.token, "[redacted-launch-token]"),
    stderr: record.stderr.replaceAll(record.token, "[redacted-launch-token]"),
    journal: await captureJournal(record),
  })),
);
const report = {
  schemaVersion: 1,
  at: new Date().toISOString(),
  status: failure ? "FAIL" : "PASS",
  failure,
  scope: {
    actualElectronMain: true,
    actualProductionStatusWindow: true,
    actualMainSourceCopyWithTestExports: true,
    preflightRegistrationDaemon: "CONTROLLED_MODEL0_FIXTURES",
    realSdkProviderOffice: "NOT_RUN",
    stableBudgetReset: "SOURCE_VM_ONLY",
    modelRequests: 0,
    businessDocuments: 0,
    installedAppChanges: 0,
    securityOptionsChanged: 0,
    uploads: 0,
  },
  sources,
  checks,
  cleanupComplete,
  children: childEvidence,
};
await mkdir(dirname(output), { recursive: true, mode: 0o700 });
await writeFile(output, JSON.stringify(report, null, 2) + "\n", { flag: "wx", mode: 0o600 });
if (cleanupComplete) await rm(root, { recursive: true, force: true });
console.log(
  JSON.stringify({
    status: report.status,
    checks: checks.length,
    failure,
    cleanupComplete,
    output,
  }),
);
if (failure || !cleanupComplete) process.exitCode = 1;
