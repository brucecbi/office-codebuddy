// Draftspect for Office — Electron menu bar shell.
//
// Wraps the daemon as a child process, exposes a tray icon with status and
// controls, provides a local desktop status window, and restarts
// the daemon if it crashes. The daemon code itself is untouched.

import {
  app,
  Tray,
  Menu,
  shell,
  dialog,
  nativeImage,
  BrowserWindow,
  Notification,
  powerMonitor,
  ipcMain,
} from "electron";
import { createHealthWatch } from "./health-watch.mjs";
import { spawn } from "node:child_process";
import { readFile, stat } from "node:fs/promises";
import fs from "node:fs";
import { homedir } from "node:os";
import { join, dirname, delimiter } from "node:path";
import { fileURLToPath } from "node:url";
import { isAddinInstalled, installAddin, uninstallAddin } from "./sideload.mjs";
import { openWordBootstrap } from "./word-bootstrap.mjs";
import { openExcelBootstrap } from "./excel-bootstrap.mjs";
import { openInMicrosoftExcel } from "./macos-excel-open.mjs";
import { readAddinRegistration, officeConnectionDetail } from "./addin-registration.mjs";
import {
  addinActivationInstructions,
  addinInstallRecoveryInstructions,
} from "./addin-guidance.mjs";
import { isSystemHomeChild } from "../daemon/system-paths.mjs";
import { stateDirectory, stateFile } from "../daemon/state-paths.mjs";
import { restartDelayMs } from "./restart-backoff.mjs";
import { installTimestampedConsole } from "../daemon/log.mjs";
import { acquireDesktopInstance } from "./single-instance.mjs";
import { desktopPreflight } from "./desktop-preflight.mjs";
import { waitForOwnedExit } from "./quit-barrier.mjs";
import { createDesktopStatusWindow } from "./desktop-status-window.mjs";
import { createDesktopLifecycle } from "./desktop-lifecycle.mjs";

installTimestampedConsole();

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const DAEMON_ENTRY = join(PROJECT_ROOT, "daemon", "index.mjs");
let SESSIONS_FILE;
let LOG_FILE;
let desktopLifecycle = null;
let desktopHeartbeat = null;

// All entry points share one owner for this state directory. A losing process
// must never register startup callbacks, open logs or touch the daemon.
let ownsDesktopInstance = false;
try {
  const stateDir = stateDirectory();
  SESSIONS_FILE = stateFile("sessions.json");
  LOG_FILE = stateFile("daemon.log");
  ownsDesktopInstance = acquireDesktopInstance({
    app,
    stateDir,
    onSecondInstance: () => {
      app
        .whenReady()
        .then(showDesktopStatus)
        .catch(() => {});
    },
  });
  desktopLifecycle = createDesktopLifecycle({
    stateDir: ownsDesktopInstance ? stateDir : undefined,
  });
  desktopLifecycle.publish(ownsDesktopInstance ? "js-started" : "duplicate");
  if (ownsDesktopInstance) {
    desktopHeartbeat = setInterval(() => {
      try {
        desktopLifecycle?.heartbeat();
      } catch {
        clearInterval(desktopHeartbeat);
        desktopIncident({ code: "DESKTOP_LAUNCH_CHANNEL_INVALID", processType: "main" });
      }
    }, 2000);
    desktopHeartbeat.unref?.();
  }
} catch {
  ownsDesktopInstance = false;
  dialog.showErrorBox(
    "Draftspect 无法启动",
    "状态目录或应用锁不可用。请检查启动配置和目录权限，或退出已有的 Draftspect 后重试。未启动本地服务。",
  );
  app.quit();
}

// Don't launch the daemon with an OS-managed $HOME child as the initial
// cwd, even if a stale sessions.json says so. Canonical set lives in
// daemon/system-paths.mjs (isSystemHomeChild).

let tray = null;
let daemonProcess = null;
let logStream = null;
let daemonStatus = "starting"; // starting | running | crashed | stopping | stopped | blocked
let currentWorkspace = null;
let restartAttempts = 0;
const MAX_RESTART = 3;
// The restart counter is only zeroed once the daemon has stayed up for this
// long. Resetting on the first "running" signal would let a daemon that
// crashes seconds after every start restart forever — the cap would never
// bite. See PR "crash-cap + restart backoff".
const STABLE_AFTER_MS = 45_000;
let stableTimer = null;
let daemonHealthTimer = null;
let daemonWatch = null;
let healthIncident = null;
let pendingDaemonStart = null;
let restartTimer = null;
const daemonExitIntent = new WeakMap();
let desktopStartupError = null;
let desktopStatusDialogOpen = false;
let appIsQuitting = false;
let desktopQuitPending = false;
let desktopQuitReady = false;
let desktopWindow = null;
let addinRegistration = [];
let officeHealth = null;
let daemonStartupTimer = null;
let fatalMainError = false;

function publishDesktopPhase(phase, fields = {}) {
  try {
    desktopLifecycle?.publish(phase, fields);
  } catch {
    desktopStartupError ||=
      "DESKTOP_LAUNCH_CHANNEL_INVALID: 启动监控通道不可用，请退出并重新打开桌面入口。";
    desktopLifecycle?.record("incident", {
      code: "DESKTOP_LAUNCH_CHANNEL_INVALID",
      processType: "main",
    });
  }
}

function desktopIncident(fields) {
  desktopLifecycle?.record("incident", fields);
  if (fields.code === "DESKTOP_RENDERER_GONE" || fields.code === "DESKTOP_STATUS_LOAD_FAILED")
    publishDesktopPhase("js-started", { desktopVisible: false });
  if (fields.reason === "responsive") {
    if (desktopStartupError?.startsWith("DESKTOP_RENDERER_UNRESPONSIVE:"))
      desktopStartupError = null;
  } else if (fields.code) {
    desktopStartupError = `${fields.code}: 桌面显示发生异常，本地服务与文档操作的状态需要分别核对；系统不会重发任务。`;
  }
  updateTray();
}

function getDesktopWindow() {
  if (!desktopWindow) {
    desktopWindow = createDesktopStatusWindow({
      BrowserWindow,
      ipcMain,
      app,
      onVisible: () => {
        if (/^DESKTOP_RENDERER_/.test(desktopStartupError || "")) desktopStartupError = null;
        publishDesktopPhase("desktop-visible");
      },
      onIncident: desktopIncident,
      readStatus: () => ({
        status: statusLabel().replace(/^●\s*/, ""),
        runtime: process.env.DRAFTSPECT_RUNTIME || "codebuddy",
        detail:
          (desktopLifecycle && !desktopLifecycle.diagnosticAvailable
            ? "DESKTOP_DIAGNOSTIC_UNAVAILABLE: 本地桌面诊断记录不可写，功能状态仍分别核对；请检查状态目录权限或磁盘空间。\n\n"
            : "") +
          officeConnectionDetail(
            addinRegistration,
            daemonStatus === "running" ? officeHealth : null,
          ) +
          "\n\n" +
          "本地服务、Office 连接与模型可用性分别验证。\n" +
          "普通加载项菜单无法激活时，可初始化合成 Word / Excel 文档；不要绑定业务文档。\n" +
          "系统不会自动重发模型任务或文档写入。\n" +
          "关闭此窗口会隐藏窗口，服务继续运行；再次双击桌面入口可重新打开。",
        error: desktopStartupError || "",
      }),
      onAction: async (action) => {
        if (appIsQuitting && action !== "quit") throw new Error("服务正在退出，请等待完成。");
        if (action === "start" || action === "restart") {
          await (action === "start" ? startDaemon() : restartDaemon());
          if (daemonStatus === "blocked") throw new Error("本地服务启动需要处理。");
          return;
        }
        if (action === "stop") return stopDaemon();
        if (action === "initialize-word") {
          if (!(await runInitializeWordAddin())) throw new Error("Word 初始化未完成。");
          return;
        }
        if (action === "initialize-excel") {
          if (!(await runInitializeExcelAddin())) throw new Error("Excel 初始化未完成。");
          return;
        }
        if (action === "open-logs") {
          if (await shell.openPath(LOG_FILE)) throw new Error("无法打开日志。");
          return;
        }
        if (action === "quit") return app.quit();
        throw new Error("不支持的操作。");
      },
    });
  }
  return desktopWindow;
}

function reportStartupFailure(error) {
  const code = /^DESKTOP_[A-Z_]+$/.test(error?.code || "") ? error.code : "DESKTOP_START_FAILED";
  const detail =
    code === "DESKTOP_START_FAILED"
      ? "本地服务无法启动。请检查项目依赖、执行权限与日志。"
      : error.message;
  desktopStartupError = `${code}: ${detail}`;
  publishDesktopPhase("startup-failed", { code, serviceReady: false });
  daemonStatus = "blocked";
  logStream?.write(`[app] ${desktopStartupError}\n`);
  updateTray();
  return showDesktopStatus().catch(() => {});
}

async function showDesktopStatus() {
  if (appIsQuitting) return;
  if (process.env.DRAFTSPECT_DESKTOP_LAUNCH === "1") {
    try {
      await getDesktopWindow().show();
      return;
    } catch {
      desktopWindow?.dispose();
      desktopWindow = null;
      desktopStartupError ||=
        "DESKTOP_STATUS_LOAD_FAILED: 本地状态窗口无法加载。请检查项目文件或查看日志。";
      logStream?.write("[app] DESKTOP_STATUS_LOAD_FAILED\n");
      // Native fallback must remain reachable even if local HTML/preload is missing.
    }
  }
  if (desktopStatusDialogOpen) return;
  desktopStatusDialogOpen = true;
  try {
    app.focus({ steal: true });
    const { response } = await dialog.showMessageBox({
      type: desktopStartupError ? "error" : "info",
      title: "Draftspect 本地服务",
      message: desktopStartupError ? "启动需要处理" : statusLabel().replace(/^●\s*/, ""),
      detail: [
        desktopStartupError,
        desktopLifecycle && !desktopLifecycle.diagnosticAvailable
          ? "DESKTOP_DIAGNOSTIC_UNAVAILABLE: 本地桌面诊断记录不可写，请检查状态目录权限或磁盘空间。"
          : "",
        `项目：${PROJECT_ROOT}`,
        `运行时：${process.env.DRAFTSPECT_RUNTIME || "codebuddy"}`,
        officeConnectionDetail(addinRegistration, officeHealth),
        "本地服务就绪与 Office 连接、模型可用性分别验证。请在目标文档侧边栏查看连接状态。",
        "系统不会自动重发模型任务或文档写入。",
      ]
        .filter(Boolean)
        .join("\n\n"),
      buttons: ["关闭", "查看日志"],
      defaultId: 0,
      cancelId: 0,
    });
    if (response === 1) await shell.openPath(LOG_FILE);
  } finally {
    desktopStatusDialogOpen = false;
  }
}

function reportHealthIncident(code, message) {
  if (healthIncident === code) return;
  healthIncident = code;
  logStream?.write(`[app] health: ${code}\n`);
  updateTray();
  if (Notification.isSupported())
    new Notification({ title: "Draftspect 需要检查", body: message }).show();
}
if (ownsDesktopInstance)
  app
    .whenReady()
    .then(() => {
      powerMonitor.on("suspend", () => daemonWatch?.suspend());
      powerMonitor.on("resume", () => {
        daemonWatch?.resume();
        healthIncident = null;
        updateTray();
      });
    })
    .catch(reportStartupFailure);

// --------------------------------------------------------------------------
// Daemon lifecycle
// --------------------------------------------------------------------------
async function findInitialWorkspace() {
  try {
    const state = JSON.parse(await readFile(SESSIONS_FILE, "utf8"));
    const folders = Object.entries(state.folders || {}).filter(([cwd]) => !isSystemHomeChild(cwd));
    if (folders.length === 0) return null;
    folders.sort(([, a], [, b]) => (b.last_used || "").localeCompare(a.last_used || ""));
    return folders[0][0];
  } catch {
    return null;
  }
}

function openLogStream() {
  fs.mkdirSync(dirname(LOG_FILE), { recursive: true });
  const stream = fs.createWriteStream(LOG_FILE, { flags: "a", mode: 0o600 });
  logStream = stream;
  stream.on("error", () => {
    if (logStream !== stream) return;
    logStream = null;
    if (appIsQuitting) return;
    stopDaemon();
    reportStartupFailure({
      code: "DESKTOP_LOG_UNAVAILABLE",
      message:
        "日志文件无法写入。已取消启动或请求停止本地服务；请检查磁盘空间和状态目录权限，并核对先前已执行的操作。",
    });
  });
  logStream.write(`\n=== ${new Date().toISOString()} app start ===\n`);
}

async function startDaemon() {
  if (appIsQuitting) return;
  cancelScheduledRestart();
  if (daemonProcess || pendingDaemonStart) return;
  officeHealth = null;
  // Fence asynchronous workspace discovery: Stop or a newer start may replace
  // this request before it resolves. A start owns the slot before its first await.
  const request = {};
  pendingDaemonStart = request;
  // Only the installed desktop entry requires the existing secure CodeBuddy
  // deployment. Developer/isolated test entry points retain their old behavior.
  if (process.env.DRAFTSPECT_DESKTOP_LAUNCH === "1") {
    try {
      const result = await desktopPreflight({ projectRoot: PROJECT_ROOT });
      if (pendingDaemonStart !== request) return;
      process.env.CODEBUDDY_CODE_PATH = result.cliPath;
      process.env.DRAFTSPECT_NODE_PATH = result.nodePath;
      process.env.PATH = [dirname(result.nodePath), process.env.PATH]
        .filter(Boolean)
        .join(delimiter);
      for (const warning of result.warnings) logStream?.write(`[app] ${warning}\n`);
    } catch (error) {
      if (pendingDaemonStart !== request) return;
      pendingDaemonStart = null;
      reportStartupFailure(error);
      return;
    }
  }
  let workspace;
  try {
    workspace = await findInitialWorkspace();
  } catch (error) {
    if (pendingDaemonStart !== request) return;
    pendingDaemonStart = null;
    await reportStartupFailure(error);
    return;
  }
  if (pendingDaemonStart !== request) return;
  pendingDaemonStart = null;
  currentWorkspace = workspace;
  desktopStartupError = null;
  daemonStatus = "starting";
  updateTray();

  const args = [DAEMON_ENTRY];
  if (workspace) args.push(workspace);
  const childEnvironment = { ...process.env, ELECTRON_RUN_AS_NODE: "1" };
  // The launcher's private handshake belongs only to this Electron owner;
  // neither the daemon nor a provider/CLI child needs that authority.
  delete childEnvironment.DRAFTSPECT_LAUNCH_STATUS_PATH;
  delete childEnvironment.DRAFTSPECT_LAUNCH_TOKEN;
  delete childEnvironment.DRAFTSPECT_LAUNCH_PARENT_PID;

  try {
    daemonProcess = spawn(process.execPath, args, {
      cwd: PROJECT_ROOT,
      // fd 3 is an IPC channel for control messages (e.g. the native folder
      // picker). The daemon uses process.send / process.on('message') over this
      // channel; stdout/stderr remain plain log streams.
      stdio: ["ignore", "pipe", "pipe", "ipc"],
      env: childEnvironment,
    });
  } catch (error) {
    reportStartupFailure(error);
    return;
  }

  // Arm the stability timer for this process. If it's still the live daemon
  // after STABLE_AFTER_MS, the start "stuck" and the restart budget is
  // refilled. A crash before then clears the timer (see the exit handler), so
  // a fast crash-loop never refills the budget and the cap eventually bites.
  if (stableTimer) clearTimeout(stableTimer);
  const thisProc = daemonProcess;
  clearTimeout(daemonStartupTimer);
  daemonStartupTimer = setTimeout(() => {
    if (daemonProcess !== thisProc || daemonStatus !== "starting") return;
    daemonExitIntent.set(thisProc, "startup-failed");
    clearTimeout(stableTimer);
    stableTimer = null;
    reportStartupFailure({
      code: "DESKTOP_DAEMON_START_TIMEOUT",
      message:
        "本地服务未在20秒内确认就绪。已请求停止自有服务；请检查状态后手动重试，不会重发文档操作。",
    });
    try {
      thisProc.kill("SIGTERM");
    } catch {
      /* retain owned process and lock until actual exit */
    }
  }, 20_000);
  daemonWatch = createHealthWatch();
  healthIncident = null;
  clearInterval(daemonHealthTimer);
  daemonHealthTimer = setInterval(() => {
    if (daemonProcess !== thisProc) return;
    const reportIpcFailure = () => {
      if (daemonProcess !== thisProc) return;
      reportHealthIncident(
        "DAEMON_UNRESPONSIVE",
        "本地服务的监控通道已断开。请检查托盘日志；系统不会自动重启服务或重试文档写入。",
      );
    };
    if (!thisProc.connected) {
      reportIpcFailure();
      return;
    }
    if (daemonWatch.unresponsive())
      reportHealthIncident(
        "DAEMON_UNRESPONSIVE",
        "本地服务没有响应。请查看托盘日志；系统不会自动重试文档写入。",
      );
    try {
      // IPC failure may be asynchronous. Providing the callback also prevents
      // an unhandled child-process error if the channel closes after the check.
      thisProc.send({ type: "health_check", id: daemonWatch.request() }, (error) => {
        if (error) reportIpcFailure();
      });
    } catch {
      reportIpcFailure();
    }
  }, 10000);
  stableTimer = null; // A live but unready process has not earned a fresh budget.

  daemonProcess.stdout.on("data", (chunk) => {
    logStream?.write(chunk);
    if (daemonProcess !== thisProc) return;
    const text = chunk.toString();
    if (
      process.env.DRAFTSPECT_DESKTOP_LAUNCH !== "1" &&
      !appIsQuitting &&
      !daemonExitIntent.has(thisProc) &&
      /\[daemon\] (?:HTTPS|HTTP) server listening on/.test(text)
    ) {
      daemonStatus = "running";
      markDaemonReady(thisProc);
      // NB: do not reset restartAttempts here — a crash-looping daemon prints
      // this on every start. The counter is only cleared by the stability
      // timer once the process has actually stayed up (see startDaemon).
      updateTray();
    }
    const m = /Starting session for (.+?) \(/.exec(text);
    if (m) {
      currentWorkspace = m[1];
      updateTray();
    }
  });

  daemonProcess.stderr.on("data", (chunk) => {
    logStream?.write(chunk);
  });

  daemonProcess.on("error", (error) => {
    if (daemonProcess !== thisProc) return;
    if (!thisProc.pid) {
      daemonProcess = null;
      clearInterval(daemonHealthTimer);
      clearTimeout(stableTimer);
      stableTimer = null;
      clearTimeout(daemonStartupTimer);
      daemonStartupTimer = null;
    }
    reportStartupFailure(error);
  });

  daemonProcess.on("message", (msg) => {
    if (daemonProcess !== thisProc) return;
    if (msg?.type === "health_result") {
      if (!daemonWatch.reply(msg.id)) return;
      officeHealth = {
        phase: msg.phase,
        receivedAt: Date.now(),
        connected_panes_by_host: msg.connected_panes_by_host,
      };
      if (msg.unresponsive_active_panes > 0)
        reportHealthIncident(
          "PANE_UNRESPONSIVE",
          "Office 侧栏的活动请求长时间没有应用层响应。它可能在后台或未响应；请检查 Word / Excel，文档操作不会自动重试。",
        );
      else if (healthIncident) {
        healthIncident = null;
        updateTray();
      }
      updateTray();
      return;
    }
    void handleDaemonMessage(msg, thisProc).catch(() => {
      if (daemonProcess !== thisProc) return;
      desktopIncident({ code: "DESKTOP_DAEMON_IPC_FAILED", processType: "daemon" });
    });
  });

  daemonProcess.on("exit", (code, signal) => {
    if (daemonProcess !== thisProc) return;
    // A daemon handling SIGTERM exits with (0, null). Record intent against
    // the actual child instead of inferring it from the signal or global UI.
    const intent = daemonExitIntent.get(thisProc);
    daemonExitIntent.delete(thisProc);
    const wasIntentional = intent === "stop";
    daemonProcess = null;
    officeHealth = null;
    clearInterval(daemonHealthTimer);
    healthIncident = null;
    clearTimeout(daemonStartupTimer);
    daemonStartupTimer = null;
    publishDesktopPhase(appIsQuitting ? "quitting" : "js-started", { serviceReady: false });
    // The process is gone before it proved stable — cancel the pending reset
    // so a fast crash-loop keeps spending the restart budget.
    if (stableTimer) {
      clearTimeout(stableTimer);
      stableTimer = null;
    }
    if (intent === "startup-failed") {
      daemonStatus = "blocked";
    } else if (intent === "restart" && !appIsQuitting) {
      restartAttempts = 0;
      void startDaemon().catch(reportStartupFailure);
      return;
    } else if (wasIntentional) {
      daemonStatus = "stopped";
    } else {
      daemonStatus = "crashed";
      logStream?.write(`\n[app] daemon exited unexpectedly (code=${code}, signal=${signal})\n`);
      desktopLifecycle?.record("incident", {
        code: "DESKTOP_DAEMON_EXIT",
        processType: "daemon",
        exitCode: code,
        reason: signal ? "killed" : "abnormal-exit",
      });
      if (!appIsQuitting && restartAttempts < MAX_RESTART) {
        restartAttempts++;
        // Retain the bounded backoff, but revoke the scheduled start when a
        // manual Stop/Start/Restart supersedes it.
        const delay = restartDelayMs(restartAttempts);
        logStream?.write(
          `[app] auto-restarting (attempt ${restartAttempts}/${MAX_RESTART}, in ${delay}ms)\n`,
        );
        const timer = setTimeout(() => {
          if (restartTimer !== timer) return;
          restartTimer = null;
          void startDaemon().catch(reportStartupFailure);
        }, delay);
        restartTimer = timer;
      } else {
        logStream?.write(`[app] giving up after ${MAX_RESTART} restart attempts\n`);
      }
    }
    updateTray();
  });
}

function cancelScheduledRestart() {
  if (restartTimer !== null) clearTimeout(restartTimer);
  restartTimer = null;
}

function stopDaemon() {
  cancelScheduledRestart();
  pendingDaemonStart = null;
  daemonStatus = daemonProcess ? "stopping" : "stopped";
  officeHealth = null;
  clearTimeout(daemonStartupTimer);
  daemonStartupTimer = null;
  publishDesktopPhase(appIsQuitting ? "quitting" : "js-started", { serviceReady: false });
  if (daemonProcess) {
    daemonExitIntent.set(daemonProcess, "stop");
    daemonProcess.kill("SIGTERM");
  }
  updateTray();
}

function restartDaemon() {
  if (appIsQuitting) return;
  cancelScheduledRestart();
  officeHealth = null;
  if (daemonProcess) {
    const proc = daemonProcess;
    daemonExitIntent.set(proc, "restart");
    daemonStatus = "starting";
    updateTray();
    proc.kill("SIGTERM");
  } else {
    restartAttempts = 0;
    return startDaemon();
  }
}

function markDaemonReady(child) {
  if (!child || daemonProcess !== child || appIsQuitting || daemonExitIntent.has(child)) return;
  clearTimeout(daemonStartupTimer);
  daemonStartupTimer = null;
  if (stableTimer) return;
  stableTimer = setTimeout(() => {
    if (daemonProcess === child && daemonStatus === "running" && restartAttempts > 0) {
      restartAttempts = 0;
      logStream?.write(
        `[app] daemon ready and stable for ${STABLE_AFTER_MS / 1000}s — restart counter reset\n`,
      );
    }
  }, STABLE_AFTER_MS);
}

// --------------------------------------------------------------------------
// IPC: handle control requests from the daemon. Currently just the native
// folder/file picker — the taskpane forwards picker requests over WebSocket
// to the daemon, which proxies them up here so we can show a real macOS
// open panel (NSOpenPanel, with sidebar shortcuts, Google Drive, iCloud,
// recent items — everything a synthetic in-page modal can't do).
// --------------------------------------------------------------------------
async function handleDaemonMessage(msg, sourceProcess = daemonProcess) {
  if (!msg) return;
  if (
    !sourceProcess ||
    sourceProcess !== daemonProcess ||
    appIsQuitting ||
    daemonExitIntent.has(sourceProcess)
  )
    return;
  // Structured readiness signal over the IPC channel — replaces sniffing
  // stdout for a log substring (fragile: log copy changes, line
  // chunking). The stdout heuristic below is kept as a fallback for any
  // path that doesn't emit this.
  if (msg.type === "daemon_ready") {
    daemonStatus = "running";
    markDaemonReady(sourceProcess);
    publishDesktopPhase("service-ready");
    // Not reset here for the same reason as the stdout heuristic: a
    // crash-looping daemon emits daemon_ready on every start. Only the
    // stability timer clears the counter.
    if (msg.cwd) currentWorkspace = msg.cwd;
    updateTray();
    return;
  }
  if (msg.type !== "pick_path") return;
  const reply = (payload) => {
    try {
      if (
        sourceProcess !== daemonProcess ||
        !sourceProcess.connected ||
        appIsQuitting ||
        daemonExitIntent.has(sourceProcess)
      )
        return;
      sourceProcess.send({ type: "pick_path_result", id: msg.id, ...payload }, (error) => {
        if (error)
          desktopLifecycle?.record("incident", {
            code: "DESKTOP_DAEMON_IPC_FAILED",
            processType: "daemon",
          });
      });
    } catch {
      desktopLifecycle?.record("incident", {
        code: "DESKTOP_DAEMON_IPC_FAILED",
        processType: "daemon",
      });
    }
  };
  try {
    // Single-mode dialogs only. Windows cannot show a combined
    // file+directory picker — ['openFile','openDirectory'] silently
    // degrades to a directory-only selector there, so files never appear
    // (the symptom: an empty chooser when adding a file context entry).
    // Callers pick exactly one mode: include_files → file, else folder.
    const properties = msg.include_files ? ["openFile"] : ["openDirectory", "createDirectory"];
    // We're tray-only (no app window), so the OS open panel surfaces
    // BEHIND whatever is frontmost (Word/Excel). macOS: force-focus the
    // app and the panel comes forward (dock is hidden, so nothing else
    // shows). Windows: app.focus() just flashes the taskbar and the
    // dialog still opens behind Excel — instead, parent the dialog to a
    // transient off-screen always-on-top window so it's brought to the
    // foreground. The helper window is positioned off-screen and
    // destroyed immediately after, so it never visibly appears.
    let dialogParent = null;
    if (process.platform === "darwin") {
      app.focus({ steal: true });
    } else if (process.platform === "win32") {
      dialogParent = new BrowserWindow({
        width: 1,
        height: 1,
        x: -4000,
        y: -4000,
        show: false,
        frame: false,
        skipTaskbar: true,
        alwaysOnTop: true,
      });
      dialogParent.showInactive();
      dialogParent.setAlwaysOnTop(true, "screen-saver");
      dialogParent.focus();
    }
    const dialogOpts = {
      title: msg.title || (msg.include_files ? "Choose a folder or file" : "Choose a folder"),
      buttonLabel: msg.button_label || "Use this",
      defaultPath: msg.default_path || undefined,
      properties,
    };
    let result;
    try {
      result = dialogParent
        ? await dialog.showOpenDialog(dialogParent, dialogOpts)
        : await dialog.showOpenDialog(dialogOpts);
    } finally {
      dialogParent?.destroy();
    }
    if (result.canceled || result.filePaths.length === 0) {
      reply({ ok: true, canceled: true });
      return;
    }
    const picked = result.filePaths[0];
    const s = await stat(picked);
    reply({ ok: true, path: picked, kind: s.isDirectory() ? "directory" : "file" });
  } catch (err) {
    reply({ ok: false, error: err.message });
  }
}

// --------------------------------------------------------------------------
// Tray UI
// --------------------------------------------------------------------------
function statusLabel() {
  if (healthIncident)
    return healthIncident === "DAEMON_UNRESPONSIVE"
      ? "● Service unresponsive — check logs"
      : "● Office pane needs attention";
  switch (daemonStatus) {
    case "starting":
      return "● Starting…";
    case "running":
      return "● Local service ready — Office/model checked separately";
    case "blocked":
      return "● Startup needs attention";
    case "crashed":
      return restartAttempts >= MAX_RESTART
        ? "● Crashed (won't restart)"
        : `● Crashed (restart ${restartAttempts}/${MAX_RESTART})`;
    case "stopped":
      return "● Stopped";
    case "stopping":
      return "● Stopping — waiting for owned service exit";
    default:
      return "● ?";
  }
}

function workspaceLabel() {
  if (!currentWorkspace) return "Workspace: (none)";
  const name = currentWorkspace.split(/[\\/]/).filter(Boolean).pop();
  return `Workspace: ${name}`;
}

function buildMenu() {
  return Menu.buildFromTemplate([
    { label: statusLabel(), enabled: false },
    { label: workspaceLabel(), enabled: false },
    ...officeConnectionDetail(addinRegistration, daemonProcess ? officeHealth : null)
      .split("\n")
      .map((label) => ({ label, enabled: false })),
    { label: "启动状态 / Startup status…", click: () => showDesktopStatus().catch(() => {}) },
    { type: "separator" },
    { label: "Open logs", click: () => shell.openPath(LOG_FILE) },
    {
      label: daemonProcess ? "Restart daemon" : "Start daemon",
      click: restartDaemon,
    },
    ...(daemonProcess ? [{ label: "Stop daemon", click: stopDaemon }] : []),
    {
      label: "Show in Finder",
      submenu: [
        { label: "Project folder", click: () => shell.openPath(PROJECT_ROOT) },
        ...(currentWorkspace
          ? [{ label: "Current workspace", click: () => shell.openPath(currentWorkspace) }]
          : []),
        { label: "Log file", click: () => shell.showItemInFolder(LOG_FILE) },
      ],
    },
    { type: "separator" },
    ...(process.platform === "darwin"
      ? [
          { label: "Initialize Word Add-in…", click: runInitializeWordAddin },
          { label: "Initialize Excel Add-in…", click: runInitializeExcelAddin },
        ]
      : []),
    { label: "Open Microsoft Word", click: () => shell.openExternal("ms-word:") },
    { label: "Open Microsoft Excel", click: () => shell.openExternal("ms-excel:") },
    { type: "separator" },
    {
      label: addinInstalled
        ? "Reinstall add-in in Word + Excel"
        : "Install add-in in Word + Excel…",
      click: () => runInstall({ interactive: true }),
    },
    ...(addinInstalled
      ? [{ label: "Uninstall add-in", click: () => runUninstall({ interactive: true }) }]
      : []),
    { type: "separator" },
    { label: "Quit Draftspect for Office", click: () => app.quit() },
  ]);
}

function updateTray() {
  desktopWindow?.refresh();
  if (!tray) return;
  tray.setToolTip(`Draftspect for Office — ${statusLabel().replace(/^●\s*/, "")}`);
  tray.setContextMenu(buildMenu());
}

// --------------------------------------------------------------------------
// Add-in sideload — install / uninstall the Word + Excel manifests so the
// user never has to manually drop files into wef/ or touch Trust Center.
// --------------------------------------------------------------------------
let addinInstalled = false;

async function runInitializeExcelAddin() {
  try {
    await openExcelBootstrap({ openPath: openInMicrosoftExcel });
    await dialog.showMessageBox({
      type: "info",
      title: "Excel initialization workbook opened",
      message: "Wait for the Draftspect pane in the synthetic workbook.",
      detail:
        "Then open your target workbook and click Home → 显示 Draftspect 窗格 there. " +
        "Close the synthetic workbook without saving.\n\n" +
        "Repeat after fully quitting and restarting Excel if the ordinary menu still cannot activate. " +
        "This fallback opens a fresh marker-only workbook, does not bind your files, " +
        "does not clear Office caches and does not start the daemon or make a model request.",
      buttons: ["OK"],
    });
    return true;
  } catch (err) {
    logStream?.write(`[app] Excel initialization failed: ${err.message}\n`);
    await dialog
      .showMessageBox({
        type: "error",
        title: "Couldn't initialize the Excel add-in",
        message: err.message,
        buttons: ["OK"],
      })
      .catch(() => {});
    return false;
  }
}

async function runInitializeWordAddin() {
  try {
    await openWordBootstrap({ openPath: (path) => shell.openPath(path) });
    await dialog.showMessageBox({
      type: "info",
      title: "Word initialization document opened",
      message: "Wait for the Draftspect pane in the synthetic document.",
      detail:
        "Then open your own document and click Home → 显示 Draftspect 窗格 there. " +
        "Close the synthetic document without saving.\n\n" +
        "Repeat Initialize Word Add-in after fully quitting and restarting Word only if " +
        "manual menu activation still fails. This is a per-process fallback, not permanent registration. " +
        "It does not start the daemon; start it separately if it is stopped.",
      buttons: ["OK"],
    });
    return true;
  } catch (err) {
    logStream?.write(`[app] Word initialization failed: ${err.message}\n`);
    await dialog
      .showMessageBox({
        type: "error",
        title: "Couldn't initialize the Word add-in",
        message: err.message,
        buttons: ["OK"],
      })
      .catch(() => {});
    return false;
  }
}

async function refreshAddinInstalled() {
  try {
    addinRegistration = await readAddinRegistration();
    addinInstalled = await isAddinInstalled();
  } catch {
    addinRegistration = [];
    addinInstalled = false;
  }
  updateTray();
}

async function runInstall({ interactive }) {
  try {
    const result = await installAddin();
    await refreshAddinInstalled();
    if (interactive) {
      dialog
        .showMessageBox({
          type: "info",
          title: "Draftspect for Office installed",
          message: "The manifests were installed; Office activation must be checked separately.",
          detail:
            addinActivationInstructions() +
            "\n\n" +
            (process.platform === "win32"
              ? `Developer manifests registered at:\n${result.catalog}`
              : `Manifests copied to each app's wef/ folder.`),
          buttons: ["OK"],
        })
        .catch(() => {});
    }
    return result;
  } catch (err) {
    logStream?.write(`[app] install failed: ${err.message}\n`);
    if (interactive) {
      dialog
        .showMessageBox({
          type: "error",
          title: "Couldn't install the add-in",
          message: err.message,
          detail: addinInstallRecoveryInstructions(),
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  }
}

async function runUninstall({ interactive }) {
  try {
    await uninstallAddin();
    await refreshAddinInstalled();
    if (interactive) {
      dialog
        .showMessageBox({
          type: "info",
          title: "Add-in uninstalled",
          message: "Draftspect for Office is no longer registered with Word or Excel.",
          detail:
            "The daemon is still running. Quit Draftspect for Office from the tray menu to stop it entirely.",
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  } catch (err) {
    logStream?.write(`[app] uninstall failed: ${err.message}\n`);
    if (interactive) {
      dialog
        .showMessageBox({
          type: "error",
          title: "Couldn't uninstall the add-in",
          message: err.message,
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  }
}

// First-run prompt — invoked once at app start if the add-in isn't already
// registered. Skipped silently on unsupported platforms (Linux, etc.).
async function offerFirstRunInstall() {
  // Preview registration is an explicit tray action, never a startup prompt.
  if (process.env.DRAFTSPECT_OFFER_INSTALL !== "1") return;
  if (process.platform !== "darwin" && process.platform !== "win32") return;
  if (addinInstalled) return;
  const { response } = await dialog
    .showMessageBox({
      type: "question",
      title: "Install Draftspect for Office in Word + Excel?",
      message:
        "Draftspect for Office can install itself in Word and Excel automatically — no manifest copying or registry editing needed.",
      detail:
        "Click Install to register the add-in now. You can install later from the tray menu " +
        "if you'd prefer.\n\n" +
        addinActivationInstructions(),
      buttons: ["Install", "Not now"],
      defaultId: 0,
      cancelId: 1,
    })
    .catch(() => ({ response: 1 }));
  if (response === 0) await runInstall({ interactive: true });
}

// --------------------------------------------------------------------------
// App lifecycle
// --------------------------------------------------------------------------
if (ownsDesktopInstance)
  app
    .whenReady()
    .then(async () => {
      // Background-only app — no dock icon on macOS.
      if (process.platform === "darwin") app.dock?.hide();

      openLogStream();
      if (process.env.DRAFTSPECT_DESKTOP_LAUNCH === "1") await getDesktopWindow().show();

      // Tray icon. macOS expects a template (monochrome with transparency, auto-
      // recolored per theme); Windows expects a full-color .ico (or PNG fallback).
      // Convention: tray-icon.png is the macOS template; tray-icon-win.ico is the
      // Windows full-color version. Falls back to the PNG if .ico is missing.
      const macIconPath = join(__dirname, "tray-icon.png");
      const winIconPath = join(__dirname, "tray-icon-win.ico");
      const iconPath =
        process.platform === "win32" && fs.existsSync(winIconPath) ? winIconPath : macIconPath;
      const icon = nativeImage.createFromPath(iconPath);
      if (process.platform === "darwin") icon.setTemplateImage(true);
      tray = new Tray(icon);
      tray.setToolTip("Draftspect for Office");

      // Check whether the add-in is already registered and refresh the tray
      // menu, then offer to install on first run.
      await refreshAddinInstalled();
      updateTray();
      if (desktopStartupError) return;
      await startDaemon();
      if (daemonStatus !== "blocked")
        offerFirstRunInstall().catch((err) =>
          logStream?.write(`[app] first-run prompt failed: ${err.message}\n`),
        );
    })
    .catch(async (error) => {
      await reportStartupFailure(error);
      if (!tray) app.quit();
    });

// Closing the local status window hides it; explicit Quit owns service shutdown.
if (ownsDesktopInstance)
  app.on("window-all-closed", (e) => {
    if (e && e.preventDefault) e.preventDefault();
  });

// JavaScript exceptions occur after native Electron initialization. Native
// pre-JS death is observed by the existing launcher, not by these handlers.
if (ownsDesktopInstance) {
  const fatal = (code) => {
    if (fatalMainError || appIsQuitting) return;
    fatalMainError = true;
    desktopLifecycle?.record("incident", { code, processType: "main" });
    publishDesktopPhase("startup-failed", { code, serviceReady: false });
    appIsQuitting = true;
    cancelScheduledRestart();
    void waitForOwnedExit(daemonProcess, stopDaemon).then(
      () => app.exit(1),
      () => app.exit(1), // parent disconnect makes the owned daemon shut down itself
    );
  };
  process.on("uncaughtException", () => fatal("DESKTOP_MAIN_EXCEPTION"));
  process.on("unhandledRejection", () => fatal("DESKTOP_MAIN_REJECTION"));
  app.on("child-process-gone", (_event, details) => {
    if (appIsQuitting) return;
    desktopIncident({
      code: "DESKTOP_CHILD_GONE",
      processType: details?.type || "unknown",
      reason: details?.reason,
      exitCode: details?.exitCode,
    });
  });
  app.on("will-quit", () => {
    clearInterval(desktopHeartbeat);
    desktopLifecycle?.close();
  });
}

async function handleBeforeQuit(event) {
  if (desktopQuitReady) return;
  event.preventDefault();
  if (desktopQuitPending) return;
  desktopQuitPending = true;
  appIsQuitting = true;
  publishDesktopPhase("quitting");
  try {
    // Keep the desktop lock until this exact daemon has exited; a new owner
    // must not race our old listener shutdown. Timeout leaves the app/lock
    // alive, with a visible error, rather than killing an unknown process.
    await waitForOwnedExit(daemonProcess, stopDaemon);
    logStream?.write(`=== ${new Date().toISOString()} app quit ===\n`);
    logStream?.end();
    desktopQuitReady = true;
    // Electron closes windows before will-quit; our close-to-hide window must
    // be destroyed only after the daemon-exit barrier, or it would cancel Quit.
    desktopWindow?.dispose();
    desktopWindow = null;
    app.quit();
  } catch (error) {
    appIsQuitting = false;
    await reportStartupFailure(error);
  } finally {
    desktopQuitPending = false;
  }
}
if (ownsDesktopInstance) app.on("before-quit", handleBeforeQuit);
