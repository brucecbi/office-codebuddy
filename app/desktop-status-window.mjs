// Local-only desktop controls. No provider, Office, daemon or user-config imports.
import { fileURLToPath } from "node:url";
import { restartDelayMs } from "./restart-backoff.mjs";

const PAGE = new URL("./desktop-status.html", import.meta.url);
const PRELOAD = fileURLToPath(new URL("./desktop-status-preload.cjs", import.meta.url));
const GET_STATUS = "draftspect:desktop-status:get";
const ACT = "draftspect:desktop-status:act";
const ACTIONS = new Set([
  "start",
  "restart",
  "stop",
  "initialize-word",
  "initialize-excel",
  "open-logs",
  "quit",
]);
const safeText = (value, limit) =>
  typeof value === "string"
    ? value.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "").slice(0, limit)
    : "";

function statusSnapshot(value) {
  const source = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return {
    status: safeText(source.status, 120) || "Unknown",
    detail: safeText(source.detail, 4000),
    runtime: safeText(source.runtime, 120),
    error: safeText(source.error, 2000),
  };
}

export function createDesktopStatusWindow({
  BrowserWindow,
  ipcMain,
  app,
  readStatus,
  onAction,
  onVisible = () => {},
  onIncident = () => {},
  timers = { setTimeout, clearTimeout },
  loadTimeoutMs = 15_000,
  stableAfterMs = 45_000,
}) {
  let window = null;
  let disposed = false;
  let ready = false;
  let visibleRequested = false;
  let pendingLoad = null;
  let actionPending = false;
  let rendererDead = false;
  let lastStatus = statusSnapshot({ status: "Starting" });
  let recoveryAttempts = 0;
  let recoveryTimer = null;
  let stabilityTimer = null;
  let cancelLoad = null;
  const cancelRecovery = () => {
    timers.clearTimeout(recoveryTimer);
    recoveryTimer = null;
    timers.clearTimeout(stabilityTimer);
    stabilityTimer = null;
  };

  function refresh() {
    if (disposed) return lastStatus;
    try {
      lastStatus = statusSnapshot(readStatus());
    } catch {
      lastStatus = statusSnapshot({
        status: "Unavailable",
        error: "无法读取本地服务状态，请查看日志。",
      });
    }
    if (window && !window.isDestroyed()) window.setTitle("Draftspect — " + lastStatus.status);
    return { ...lastStatus };
  }

  function trusted(event) {
    if (disposed || !window || window.isDestroyed()) return false;
    const contents = window.webContents;
    try {
      return (
        event?.sender === contents &&
        event.senderFrame === contents.mainFrame &&
        event.senderFrame.url === PAGE.href &&
        contents.getURL() === PAGE.href
      );
    } catch {
      return false;
    }
  }

  const getStatus = (event, ...args) => {
    if (!trusted(event) || args.length !== 0) throw new Error("DESKTOP_IPC_DENIED");
    return refresh();
  };
  const act = async (event, ...args) => {
    if (
      !trusted(event) ||
      args.length !== 1 ||
      typeof args[0] !== "string" ||
      !ACTIONS.has(args[0])
    )
      throw new Error("DESKTOP_IPC_DENIED");
    if (actionPending) return { ok: false, error: "操作正在进行，请稍候。" };
    actionPending = true;
    try {
      await onAction(args[0]);
      return { ok: true, status: refresh() };
    } catch {
      return { ok: false, error: "操作未完成，请检查服务状态或查看日志。", status: refresh() };
    } finally {
      actionPending = false;
    }
  };
  ipcMain.handle(GET_STATUS, getStatus);
  ipcMain.handle(ACT, act);

  function reveal() {
    if (!window || window.isDestroyed() || !visibleRequested || disposed) return;
    if (window.isMinimized()) window.restore();
    window.show();
    window.focus();
    onVisible();
  }

  function show() {
    if (disposed) return Promise.resolve();
    visibleRequested = true;
    if (window && !window.isDestroyed() && rendererDead) {
      const dead = window;
      window = null;
      ready = false;
      pendingLoad = null;
      cancelRecovery();
      dead.destroy();
    }
    if (window && !window.isDestroyed()) {
      refresh();
      if (ready) reveal();
      return pendingLoad ?? Promise.resolve();
    }
    return createWindow();
  }

  function createWindow() {
    if (disposed) return Promise.resolve();
    ready = false;
    rendererDead = false;
    window = new BrowserWindow({
      width: 620,
      height: 610,
      minWidth: 480,
      minHeight: 510,
      title: "Draftspect",
      show: false,
      autoHideMenuBar: true,
      backgroundColor: "#f5f7f8",
      webPreferences: {
        preload: PRELOAD,
        contextIsolation: true,
        sandbox: true,
        nodeIntegration: false,
        webSecurity: true,
        allowRunningInsecureContent: false,
        webviewTag: false,
        partition: "draftspect-local-status",
      },
    });
    const owned = window;
    owned.setMenu?.(null);
    owned.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
    for (const event of ["will-navigate", "will-frame-navigate", "will-attach-webview"])
      owned.webContents.on(event, (navigation) => navigation.preventDefault());
    owned.webContents.session.setPermissionRequestHandler((_contents, _permission, callback) =>
      callback(false),
    );
    owned.webContents.session.setPermissionCheckHandler(() => false);
    owned.on("close", (event) => {
      if (disposed) return;
      event.preventDefault();
      visibleRequested = false;
      cancelRecovery();
      owned.hide();
    });
    owned.on("closed", () => {
      if (window === owned) {
        window = null;
        ready = false;
        pendingLoad = null;
      }
    });
    owned.once("ready-to-show", () => {
      if (window !== owned) return;
      ready = true;
      reveal();
      completeLoad();
      timers.clearTimeout(stabilityTimer);
      stabilityTimer = timers.setTimeout(() => {
        if (!disposed && window === owned && ready) recoveryAttempts = 0;
      }, stableAfterMs);
    });
    const recoverRenderer = (details) => {
      if (disposed || window !== owned) return;
      rendererDead = true;
      onIncident({
        code: "DESKTOP_RENDERER_GONE",
        processType: "renderer",
        reason: details?.reason,
        exitCode: details?.exitCode,
      });
      cancelRecovery();
      failLoad("DESKTOP_RENDERER_GONE");
      ready = false;
      if (!visibleRequested) return;
      if (recoveryAttempts >= 3) {
        onIncident({
          code: "DESKTOP_RENDERER_RECOVERY_EXHAUSTED",
          processType: "renderer",
          attempt: recoveryAttempts,
        });
        return;
      }
      const delay = restartDelayMs(++recoveryAttempts);
      recoveryTimer = timers.setTimeout(() => {
        recoveryTimer = null;
        if (disposed || window !== owned || !visibleRequested) return;
        // Only the local status page is recreated. Daemon requests and Office
        // operations never cross this recovery path.
        window = null;
        owned.destroy();
        void createWindow().catch(() => {
          onIncident({ code: "DESKTOP_STATUS_LOAD_FAILED", processType: "renderer" });
        });
      }, delay);
    };
    owned.webContents.on("render-process-gone", (_event, details) => recoverRenderer(details));
    owned.on("unresponsive", () => {
      if (!disposed && window === owned)
        onIncident({
          code: "DESKTOP_RENDERER_UNRESPONSIVE",
          processType: "renderer",
          reason: "unresponsive",
        });
    });
    owned.on("responsive", () => {
      if (!disposed && window === owned)
        onIncident({ processType: "renderer", reason: "responsive" });
    });
    refresh();
    let loaded = false,
      settled = false,
      resolveLoad,
      rejectLoad;
    const completeLoad = () => {
      if (!loaded || !ready || settled) return;
      settled = true;
      timers.clearTimeout(loadTimer);
      cancelLoad = null;
      resolveLoad();
    };
    const failLoad = (code) => {
      if (settled) return;
      settled = true;
      timers.clearTimeout(loadTimer);
      cancelLoad = null;
      rejectLoad(Object.assign(new Error(code), { code }));
    };
    pendingLoad = new Promise((resolve, reject) => {
      resolveLoad = resolve;
      rejectLoad = reject;
    });
    cancelLoad = () => failLoad("DESKTOP_STATUS_LOAD_FAILED");
    const loadTimer = timers.setTimeout(() => {
      failLoad("DESKTOP_STATUS_TIMEOUT");
    }, loadTimeoutMs);
    // An owned destruction cancels startup; it cannot leave an infinite show().
    owned.once("closed", () => failLoad("DESKTOP_STATUS_LOAD_FAILED"));
    Promise.resolve()
      .then(() => owned.loadFile(fileURLToPath(PAGE)))
      .then(
        () => {
          loaded = true;
          completeLoad();
        },
        () => {
          failLoad("DESKTOP_STATUS_LOAD_FAILED");
        },
      );
    return pendingLoad;
  }

  function dispose() {
    if (disposed) return;
    disposed = true;
    visibleRequested = false;
    cancelRecovery();
    cancelLoad?.();
    app.removeListener("will-quit", dispose);
    // Destroy the renderer first so its polling stops before removing IPC.
    // Security checks remain active throughout teardown.
    if (window && !window.isDestroyed()) window.destroy();
    window = null;
    ipcMain.removeHandler(GET_STATUS);
    ipcMain.removeHandler(ACT);
  }
  app.on("will-quit", dispose);
  return { show, refresh, dispose };
}
