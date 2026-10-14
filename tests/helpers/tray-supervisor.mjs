// Execute the production supervisor functions without Electron/UI. Child
// processes and timers may be virtual or real, depending on the test.
import { readFile } from "node:fs/promises";
import { EventEmitter } from "node:events";
import vm from "node:vm";
import { dirname, delimiter } from "node:path";
import { createHealthWatch } from "../../app/health-watch.mjs";
import { restartDelayMs } from "../../app/restart-backoff.mjs";
import { officeConnectionDetail } from "../../app/addin-registration.mjs";

export async function trayHarness({
  findWorkspace = async () => null,
  spawnChild,
  overrides = {},
} = {}) {
  const source = await readFile(new URL("../../app/main.mjs", import.meta.url), "utf8");
  const state = source.slice(
    source.indexOf("let tray = null;"),
    source.indexOf("\nfunction reportHealthIncident("),
  );
  const start = source.indexOf("async function startDaemon(");
  const end = source.indexOf(
    "\n// --------------------------------------------------------------------------\n// IPC:",
    start,
  );
  if (!state || start < 0 || end < start)
    throw new Error("Production supervisor boundaries changed");
  const children = [],
    timers = [],
    logs = [],
    desktopWindows = [];
  const context = vm.createContext({
    findInitialWorkspace: findWorkspace,
    spawn: (...args) => {
      const child = spawnChild ? spawnChild(...args) : new EventEmitter();
      if (!spawnChild)
        Object.assign(child, {
          stdout: new EventEmitter(),
          stderr: new EventEmitter(),
          connected: true,
          signals: [],
          kill(signal) {
            this.signals.push(signal);
            return true;
          },
          send() {},
        });
      children.push(child);
      return child;
    },
    process: { execPath: "/synthetic/node", env: {} },
    PROJECT_ROOT: "/synthetic",
    DAEMON_ENTRY: "/synthetic/daemon.mjs",
    desktopLifecycle: { publish() {}, record() {} },
    createHealthWatch,
    restartDelayMs,
    officeConnectionDetail,
    setInterval: () => ({}),
    clearInterval() {},
    setTimeout: (fn, ms) => {
      const timer = { fn, ms, cancelled: false };
      timers.push(timer);
      return timer;
    },
    clearTimeout: (timer) => {
      if (timer) timer.cancelled = true;
    },
    updateTray() {},
    handleDaemonMessage() {},
    reportHealthIncident() {},
    app: { focus() {} },
    BrowserWindow: class {},
    ipcMain: {},
    createDesktopStatusWindow: (options) => {
      const window = {
        options,
        shown: [],
        show: async () => {
          window.shown.push(options.readStatus());
        },
        refresh: () => options.readStatus(),
        dispose() {},
      };
      desktopWindows.push(window);
      return window;
    },
    dialog: { showMessageBox: async () => ({ response: 0 }) },
    LOG_FILE: "/synthetic/daemon.log",
    dirname,
    delimiter,
    statusLabel: () => "Synthetic local status",
    __logs: logs,
    ...overrides,
  });
  vm.runInContext(
    state +
      "\n" +
      source.slice(start, end) +
      "\n" +
      source.slice(
        source.indexOf("async function handleDaemonMessage("),
        source.indexOf(
          "\n// --------------------------------------------------------------------------\n// Tray UI",
        ),
      ) +
      `
    logStream = { write: value => __logs.push(String(value).trim()) };
    globalThis.api = { start: startDaemon, stop: stopDaemon, restart: restartDaemon,
      owner: () => daemonProcess, status: () => daemonStatus,
      showDesktopStatus, getDesktopWindow };
  `,
    context,
  );
  return { ...context.api, children, timers, logs, desktopWindows };
}

export const tick = () => new Promise((resolve) => setImmediate(resolve));
