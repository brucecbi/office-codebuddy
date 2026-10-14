// Metadata-only local diagnostics and the private native-launcher handshake.
// No crash upload, document/session payload, exception text or environment dump.
import fs from "node:fs";
import { dirname, join, resolve } from "node:path";

export const DESKTOP_FAILURE_CODES = new Set([
  "DESKTOP_START_FAILED",
  "DESKTOP_LAUNCH_CHANNEL_INVALID",
  "DESKTOP_STATUS_LOAD_FAILED",
  "DESKTOP_STATUS_TIMEOUT",
  "DESKTOP_RENDERER_GONE",
  "DESKTOP_RENDERER_UNRESPONSIVE",
  "DESKTOP_RENDERER_RECOVERY_EXHAUSTED",
  "DESKTOP_CHILD_GONE",
  "DESKTOP_MAIN_EXCEPTION",
  "DESKTOP_MAIN_REJECTION",
  "DESKTOP_DAEMON_START_TIMEOUT",
  "DESKTOP_DAEMON_EXIT",
  "DESKTOP_DAEMON_IPC_FAILED",
  "DESKTOP_DIAGNOSTIC_UNAVAILABLE",
  "DESKTOP_LOG_UNAVAILABLE",
  "DESKTOP_SINGLE_INSTANCE_FAILED",
  "DESKTOP_STOP_TIMEOUT",
  "DESKTOP_STOP_FAILED",
  "DESKTOP_TLS_REQUIRED",
  "DESKTOP_NODE_UNAVAILABLE",
  "DESKTOP_CLI_UNAVAILABLE",
  "DESKTOP_PROJECT_UNAVAILABLE",
  "DESKTOP_STATE_UNAVAILABLE",
  "DESKTOP_TRANSPORT_INVALID",
  "DESKTOP_PORT_INVALID",
  "DESKTOP_PORT_CHECK_TIMEOUT",
  "DESKTOP_PORT_IN_USE",
  "DESKTOP_PORT_UNAVAILABLE",
  "DESKTOP_ENV_INVALID",
  "DESKTOP_PROJECT_INVALID",
  "DESKTOP_ASSET_MISSING",
  "DESKTOP_DEPENDENCY_MISSING",
  "DESKTOP_TLS_INVALID",
  "DESKTOP_HTTPS_REQUIRED",
  "DESKTOP_CLI_INVALID",
  "DESKTOP_CLI_NOT_FOUND",
  "DESKTOP_NODE_INVALID",
  "DESKTOP_NODE_NOT_FOUND",
]);
const PHASES = new Set([
  "js-started",
  "desktop-visible",
  "service-ready",
  "startup-failed",
  "duplicate",
  "quitting",
]);
const REASONS = new Set([
  "clean-exit",
  "abnormal-exit",
  "killed",
  "crashed",
  "oom",
  "launch-failed",
  "integrity-failure",
  "unresponsive",
  "responsive",
]);
const TYPES = new Set([
  "main",
  "renderer",
  "daemon",
  "GPU",
  "Utility",
  "Browser",
  "Zygote",
  "Sandbox helper",
  "unknown",
]);
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const codeOf = (value) => (DESKTOP_FAILURE_CODES.has(value) ? value : "DESKTOP_START_FAILED");
const invalid = () =>
  Object.assign(new Error("Private desktop launch channel is invalid."), {
    code: "DESKTOP_LAUNCH_CHANNEL_INVALID",
  });

export function createDesktopLifecycle({
  env = process.env,
  stateDir,
  processInfo = process,
  now = () => new Date(),
  fileSystem = fs,
} = {}) {
  const uid = processInfo.getuid?.();
  let channel = null,
    channelPath = null,
    channelIdentity = null,
    token = null,
    journal = null;
  let state = {
    schemaVersion: 1,
    pid: processInfo.pid,
    phase: "js-started",
    desktopVisible: false,
    serviceReady: false,
    heartbeat: 0,
  };
  let journalHealthy = true;
  const validFile = (stat, mode) =>
    stat.isFile() &&
    !stat.isSymbolicLink?.() &&
    stat.nlink === 1 &&
    (uid === undefined || stat.uid === uid) &&
    (stat.mode & 0o777) === mode;
  const privateParent = (path) => {
    const parent = dirname(path),
      stat = fileSystem.lstatSync(parent);
    return (
      stat.isDirectory() &&
      !stat.isSymbolicLink() &&
      (uid === undefined || stat.uid === uid) &&
      (stat.mode & 0o077) === 0 &&
      fileSystem.realpathSync(parent) === resolve(parent)
    );
  };
  const requested = [
    env.DRAFTSPECT_LAUNCH_STATUS_PATH,
    env.DRAFTSPECT_LAUNCH_TOKEN,
    env.DRAFTSPECT_LAUNCH_PARENT_PID,
  ];
  if (requested.some((value) => value !== undefined)) {
    try {
      const parentPid = Number(requested[2]);
      if (
        typeof requested[0] !== "string" ||
        !requested[0].startsWith("/") ||
        !UUID.test(requested[1] || "") ||
        !Number.isSafeInteger(parentPid) ||
        parentPid < 2 ||
        parentPid !== processInfo.ppid ||
        !privateParent(requested[0])
      )
        throw invalid();
      channelPath = requested[0];
      channel = fileSystem.openSync(
        channelPath,
        fileSystem.constants.O_RDWR |
          fileSystem.constants.O_NOFOLLOW |
          fileSystem.constants.O_NONBLOCK,
      );
      channelIdentity = fileSystem.fstatSync(channel);
      if (!validFile(channelIdentity, 0o600) || channelIdentity.size > 4096) throw invalid();
      const initial = JSON.parse(fileSystem.readFileSync(channel, "utf8"));
      if (initial.schemaVersion !== 1 || initial.token !== requested[1]) throw invalid();
      token = requested[1];
    } catch {
      if (channel !== null) fileSystem.closeSync(channel);
      throw invalid();
    }
  }
  if (stateDir) {
    try {
      const path = join(stateDir, "desktop", "desktop-events.jsonl");
      if (!privateParent(path)) throw invalid();
      journal = fileSystem.openSync(
        path,
        fileSystem.constants.O_WRONLY |
          fileSystem.constants.O_CREAT |
          fileSystem.constants.O_APPEND |
          fileSystem.constants.O_NOFOLLOW |
          fileSystem.constants.O_NONBLOCK,
        0o600,
      );
      if (!validFile(fileSystem.fstatSync(journal), 0o600)) throw invalid();
    } catch {
      if (journal !== null) fileSystem.closeSync(journal);
      journal = null;
    }
  }
  function record(event, fields = {}) {
    const value = {
      schemaVersion: 1,
      time: now().toISOString(),
      pid: processInfo.pid,
      event: PHASES.has(event) ? event : "incident",
    };
    if (fields.code) value.code = codeOf(fields.code);
    if (REASONS.has(fields.reason)) value.reason = fields.reason;
    if (TYPES.has(fields.processType)) value.processType = fields.processType;
    if (Number.isSafeInteger(fields.exitCode)) value.exitCode = fields.exitCode;
    if (Number.isSafeInteger(fields.attempt) && fields.attempt >= 0 && fields.attempt <= 3)
      value.attempt = fields.attempt;
    for (const name of ["electron", "chrome", "node"])
      if (/^\d+(?:\.\d+){1,3}$/.test(processInfo.versions?.[name] || ""))
        value[name] = processInfo.versions[name];
    if (journal !== null)
      try {
        fileSystem.writeSync(journal, JSON.stringify(value) + "\n");
      } catch {
        journalHealthy = false;
        /* local best effort; no remote fallback */
      }
    return value;
  }
  function publish(phase, fields = {}) {
    if (!PHASES.has(phase)) throw new TypeError("Unknown desktop launch phase");
    state.phase = phase;
    if (phase === "desktop-visible") state.desktopVisible = true;
    if (phase === "service-ready") state.serviceReady = true;
    if (typeof fields.desktopVisible === "boolean") state.desktopVisible = fields.desktopVisible;
    if (typeof fields.serviceReady === "boolean") state.serviceReady = fields.serviceReady;
    if (fields.code) state.code = codeOf(fields.code);
    if (phase === "startup-failed") state.failureCode = codeOf(fields.code);
    if (phase === "quitting") state.serviceReady = false;
    record(phase, fields);
    writeChannel(true);
    return { ...state };
  }
  function writeChannel(durable = false) {
    if (channel !== null) {
      try {
        const current = fileSystem.lstatSync(channelPath),
          owned = fileSystem.fstatSync(channel);
        if (
          !privateParent(channelPath) ||
          !validFile(current, 0o600) ||
          !validFile(owned, 0o600) ||
          current.ino !== channelIdentity.ino ||
          current.dev !== channelIdentity.dev ||
          owned.ino !== current.ino ||
          owned.dev !== current.dev
        )
          throw invalid();
        const data = Buffer.from(JSON.stringify({ ...state, token }) + "\n");
        let offset = 0;
        while (offset < data.length)
          offset += fileSystem.writeSync(channel, data, offset, data.length - offset, offset);
        fileSystem.ftruncateSync(channel, data.length);
        if (durable) fileSystem.fsyncSync(channel);
      } catch {
        throw invalid();
      }
    }
  }
  function heartbeat() {
    if (state.phase === "quitting" || channel === null) return;
    state.heartbeat++;
    // Liveness only: no journal growth or per-tick disk flush. Phase/exit
    // records remain durable; the monitor tolerates transient partial reads.
    writeChannel();
  }
  function close() {
    for (const fd of [channel, journal])
      if (fd !== null)
        try {
          fileSystem.closeSync(fd);
        } catch {
          /* closing only descriptors opened here */
        }
    channel = journal = null;
  }
  return {
    publish,
    heartbeat,
    record,
    close,
    snapshot: () => ({ ...state }),
    get diagnosticAvailable() {
      return journal !== null && journalHealthy;
    },
  };
}
