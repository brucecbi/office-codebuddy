// SDK 0.3.243 has no public child-exit promise: close() releases its lock and
// sends SIGTERM synchronously. Keep this private-API dependency isolated and
// version/shape checked. Observe only this session's own direct ChildProcess;
// never patch global spawn, scan processes, or infer exit from a closed stream.
import { ChildProcess } from "node:child_process";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";

const require = createRequire(import.meta.url);
const sdkVersion = JSON.parse(
  readFileSync(join(dirname(require.resolve("@tencent-ai/agent-sdk")), "..", "package.json")),
).version;

export function observeCodeBuddyClose(session, { version = sdkVersion, onProcess } = {}) {
  const transport = session?._transport;
  if (
    version !== "0.3.243" ||
    !transport ||
    !Object.hasOwn(transport, "process") ||
    transport.process !== null ||
    typeof transport.start !== "function" ||
    typeof transport.sendControlResponse !== "function" ||
    typeof transport.sendControlErrorResponse !== "function" ||
    typeof session.close !== "function"
  )
    throw new Error("CodeBuddy process-exit adapter unsupported; refusing an unobservable session");
  const originalStart = transport.start;
  let closing = false,
    child = null,
    exit = null,
    exitEvidence = null,
    observationError = null;
  // SDK 0.3.243 fires MCP requests without awaiting their handlers. Closing
  // rejects pending MCP work, whose late handler then attempts a control reply
  // through a cleared process and rejects unhandled ("Transport not started").
  // Only replies to this closed transport are obsolete. Keep active errors
  // visible; do not install a global unhandled-rejection handler or weaken the
  // child-exit barrier. Also cover a close initiated inside the SDK itself.
  for (const name of ["sendControlResponse", "sendControlErrorResponse"]) {
    const original = transport[name];
    transport[name] = function (...args) {
      if (closing || this.closed === true) return;
      return original.apply(this, args);
    };
  }
  transport.start = function (...args) {
    // ensureReady() asynchronously resolves CLI metadata before start(). A Stop
    // during that await must also prevent a child being spawned after close().
    if (closing) throw new Error("CodeBuddy session closed before process start");
    try {
      return originalStart.apply(this, args);
    } finally {
      if (this.process && this.process !== child) {
        if (child || !(this.process instanceof ChildProcess)) {
          observationError = new Error("CodeBuddy changed its child-process contract");
          throw observationError;
        }
        child = this.process;
        exit = new Promise((resolve) => {
          const observed = (evidence) => {
            exitEvidence = evidence;
            resolve(evidence);
          };
          child.once("exit", (code, signal) => observed({ pid: child.pid, code, signal }));
          child.once("error", (error) => {
            // spawn failure with no PID means no OS child ever existed.
            if (!child.pid) observed({ pid: null, spawn_error: error.code || error.name });
            // With a PID, only a later exit event proves the child has ended.
          });
          if (child.exitCode !== null || child.signalCode !== null)
            observed({ pid: child.pid, code: child.exitCode, signal: child.signalCode });
        });
        exit.catch(() => {});
        onProcess?.(child);
      }
    }
  };
  const getCloseStatus = () => ({
    phase: observationError
      ? "unobservable"
      : exitEvidence
        ? "exited"
        : child
          ? closing
            ? "closing"
            : "running"
          : "not_started",
    close_requested: closing,
    pid: child?.pid ?? null,
    can_terminate: closing && !!child?.pid && !exitEvidence && !observationError,
    ...(exitEvidence ? { exit: exitEvidence } : {}),
  });
  const failRecovery = (code, message) => {
    throw Object.assign(new Error(message), { code });
  };
  return {
    getCloseStatus,
    terminateOwnedProcess({ expectedPid, signal = "SIGTERM", confirmForce = false } = {}) {
      if (!closing)
        failRecovery(
          "RUNTIME_NOT_CLOSING",
          "Stop/release this session before requesting process recovery.",
        );
      if (observationError)
        failRecovery(
          "RUNTIME_EXIT_UNOBSERVABLE",
          "Runtime process identity cannot be verified; no signal was sent.",
        );
      if (!["SIGTERM", "SIGKILL"].includes(signal))
        failRecovery(
          "RUNTIME_SIGNAL_INVALID",
          "Only SIGTERM or explicitly confirmed SIGKILL is supported.",
        );
      if (signal === "SIGKILL" && confirmForce !== true)
        failRecovery(
          "RUNTIME_FORCE_CONFIRMATION_REQUIRED",
          "Force termination needs an explicit confirmation for this captured process.",
        );
      if (!child?.pid || !Number.isSafeInteger(expectedPid) || expectedPid !== child.pid)
        failRecovery(
          "RUNTIME_PROCESS_CHANGED",
          "The captured runtime process changed or never started; refresh its status first.",
        );
      if (exitEvidence) return { ...getCloseStatus(), signal_sent: false };
      // ChildProcess.kill addresses only this captured direct child, never a
      // discovered PID/process name. Sending a signal is NOT exit evidence.
      const sent = child.kill(signal);
      if (!sent)
        failRecovery(
          "RUNTIME_SIGNAL_FAILED",
          "The runtime signal was not sent; process exit is still unconfirmed.",
        );
      return { ...getCloseStatus(), signal_sent: true, signal };
    },
    close() {
      closing = true;
      session.close();
    },
    async waitForExit() {
      if (!closing) throw new Error("Close the CodeBuddy session before awaiting process exit");
      if (observationError) throw observationError;
      return exit ? await exit : { pid: null, process_never_started: true };
    },
  };
}
