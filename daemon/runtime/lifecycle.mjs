// One generation owns one pane. The handle exists before asynchronous setup.
// Stop fences late messages immediately; transport teardown still runs.
//
// Turn timeout model (2026-09-21, after the 2026-09-20 incident where a long
// read-only contract analysis was killed mid-progress by a fixed 180s wall):
//  - idleMs: the turn is stopped only after this much INACTIVITY. Every SDK
//    message and every Office tool event must call activity() — a turn that
//    is making progress is never interrupted.
//  - maxMs: a hard ceiling that fires no matter how active the turn is, so a
//    genuinely runaway loop still terminates. Back-compat: the legacy
//    `timeoutMs` option sets this ceiling.
//  - writesPending(): reports outstanding edit calls, not past mutations.
//    No pending call does not prove the document is unchanged or clear an
//    earlier RESULT_UNKNOWN. Prior tool receipts remain authoritative.
// Defaults come from DRAFTSPECT_TURN_IDLE_MS / DRAFTSPECT_TURN_MAX_MS.

function envMs(name) {
  const value = Number(globalThis.process?.env?.[name]);
  return Number.isFinite(value) && value >= 1 ? Math.floor(value) : null;
}

export function timeoutCopy(kind, writesPending, { idleMs, maxMs }) {
  const idleSecs = Math.round(idleMs / 1000);
  const maxMins = Math.round((maxMs / 60000) * 10) / 10;
  const reason =
    kind === "idle"
      ? `no activity from the model for ${idleSecs}s`
      : `this turn exceeded the ${maxMins}-minute limit`;
  if (writesPending) {
    return `Request timed out: ${reason}. Pending Office edits may have an unknown outcome; inspect the document before retrying.`;
  }
  return `Request timed out: ${reason}, so this turn was stopped. No Office edit call is pending. Check earlier tool results before continuing; any RESULT_UNKNOWN still requires verification.`;
}

export function createLifecycle({ emit, cleanup, timeoutMs, idleMs, writesPending }) {
  const maxMs = timeoutMs ?? envMs("DRAFTSPECT_TURN_MAX_MS") ?? 900000;
  const idle = idleMs ?? envMs("DRAFTSPECT_TURN_IDLE_MS") ?? 120000;
  let handle = null;
  let closed = false;
  let idleTimer = null;
  let maxTimer = null;
  let turn = null;
  const clearTimers = () => {
    clearTimeout(idleTimer);
    clearTimeout(maxTimer);
  };
  const timeout = (kind) => {
    if (!turn || turn.done) return;
    // Conservative default: without write-tracking information, keep the
    // unknown-outcome warning rather than falsely claiming a clean document.
    const pending = typeof writesPending === "function" ? writesPending() !== false : true;
    finish({
      subtype: "timeout",
      error: timeoutCopy(kind, pending, { idleMs: idle, maxMs }),
    });
    close();
  };
  const armIdle = () => {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => timeout("idle"), idle);
    idleTimer.unref?.();
  };
  const finish = (outcome) => {
    if (!turn || turn.done) return false;
    turn.done = true;
    clearTimers();
    emit({ event: "turn_complete", turn_id: turn.id, ...outcome });
    return true;
  };
  const close = () => {
    if (closed) return;
    closed = true;
    clearTimers();
    cleanup();
    try {
      handle?.close();
    } catch {
      /* fenced even if SDK teardown throws */
    }
  };
  return {
    get closed() {
      return closed;
    },
    get active() {
      return !!turn && !turn.done;
    },
    get turnId() {
      return turn?.id ?? null;
    },
    begin(id) {
      if (closed) return false;
      if (turn && !turn.done) throw new Error("Previous turn is still running");
      turn = { id, done: false };
      armIdle();
      maxTimer = setTimeout(() => timeout("max"), maxMs);
      maxTimer.unref?.();
      return true;
    },
    activity() {
      if (!turn || turn.done || closed) return;
      armIdle();
    },
    attach(h) {
      handle = h;
      if (closed) {
        try {
          h.close();
        } catch {}
      }
    },
    finish,
    fail(error, subtype = "error", extra = {}) {
      finish({ subtype, error: String(error), ...extra });
      close();
    },
    cancel() {
      finish({ subtype: "interrupted", interrupted: true });
      close();
    },
    close,
  };
}
