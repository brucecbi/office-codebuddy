// Keep the Electron owner alive while its captured daemon finishes exiting.
// This helper never kills a process, releases an ownership lock, or restarts.
function stopFailure(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function hasExited(child) {
  return !child || child.exitCode != null || child.signalCode != null;
}

export async function waitForOwnedExit(child, requestStop, { timeoutMs = 5_000 } = {}) {
  return new Promise((resolve, reject) => {
    let settled = false;
    let stopCompleted = false;
    let exited = hasExited(child);
    const duration = Number.isFinite(timeoutMs) && timeoutMs >= 0 ? timeoutMs : 5_000;
    const cleanup = () => {
      clearTimeout(timer);
      child?.removeListener("exit", onExit);
      child?.removeListener("error", onError);
    };
    const finish = (error) => {
      if (settled) return;
      settled = true;
      cleanup();
      if (error) reject(error);
      else resolve();
    };
    const finishIfStopped = () => {
      if (stopCompleted && exited) finish();
    };
    const onExit = () => {
      exited = true;
      finishIfStopped();
    };
    const onError = () => {
      // A failed spawn has no process to await. Other child errors (including
      // a failed kill/send) do not establish that a live PID has exited.
      if (child.pid == null) onExit();
    };
    const timer = setTimeout(() => {
      finish(
        stopFailure(
          "DESKTOP_STOP_TIMEOUT",
          "The owned local service has not exited. The application must stay open; retry Quit after checking its status. No process was force-stopped.",
        ),
      );
    }, duration);

    // Install these before Stop: an already queued or synchronous exit must
    // not be lost when Stop clears the caller's mutable daemon reference.
    child?.on("exit", onExit);
    child?.on("error", onError);
    try {
      // Always cancel pending starts/backoff, including when no child exists.
      const requested = requestStop();
      Promise.resolve(requested).then(
        () => {
          stopCompleted = true;
          exited ||= hasExited(child);
          finishIfStopped();
        },
        () =>
          finish(
            stopFailure(
              "DESKTOP_STOP_FAILED",
              "The local service stop request failed. The application must stay open; check its status before retrying Quit.",
            ),
          ),
      );
    } catch {
      // Do not expose raw exceptions: a caller may include paths or secrets.
      finish(
        stopFailure(
          "DESKTOP_STOP_FAILED",
          "The local service stop request failed. The application must stay open; check its status before retrying Quit.",
        ),
      );
    }
  });
}
