// Retain teardown independently of the UI lifecycle: Stop fences/removes a
// live session immediately, but a later archive/resume still has to await exit.
export function createRuntimeCloseRegistry() {
  const pending = new Map();
  let closing = false;
  const registry = {
    has: (key) => pending.has(key),
    track(key, closed, controller = null) {
      if (!closed || typeof closed.then !== "function")
        throw new Error("CodeBuddy runtime did not provide a process-exit barrier");
      const group = pending.get(key) || new Set();
      pending.set(key, group);
      const entry = { closed, controller };
      group.add(entry);
      closed.then(
        () => {
          group.delete(entry);
          if (!group.size && pending.get(key) === group) pending.delete(key);
        },
        () => {}, // An unverified close remains a blocking entry, never a pass.
      );
      if (closing) controller?.close?.();
    },
    async closeAll(timeoutMs = 2500) {
      closing = true;
      // Includes removed/released UI sessions, not just the current session map.
      const deadline = Date.now() + timeoutMs;
      for (const group of pending.values())
        for (const { controller } of group) {
          try {
            controller?.close?.();
          } catch {
            // One broken close must not skip the other owned children. The
            // actual-exit barrier below still fails if this child remains.
          }
        }
      while (pending.size)
        await Promise.all(
          [...pending.keys()].map((key) => registry.wait(key, Math.max(1, deadline - Date.now()))),
        );
    },
    status(key) {
      return [...(pending.get(key) || [])].map(({ controller }) => {
        try {
          return (
            controller?.getCloseStatus?.() || { phase: "exit_unconfirmed", can_terminate: false }
          );
        } catch {
          return { phase: "unobservable", can_terminate: false };
        }
      });
    },
    async recover(
      key,
      { expectedPid, signal = "SIGTERM", confirmForce = false, timeoutMs = 5000 } = {},
    ) {
      const matches = [...(pending.get(key) || [])].filter(
        ({ controller }) =>
          controller?.getCloseStatus?.().pid === expectedPid && Number.isSafeInteger(expectedPid),
      );
      if (
        matches.length !== 1 ||
        typeof matches[0].controller?.terminateOwnedProcess !== "function"
      )
        throw Object.assign(
          new Error(
            "No unique captured runtime process matches; refresh the session status before recovery.",
          ),
          { code: "RUNTIME_PROCESS_CHANGED" },
        );
      matches[0].controller.terminateOwnedProcess({ expectedPid, signal, confirmForce });
      await registry.wait(key, timeoutMs);
      return { exited: true };
    },
    async wait(key, timeoutMs = 5000) {
      let timer;
      try {
        await Promise.race([
          (async () => {
            while (pending.has(key))
              await Promise.all([...pending.get(key)].map(({ closed }) => closed));
          })(),
          new Promise((_, reject) => {
            timer = setTimeout(
              () =>
                reject(
                  Object.assign(
                    new Error(
                      "Runtime process exit not confirmed; keep the session released and retry after it exits.",
                    ),
                    { code: "RUNTIME_EXIT_UNCONFIRMED" },
                  ),
                ),
              timeoutMs,
            );
          }),
        ]);
      } finally {
        clearTimeout(timer);
      }
    },
  };
  return registry;
}
