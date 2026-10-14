// Timestamped console for long-running processes (daemon, Electron shell).
// daemon.log historically had per-line timestamps only on app-level markers,
// which made the 2026-09-20 turn-timeout incident impossible to reconstruct
// precisely. Install once at process entry; every console.* line gets an
// ISO-8601 prefix. Idempotent per target object (safe for tests).

const INSTALLED = Symbol("draftspect.timestampedConsole");

export function installTimestampedConsole(target = globalThis.console) {
  if (!target || target[INSTALLED]) return target;
  for (const level of ["log", "info", "warn", "error"]) {
    const original = target[level];
    if (typeof original !== "function") continue;
    target[level] = (...args) => original.call(target, new Date().toISOString(), ...args);
  }
  Object.defineProperty(target, INSTALLED, { value: true });
  return target;
}
