// Restart backoff for the Electron shell's daemon supervisor. A fixed 1s
// delay raced the dying daemon's sockets: three instant EADDRINUSE crashes
// exhausted the restart budget and left no daemon running (2026-09-20).
// Doubling the delay gives the old process time to release its ports.

export function restartDelayMs(attempt, { baseMs = 1000, maxMs = 8000 } = {}) {
  const exponent = Math.max(0, Math.floor(attempt) - 1);
  return Math.min(maxMs, baseMs * 2 ** exponent);
}
