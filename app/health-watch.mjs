// Process-local, metadata-only watchdog. It never restarts or replays work.
export function createHealthWatch({ now = Date.now, timeoutMs = 35000 } = {}) {
  let lastReply = now(),
    sequence = 0,
    acknowledged = 0,
    suspended = false;
  return {
    request() {
      return ++sequence;
    },
    reply(id) {
      if (!Number.isSafeInteger(id) || id <= acknowledged || id > sequence) return false;
      acknowledged = id;
      lastReply = now();
      return true;
    },
    suspend() {
      suspended = true;
    },
    resume() {
      suspended = false;
      lastReply = now();
      acknowledged = sequence;
    },
    unresponsive() {
      return !suspended && now() - lastReply > timeoutMs;
    },
  };
}
