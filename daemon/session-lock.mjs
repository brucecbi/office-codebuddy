// Installation-local session-store lock. All acquisition/recovery attempts use
// a short exclusive fence, so two dead-owner reclaimers cannot unlink a newer
// owner's lock. We deliberately never reap the fence itself: a crash in that
// short critical section needs controlled recovery with this application shut
// down, rather than another recursively unsafe stale-lock algorithm. A live
// holder, or a fence we cannot yet prove was abandoned, is reported as busy.
const FENCE_STALE_MS = 15_000;
import { open, lstat, unlink } from "node:fs/promises";
import { constants } from "node:fs";
import { randomUUID } from "node:crypto";

export class SessionLockError extends Error {
  constructor(code, message, recovery = {}) {
    super(message);
    this.name = "SessionLockError";
    this.code = code;
    // Only operational metadata. No file paths, document names or session IDs.
    this.recovery = recovery;
  }
}

const sameFile = (a, b) => a?.dev === b?.dev && a?.ino === b?.ino;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const lostOwnership = () =>
  new SessionLockError(
    "SESSION_LOCK_OWNERSHIP_CHANGED",
    "Session lock ownership changed; persistence was stopped. Keep the lock and inspect the store before retrying.",
  );

async function snapshot(path) {
  let handle;
  try {
    const entry = await lstat(path);
    if (!entry.isFile() || entry.isSymbolicLink() || entry.size > 4096)
      throw new SessionLockError(
        "SESSION_LOCK_INVALID",
        "Session lock metadata is invalid; controlled recovery is required.",
      );
    handle = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW);
    const stat = await handle.stat();
    if (!sameFile(entry, stat)) throw lostOwnership();
    const raw = await handle.readFile("utf8");
    let owner;
    try {
      owner = JSON.parse(raw);
    } catch {}
    if (
      !owner ||
      !Number.isSafeInteger(owner.pid) ||
      owner.pid <= 0 ||
      typeof owner.at !== "string" ||
      !Number.isFinite(Date.parse(owner.at)) ||
      (owner.nonce !== undefined && (typeof owner.nonce !== "string" || !owner.nonce))
    )
      throw new SessionLockError(
        "SESSION_LOCK_INVALID",
        "Session lock metadata is incomplete or invalid; it was not removed.",
      );
    return { stat, raw, owner };
  } catch (error) {
    if (error.code === "ENOENT") return null;
    throw error;
  } finally {
    await handle?.close();
  }
}

async function ownedFile(path) {
  const handle = await open(path, "wx", 0o600);
  let stat;
  const owner = { pid: process.pid, nonce: randomUUID(), at: new Date().toISOString() };
  const raw = JSON.stringify(owner);
  try {
    stat = await handle.stat();
    await handle.writeFile(raw);
  } catch (error) {
    // The exclusive create succeeded, so only remove our still-identical inode.
    try {
      const current = await lstat(path).catch(() => null);
      if (stat && sameFile(stat, current)) await unlink(path);
    } finally {
      await handle.close();
    }
    throw error;
  }
  let released = false;
  const assertOwned = async () => {
    if (released) throw lostOwnership();
    const current = await snapshot(path);
    if (
      !current ||
      !sameFile(stat, current.stat) ||
      current.raw !== raw ||
      current.owner.nonce !== owner.nonce
    )
      throw lostOwnership();
  };
  return {
    owner,
    assertOwned,
    async release() {
      if (released) return;
      try {
        await assertOwned();
        await unlink(path);
      } finally {
        released = true;
        await handle.close();
      }
    },
  };
}

function busyLock(
  message = "Session store is locked by another process; retry after it finishes.",
  recovery = {},
) {
  return new SessionLockError("SESSION_LOCK_BUSY", message, recovery);
}

function blockedFence() {
  return new SessionLockError(
    "SESSION_LOCK_RECOVERY_BLOCKED",
    "Session lock recovery is already in progress or was interrupted. Retry; if it persists, stop this application's daemons before controlled recovery.",
    { phase: "recovery_fence", requires_stopped_application: true },
  );
}

// A held fence is not evidence of a crash by itself. Say "busy" while its
// owner is alive, or while an unverifiable fence is still young. The
// recovery message is only for a dead owner, or an unverifiable fence that
// has already been around longer than a normal critical section.
async function classifyHeldFence(fencePath, probeProcess, sightings) {
  let seen;
  try {
    seen = await snapshot(fencePath);
  } catch (error) {
    if (error.code === "SESSION_LOCK_OWNERSHIP_CHANGED") return busyLock();
    if (error.code !== "SESSION_LOCK_INVALID") throw error;
    const entry = await lstat(fencePath).catch(() => null);
    const age = entry ? Date.now() - entry.mtimeMs : 0;
    return sightings > 1 && age >= FENCE_STALE_MS ? blockedFence() : busyLock();
  }
  if (!seen) return busyLock();
  const state = ownerState(seen.owner, probeProcess);
  if (state === "alive") return busyLock();
  if (state === "dead") return blockedFence();
  const age = Date.now() - Date.parse(seen.owner.at);
  return sightings > 1 && age >= FENCE_STALE_MS ? blockedFence() : busyLock();
}

function ownerState(owner, probeProcess) {
  try {
    probeProcess(owner.pid);
    return "alive";
  } catch (error) {
    // EPERM and every other error are uncertainty, never evidence of death.
    return error.code === "ESRCH" ? "dead" : "unknown";
  }
}

export async function acquireSessionLock(
  lockPath,
  { timeoutMs = 2000, retryMs = 25, probeProcess = (pid) => process.kill(pid, 0), hooks = {} } = {},
) {
  const deadline = Date.now() + timeoutMs;
  const fencePath = lockPath + ".recovery";
  let fenceSightings = 0;
  let last = busyLock();
  const finishAcquire = async () => {
    const lock = await ownedFile(lockPath);
    try {
      await hooks.afterAcquire?.(lock);
      return lock;
    } catch (error) {
      await lock.release();
      throw error;
    }
  };
  for (;;) {
    let fence;
    try {
      fence = await ownedFile(fencePath);
    } catch (error) {
      if (error.code !== "EEXIST") throw error;
      fenceSightings++;
      last = await classifyHeldFence(fencePath, probeProcess, fenceSightings);
    }
    if (fence) {
      try {
        const previous = await snapshot(lockPath);
        if (previous) {
          const state = ownerState(previous.owner, probeProcess);
          if (state === "unknown")
            throw new SessionLockError(
              "SESSION_LOCK_OWNER_UNKNOWN",
              "The session lock owner's status cannot be verified; the lock was kept.",
              { phase: "owner_probe", owner_pid: previous.owner.pid },
            );
          if (state === "alive") {
            last = busyLock("Session store is locked by a live process; retry after it finishes.", {
              owner_pid: previous.owner.pid,
            });
          } else {
            await hooks.beforeReclaim?.(previous);
            const current = await snapshot(lockPath);
            if (!current || !sameFile(previous.stat, current.stat) || previous.raw !== current.raw)
              throw lostOwnership();
            await fence.assertOwned();
            await unlink(lockPath);
          }
          // A live owner is never evicted, however old its timestamp is.
          if (state !== "alive") return await finishAcquire();
        } else {
          return await finishAcquire();
        }
      } finally {
        await fence.release();
      }
    }
    if (Date.now() >= deadline) throw last;
    await sleep(Math.min(retryMs, Math.max(1, deadline - Date.now())));
  }
}
