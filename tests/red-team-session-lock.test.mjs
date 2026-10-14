import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, readFile, unlink, rm, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { acquireSessionLock } from "../daemon/session-lock.mjs";

const dead = () => {
  throw Object.assign(new Error("No such process"), { code: "ESRCH" });
};
const denied = () => {
  throw Object.assign(new Error("Permission denied"), { code: "EPERM" });
};
const owner = (pid = 2147483647) => ({ pid, at: "2000-01-01T00:00:00Z", nonce: "original-owner" });
async function isolated(fn) {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-red-lock-"));
  try {
    await fn(join(dir, "sessions.json.lock"), dir);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

test("RED age never evicts live or unverifiable lock owners", async () => {
  await isolated(async (path) => {
    const bytes = JSON.stringify(owner(process.pid));
    await writeFile(path, bytes);
    await assert.rejects(
      acquireSessionLock(path, { timeoutMs: 15, retryMs: 2, probeProcess: () => {} }),
      (e) => e.code === "SESSION_LOCK_BUSY",
    );
    assert.equal(await readFile(path, "utf8"), bytes);
    await assert.rejects(
      acquireSessionLock(path, { timeoutMs: 15, probeProcess: denied }),
      (e) => e.code === "SESSION_LOCK_OWNER_UNKNOWN",
    );
    assert.equal(await readFile(path, "utf8"), bytes);
  });
});

test("RED malformed and symlink lock metadata cannot be reclaimed as dead", async () => {
  await isolated(async (path, dir) => {
    for (const bytes of [
      "{",
      JSON.stringify({ pid: 0, at: "2000-01-01" }),
      JSON.stringify({ pid: 1, at: "unknown" }),
    ]) {
      await writeFile(path, bytes);
      await assert.rejects(
        acquireSessionLock(path, { timeoutMs: 10, probeProcess: dead }),
        (e) => e.code === "SESSION_LOCK_INVALID",
      );
      assert.equal(await readFile(path, "utf8"), bytes);
    }
    const target = join(dir, "unrelated.txt");
    await writeFile(target, "preserve");
    await unlink(path);
    await symlink(target, path);
    await assert.rejects(
      acquireSessionLock(path, { timeoutMs: 10, probeProcess: dead }),
      (e) => e.code === "SESSION_LOCK_INVALID",
    );
    assert.equal(await readFile(target, "utf8"), "preserve");
  });
});

test("RED concurrent reclaimers preserve the new owner rather than deleting it", async () => {
  await isolated(async (path) => {
    await writeFile(path, JSON.stringify(owner()));
    let entered, releaseFence;
    const atFence = new Promise((r) => (entered = r)),
      gate = new Promise((r) => (releaseFence = r));
    const first = acquireSessionLock(path, {
      timeoutMs: 500,
      retryMs: 1,
      probeProcess: dead,
      hooks: {
        beforeReclaim: async () => {
          entered();
          await gate;
        },
      },
    });
    await atFence;
    const second = acquireSessionLock(path, {
      timeoutMs: 500,
      retryMs: 1,
      probeProcess: (pid) => {
        if (pid !== process.pid) dead();
      },
    });
    releaseFence();
    const one = await first;
    await new Promise((r) => setTimeout(r, 20));
    assert.equal(JSON.parse(await readFile(path, "utf8")).nonce, one.owner.nonce);
    await one.release();
    const two = await second;
    assert.notEqual(one.owner.nonce, two.owner.nonce);
    await two.assertOwned();
    await two.release();
  });
});

test("RED replacement during reclaim or release cannot be erased by the old owner", async () => {
  await isolated(async (path) => {
    await writeFile(path, JSON.stringify(owner()));
    const replacement = JSON.stringify({ ...owner(process.pid), nonce: "replacement" });
    await assert.rejects(
      acquireSessionLock(path, {
        timeoutMs: 20,
        probeProcess: dead,
        hooks: {
          beforeReclaim: async () => {
            await unlink(path);
            await writeFile(path, replacement);
          },
        },
      }),
      (e) => e.code === "SESSION_LOCK_OWNERSHIP_CHANGED",
    );
    assert.equal(await readFile(path, "utf8"), replacement);
    await unlink(path);
    const held = await acquireSessionLock(path);
    await unlink(path);
    await writeFile(path, replacement);
    await assert.rejects(held.release(), (e) => e.code === "SESSION_LOCK_OWNERSHIP_CHANGED");
    assert.equal(await readFile(path, "utf8"), replacement);
  });
});

test("RED a crashed recovery fence fails closed even when its PID is dead", async () => {
  await isolated(async (path) => {
    await writeFile(path + ".recovery", JSON.stringify(owner()));
    await assert.rejects(
      acquireSessionLock(path, { timeoutMs: 10, retryMs: 2, probeProcess: dead }),
      (e) => e.code === "SESSION_LOCK_RECOVERY_BLOCKED",
    );
    assert.equal(JSON.parse(await readFile(path + ".recovery", "utf8")).nonce, "original-owner");
  });
});
