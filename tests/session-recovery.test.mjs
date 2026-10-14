import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, rm, rename, lstat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { acquireSessionLock } from "../daemon/session-lock.mjs";
import { createSessionStore } from "../daemon/sessions.mjs";

async function fixture(fn) {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-session-recovery-"));
  const file = join(dir, "sessions.json");
  try {
    await fn({ file, lockPath: file + ".lock" });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}
const oldOwner = (pid = 2147483647) => ({ pid, at: "2000-01-01T00:00:00.000Z" });
const deadProbe = () => {
  throw Object.assign(new Error("No process"), { code: "ESRCH" });
};
const absent = async (path) => assert.rejects(lstat(path), { code: "ENOENT" });

test("session store reclaims a confirmed-dead legacy owner and persists normally", async () => {
  await fixture(async ({ file, lockPath }) => {
    await writeFile(lockPath, JSON.stringify(oldOwner()));
    const store = createSessionStore({
      file,
      legacyFile: null,
      lockOptions: { probeProcess: deadProbe },
    });
    await store.saveSessionId("word", "/synthetic", "saved", "codebuddy", "doc");
    assert.equal(await store.getSessionId("word", "/synthetic", "codebuddy", "doc"), "saved");
    await absent(lockPath);
    await absent(lockPath + ".recovery");
  });
});

test("old live owner is never expired by age", async () => {
  await fixture(async ({ lockPath }) => {
    const original = JSON.stringify(oldOwner(process.pid));
    await writeFile(lockPath, original);
    await assert.rejects(acquireSessionLock(lockPath, { timeoutMs: 10, retryMs: 1 }), {
      code: "SESSION_LOCK_BUSY",
    });
    assert.equal(await readFile(lockPath, "utf8"), original);
  });
});

test("EPERM and unknown PID probe failures keep lock intact", async () => {
  for (const code of ["EPERM", "EIO"]) {
    await fixture(async ({ lockPath }) => {
      const original = JSON.stringify(oldOwner());
      await writeFile(lockPath, original);
      await assert.rejects(
        acquireSessionLock(lockPath, {
          probeProcess: () => {
            throw Object.assign(new Error("Unavailable"), { code });
          },
        }),
        { code: "SESSION_LOCK_OWNER_UNKNOWN" },
      );
      assert.equal(await readFile(lockPath, "utf8"), original);
      await absent(lockPath + ".recovery");
    });
  }
});

test("malformed or partially written owner is fail-closed", async () => {
  for (const original of ["", "{", "{}", '{"pid":0,"at":"2020-01-01"}', '{"pid":9,"at":"bad"}']) {
    await fixture(async ({ lockPath }) => {
      await writeFile(lockPath, original);
      await assert.rejects(acquireSessionLock(lockPath, { probeProcess: deadProbe }), {
        code: "SESSION_LOCK_INVALID",
      });
      assert.equal(await readFile(lockPath, "utf8"), original);
    });
  }
});

test("concurrent dead-owner recoverers cannot unlink a new live owner", async () => {
  await fixture(async ({ lockPath }) => {
    await writeFile(lockPath, JSON.stringify(oldOwner()));
    const probeProcess = (pid) => (pid === process.pid ? undefined : deadProbe());
    let releaseReclaimer;
    let firstInspect;
    const inspected = new Promise((resolve) => {
      firstInspect = resolve;
    });
    const proceed = new Promise((resolve) => {
      releaseReclaimer = resolve;
    });
    const first = acquireSessionLock(lockPath, {
      probeProcess,
      hooks: {
        beforeReclaim: async () => {
          firstInspect();
          await proceed;
        },
      },
    });
    await inspected;
    // The timeout must outlast the first acquirer's fence so the second one
    // observes the new live owner (BUSY) rather than only the held fence.
    const second = assert.rejects(
      acquireSessionLock(lockPath, { probeProcess, retryMs: 1, timeoutMs: 500 }),
      { code: "SESSION_LOCK_BUSY" },
    );
    releaseReclaimer();
    const lock = await first;
    await second;
    const savedOwner = JSON.parse(await readFile(lockPath, "utf8"));
    assert.equal(savedOwner.nonce, lock.owner.nonce);
    await lock.release();
  });
});

test("a replaced owner is not reclaimed after the first stale snapshot", async () => {
  await fixture(async ({ lockPath }) => {
    await writeFile(lockPath, JSON.stringify(oldOwner()));
    const replacement = JSON.stringify({ ...oldOwner(process.pid), nonce: "new-owner" });
    await assert.rejects(
      acquireSessionLock(lockPath, {
        probeProcess: deadProbe,
        hooks: {
          beforeReclaim: async () => {
            await rename(lockPath, lockPath + ".original");
            await writeFile(lockPath, replacement);
          },
        },
      }),
      { code: "SESSION_LOCK_OWNERSHIP_CHANGED" },
    );
    assert.equal(await readFile(lockPath, "utf8"), replacement);
  });
});

test("release and commit assertions do not remove or accept replacement ownership", async () => {
  await fixture(async ({ lockPath }) => {
    const lock = await acquireSessionLock(lockPath);
    assert.equal(typeof lock.owner.nonce, "string");
    await rename(lockPath, lockPath + ".original");
    const replacement = JSON.stringify({ ...oldOwner(process.pid), nonce: "replacement" });
    await writeFile(lockPath, replacement);
    await assert.rejects(lock.assertOwned(), { code: "SESSION_LOCK_OWNERSHIP_CHANGED" });
    await assert.rejects(lock.release(), { code: "SESSION_LOCK_OWNERSHIP_CHANGED" });
    assert.equal(await readFile(lockPath, "utf8"), replacement);
  });
});

test("a live recovery fence is reported as busy and is not removed", async () => {
  await fixture(async ({ lockPath }) => {
    const held = JSON.stringify({
      pid: process.pid,
      nonce: "live-fence",
      at: new Date().toISOString(),
    });
    await writeFile(lockPath + ".recovery", held);
    await assert.rejects(
      acquireSessionLock(lockPath, {
        timeoutMs: 30,
        retryMs: 5,
        probeProcess: (pid) => {
          if (pid !== process.pid) deadProbe();
        },
      }),
      { code: "SESSION_LOCK_BUSY" },
    );
    assert.equal(await readFile(lockPath + ".recovery", "utf8"), held);
    await absent(lockPath);
  });
});

test("an unverifiable young fence stays busy; an old one asks for controlled recovery", async () => {
  await fixture(async ({ lockPath }) => {
    const denied = () => {
      throw Object.assign(new Error("Unavailable"), { code: "EPERM" });
    };
    const young = JSON.stringify({ pid: 2147483646, nonce: "young", at: new Date().toISOString() });
    await writeFile(lockPath + ".recovery", young);
    await assert.rejects(
      acquireSessionLock(lockPath, { timeoutMs: 40, retryMs: 5, probeProcess: denied }),
      { code: "SESSION_LOCK_BUSY" },
    );
    assert.equal(await readFile(lockPath + ".recovery", "utf8"), young);
    const old = JSON.stringify({ pid: 2147483646, nonce: "old", at: "2000-01-01T00:00:00.000Z" });
    await writeFile(lockPath + ".recovery", old);
    await assert.rejects(
      acquireSessionLock(lockPath, { timeoutMs: 40, retryMs: 5, probeProcess: denied }),
      { code: "SESSION_LOCK_RECOVERY_BLOCKED" },
    );
    assert.equal(await readFile(lockPath + ".recovery", "utf8"), old);
  });
});

test("interrupted recovery fence stays fail-closed instead of recursively evicting it", async () => {
  await fixture(async ({ lockPath }) => {
    const original = JSON.stringify(oldOwner());
    await writeFile(lockPath + ".recovery", original);
    await assert.rejects(
      acquireSessionLock(lockPath, { timeoutMs: 5, retryMs: 1, probeProcess: deadProbe }),
      { code: "SESSION_LOCK_RECOVERY_BLOCKED" },
    );
    assert.equal(await readFile(lockPath + ".recovery", "utf8"), original);
    await absent(lockPath);
  });
});

test("transaction failure releases its own lock, leaves original bytes, and later work can recover", async () => {
  await fixture(async ({ file, lockPath }) => {
    const store = createSessionStore({ file, legacyFile: null });
    const original = "corrupt source must survive";
    await writeFile(file, original);
    await assert.rejects(store.saveSessionId("word", "/synthetic", "not-saved"), SyntaxError);
    assert.equal(await readFile(file, "utf8"), original);
    await absent(lockPath);
    await absent(lockPath + ".recovery");
    await writeFile(file, JSON.stringify({ version: 4, folders: {} }));
    await store.saveSessionId("word", "/synthetic", "saved");
    assert.equal(await store.getSessionId("word", "/synthetic"), "saved");
  });
});

test("parallel stores preserve independent session updates after stale recovery", async () => {
  await fixture(async ({ file, lockPath }) => {
    await writeFile(lockPath, JSON.stringify(oldOwner()));
    const stores = Array.from({ length: 8 }, () =>
      createSessionStore({
        file,
        legacyFile: null,
        lockOptions: {
          retryMs: 1,
          probeProcess: (pid) => (pid === process.pid ? undefined : deadProbe()),
        },
      }),
    );
    await Promise.all(
      stores.map((store, i) =>
        store.saveSessionId("word", "/synthetic", "s" + i, "codebuddy", "doc" + i),
      ),
    );
    for (let i = 0; i < stores.length; i++)
      assert.equal(
        await stores[i].getSessionId("word", "/synthetic", "codebuddy", "doc" + i),
        "s" + i,
      );
    await absent(lockPath);
    await absent(lockPath + ".recovery");
  });
});

test("ownership replacement during a transaction refuses commit and preserves both new lock and previous state", async () => {
  await fixture(async ({ file, lockPath }) => {
    const original = JSON.stringify({ version: 4, folders: {} });
    await writeFile(file, original);
    const replacement = JSON.stringify({ ...oldOwner(process.pid), nonce: "new-lock-owner" });
    const store = createSessionStore({
      file,
      legacyFile: null,
      lockOptions: {
        hooks: {
          afterAcquire: async () => {
            await rename(lockPath, lockPath + ".original");
            await writeFile(lockPath, replacement);
          },
        },
      },
    });
    await assert.rejects(store.saveSessionId("word", "/synthetic", "not-saved"), {
      code: "SESSION_LOCK_OWNERSHIP_CHANGED",
    });
    assert.equal(await readFile(file, "utf8"), original);
    assert.equal(await readFile(lockPath, "utf8"), replacement);
  });
});

test("isolated process-crash E2E: actual dead child lock recovers and competing processes preserve updates", async () => {
  await fixture(async ({ file, lockPath }) => {
    const lockUrl = new URL("../daemon/session-lock.mjs", import.meta.url).href;
    const storeUrl = new URL("../daemon/sessions.mjs", import.meta.url).href;
    const child = spawn(
      process.execPath,
      [
        "--input-type=module",
        "-e",
        `
      import { acquireSessionLock } from ${JSON.stringify(lockUrl)};
      await acquireSessionLock(process.argv[1]);
      process.stdin.resume();
      process.stdout.write("ready");
    `,
        lockPath,
      ],
      { stdio: ["pipe", "pipe", "pipe"] },
    );
    try {
      await once(child.stdout, "data");
      assert.equal(JSON.parse(await readFile(lockPath, "utf8")).pid, child.pid);
      const exited = once(child, "exit");
      child.kill("SIGKILL"); // Only this test's own synthetic child.
      await exited;
      assert.throws(() => process.kill(child.pid, 0), { code: "ESRCH" });
      const jobs = Array.from({ length: 5 }, (_, i) => {
        const writer = spawn(
          process.execPath,
          [
            "--input-type=module",
            "-e",
            `
          import { createSessionStore } from ${JSON.stringify(storeUrl)};
          const store = createSessionStore({ file: process.argv[1], legacyFile: null });
          await store.saveSessionId("word", "/synthetic", "session-" + process.argv[2], "codebuddy", "doc-" + process.argv[2]);
        `,
            file,
            String(i),
          ],
          { stdio: ["ignore", "ignore", "pipe"] },
        );
        let stderr = "";
        writer.stderr.on("data", (data) => {
          stderr += data;
        });
        return once(writer, "exit").then(([code, signal]) => {
          assert.equal(code, 0, stderr);
          assert.equal(signal, null);
        });
      });
      await Promise.all(jobs);
      const store = createSessionStore({ file, legacyFile: null });
      for (let i = 0; i < jobs.length; i++)
        assert.equal(
          await store.getSessionId("word", "/synthetic", "codebuddy", "doc-" + i),
          "session-" + i,
        );
      await absent(lockPath);
      await absent(lockPath + ".recovery");
    } finally {
      if (child.exitCode === null && child.signalCode === null) {
        const exited = once(child, "exit");
        child.kill("SIGKILL");
        await exited;
      }
    }
  });
});
