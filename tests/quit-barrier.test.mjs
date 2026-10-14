import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { spawn } from "node:child_process";
import { waitForOwnedExit } from "../app/quit-barrier.mjs";

function childFixture() {
  const child = new EventEmitter();
  child.pid = 424242;
  child.exitCode = null;
  child.signalCode = null;
  child.kill = () => {
    throw new Error("The barrier must never kill a process");
  };
  return child;
}

function assertClean(child) {
  assert.equal(child.listenerCount("exit"), 0);
  assert.equal(child.listenerCount("error"), 0);
}

function assertFailure(code) {
  return (error) => {
    assert.equal(error.code, code);
    assert.ok(!error.message.includes("synthetic-secret"));
    assert.equal(error.cause, undefined);
    return true;
  };
}

test("subscribes before Stop and captures synchronous exit before the mutable owner is cleared", async () => {
  const child = childFixture();
  let owner = child;
  let stops = 0;
  await waitForOwnedExit(owner, () => {
    stops++;
    assert.equal(child.listenerCount("exit"), 1);
    child.emit("exit", 0, null);
    owner = null;
  });
  assert.equal(stops, 1);
  assert.equal(owner, null);
  assertClean(child);
});

test("does not complete until the captured child asynchronously exits", async () => {
  const child = childFixture();
  let resolved = false;
  const waiting = waitForOwnedExit(child, () => {}).then(() => {
    resolved = true;
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(resolved, false);
  child.emit("exit", 0, null);
  await waiting;
  assert.equal(resolved, true);
  assertClean(child);
});

test("already-exited and null children still request Stop to cancel pending starts", async () => {
  for (const mode of ["code", "signal", "null"]) {
    const child = mode === "null" ? null : childFixture();
    if (mode === "code") child.exitCode = 0;
    if (mode === "signal") child.signalCode = "SIGTERM";
    let stops = 0;
    await waitForOwnedExit(
      child,
      () => {
        stops++;
      },
      { timeoutMs: 100 },
    );
    assert.equal(stops, 1);
    if (child) assertClean(child);
  }
});

test("a killed flag alone is not exit evidence; timeout never signals processes", async () => {
  const child = childFixture();
  child.killed = true;
  const other = childFixture();
  let ownedSignals = 0;
  let otherSignals = 0;
  child.kill = () => {
    ownedSignals++;
  };
  other.kill = () => {
    otherSignals++;
  };
  let stops = 0;
  await assert.rejects(
    waitForOwnedExit(
      child,
      () => {
        stops++;
      },
      { timeoutMs: 15 },
    ),
    assertFailure("DESKTOP_STOP_TIMEOUT"),
  );
  assert.equal(stops, 1);
  assert.equal(ownedSignals, 0);
  assert.equal(otherSignals, 0);
  assertClean(child);
  assertClean(other);
  // The former captured child may exit after the timeout without triggering
  // any barrier restart/quit callback or keeping a listener alive.
  child.emit("exit", 0, null);
  assertClean(child);
});

test("synchronous Stop exceptions are sanitized and listeners are removed", async () => {
  const child = childFixture();
  await assert.rejects(
    waitForOwnedExit(child, () => {
      throw new Error("synthetic-secret /private/config");
    }),
    assertFailure("DESKTOP_STOP_FAILED"),
  );
  assertClean(child);
});

test("an exit emitted just before Stop throws does not hide the stop failure", async () => {
  const child = childFixture();
  await assert.rejects(
    waitForOwnedExit(child, () => {
      child.emit("exit", 0, null);
      throw new Error("synthetic-secret");
    }),
    assertFailure("DESKTOP_STOP_FAILED"),
  );
  assertClean(child);
});

test("asynchronous Stop rejection is sanitized and listeners are removed", async () => {
  const child = childFixture();
  await assert.rejects(
    waitForOwnedExit(child, async () => {
      throw new Error("synthetic-secret");
    }),
    assertFailure("DESKTOP_STOP_FAILED"),
  );
  assertClean(child);
});

test("an error with a live PID cannot falsely report that the child exited", async () => {
  const child = childFixture();
  await assert.rejects(
    waitForOwnedExit(
      child,
      () => {
        child.emit("error", new Error("synthetic-secret kill failure"));
      },
      { timeoutMs: 15 },
    ),
    assertFailure("DESKTOP_STOP_TIMEOUT"),
  );
  assertClean(child);
});

test("a no-PID spawn error completes without waiting for a nonexistent exit", async () => {
  const child = childFixture();
  child.pid = undefined;
  await waitForOwnedExit(child, () => {
    child.emit("error", new Error("synthetic-secret spawn ENOENT"));
  });
  assertClean(child);
});

test("real failed spawn has no PID and its error permits clean completion", async () => {
  const child = spawn("/draftspect-synthetic-not-existing-quit-test/program", [], {
    stdio: "ignore",
  });
  let stops = 0;
  await waitForOwnedExit(
    child,
    () => {
      stops++;
    },
    { timeoutMs: 500 },
  );
  assert.equal(child.pid, undefined);
  assert.equal(stops, 1);
  assertClean(child);
});

test("cleanup preserves unrelated owner exit/error listeners", async () => {
  const child = childFixture();
  const exitHandler = () => {};
  const errorHandler = () => {};
  child.on("exit", exitHandler);
  child.on("error", errorHandler);
  await waitForOwnedExit(child, () => child.emit("exit", 0, null));
  assert.deepEqual(child.listeners("exit"), [exitHandler]);
  assert.deepEqual(child.listeners("error"), [errorHandler]);
});
