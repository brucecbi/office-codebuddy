import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { createRuntimeCloseRegistry } from "../daemon/runtime/close-registry.mjs";
import { loadProcessTransport } from "./helpers/codebuddy-transport.mjs";
const ProcessTransport = await loadProcessTransport();
// The adapter reads the installed SDK version on import. Do not import it when
// the SDK is absent; SDK-independent registry tests below must still execute.
const observeCodeBuddyClose = ProcessTransport
  ? (await import("../daemon/runtime/codebuddy-process-close.mjs")).observeCodeBuddyClose
  : null;
const sdkTest = (name, fn) =>
  test(
    name,
    {
      skip:
        !ProcessTransport && "CodeBuddy SDK/version adapter absent; run npm ci for pinned checks",
    },
    fn,
  );

const controlReplies = { sendControlResponse() {}, sendControlErrorResponse() {} };
test("global shutdown awaits removed sessions and immediately closes late registrations", async () => {
  const registry = createRuntimeCloseRegistry();
  let firstExit,
    secondExit,
    closes = 0;
  registry.track(
    "removed",
    new Promise((r) => {
      firstExit = r;
    }),
    {
      close() {
        closes++;
      },
    },
  );
  const drained = registry.closeAll(100);
  registry.track(
    "late",
    new Promise((r) => {
      secondExit = r;
    }),
    {
      close() {
        closes++;
        secondExit();
      },
    },
  );
  assert.equal(closes, 2);
  let settled = false;
  drained.then(() => {
    settled = true;
  });
  await new Promise((r) => setImmediate(r));
  assert.equal(settled, false);
  firstExit();
  await drained;
  assert.equal(registry.has("removed"), false);
  assert.equal(registry.has("late"), false);
});

sdkTest(
  "CodeBuddy close barrier waits for the captured OS child exit, not sync close",
  async () => {
    const transport = {
      ...controlReplies,
      process: null,
      start() {
        this.process = spawn(
          process.execPath,
          ["-e", "process.stdin.resume(); process.stdout.write('ready');"],
          {
            stdio: ["pipe", "pipe", "pipe"],
          },
        );
      },
    };
    const session = { _transport: transport, close() {} }; // Deliberately does not stop child.
    const observer = observeCodeBuddyClose(session);
    transport.start();
    const child = transport.process;
    try {
      await once(child.stdout, "data");
      observer.close();
      let settled = false;
      const exit = observer.waitForExit().then((evidence) => {
        settled = true;
        return evidence;
      });
      await new Promise((r) => setImmediate(r));
      assert.equal(settled, false);
      child.stdin.end();
      const evidence = await exit;
      assert.equal(evidence.pid, child.pid);
      assert.equal(evidence.code, 0);
    } finally {
      if (child.exitCode === null && child.signalCode === null) child.kill();
    }
  },
);

sdkTest(
  "Stop before asynchronous CLI startup fences future spawn and proves no child",
  async () => {
    let starts = 0;
    const transport = {
      ...controlReplies,
      process: null,
      start() {
        starts++;
      },
    };
    const observer = observeCodeBuddyClose({ _transport: transport, close() {} });
    observer.close();
    assert.throws(() => transport.start(), /closed before process start/);
    assert.equal(starts, 0);
    assert.deepEqual(await observer.waitForExit(), { pid: null, process_never_started: true });
  },
);

sdkTest("SDK version/shape drift is refused before observation can be claimed", () => {
  assert.throws(() => observeCodeBuddyClose({ close() {} }), /unsupported/);
  assert.throws(
    () =>
      observeCodeBuddyClose(
        { _transport: { process: null, start() {} }, close() {} },
        { version: "future" },
      ),
    /unsupported/,
  );
});

test("pending child close survives removal of live session and timeout; actual exit permits retry", async () => {
  const registry = createRuntimeCloseRegistry();
  let exit;
  registry.track(
    "pane",
    new Promise((r) => {
      exit = r;
    }),
  );
  await assert.rejects(registry.wait("pane", 5), /exit not confirmed/);
  assert.equal(registry.has("pane"), true);
  exit();
  await registry.wait("pane");
  assert.equal(registry.has("pane"), false);
});

test("unverified child close never becomes a successful released session", async () => {
  const registry = createRuntimeCloseRegistry();
  registry.track("pane", Promise.reject(new Error("unobservable child")));
  await assert.rejects(registry.wait("pane"), /unobservable/);
  assert.equal(registry.has("pane"), true);
  assert.throws(() => registry.track("other", null), /process-exit barrier/);
});

for (const outcome of ["success", "error"]) {
  test(
    `actual SDK late MCP ${outcome} response after close cannot crash the daemon`,
    { skip: !ProcessTransport && "CodeBuddy SDK absent; run npm ci for pinned SDK checks" },
    async () => {
      const transport = new ProcessTransport();
      let resolveReply, rejectReply;
      const response = new Promise((resolve, reject) => {
        resolveReply = resolve;
        rejectReply = reject;
      });
      transport.sdkMcpTransports.set("office", { close: async () => {} });
      transport.handleMcpControlRequest = () => response;
      const observer = observeCodeBuddyClose({
        _transport: transport,
        close: () => transport.close(),
      });
      // While open, real transport errors must not be swallowed by the fence.
      assert.throws(() => transport.sendControlResponse("active", {}), /Transport not started/);
      const pending = transport.handleMcpMessageRequest({
        request_id: "pending-office",
        request: {
          server_name: "office",
          message: { jsonrpc: "2.0", method: "tools/call", id: 1 },
        },
      });
      observer.close();
      if (outcome === "success") resolveReply({ jsonrpc: "2.0", id: 1, result: {} });
      else rejectReply(new Error("Office request cancelled during Stop"));
      await assert.doesNotReject(pending);
      assert.deepEqual(await observer.waitForExit(), { pid: null, process_never_started: true });
    },
  );
}

test(
  "SDK-initiated closure also fences late control replies before runtime cleanup",
  { skip: !ProcessTransport && "CodeBuddy SDK absent; run npm ci for pinned SDK checks" },
  () => {
    const transport = new ProcessTransport();
    observeCodeBuddyClose({ _transport: transport, close: () => transport.close() });
    transport.close();
    assert.doesNotThrow(() => transport.sendControlErrorResponse("late", "cancelled"));
  },
);

sdkTest(
  "owned-process recovery retains exit barrier through ignored TERM and explicit KILL",
  async () => {
    const transport = {
      ...controlReplies,
      process: null,
      start() {
        this.process = spawn(
          process.execPath,
          [
            "-e",
            "process.on('SIGTERM',()=>{});process.stdin.resume();process.stdout.write('ready');",
          ],
          { stdio: ["pipe", "pipe", "pipe"] },
        );
      },
    };
    const observer = observeCodeBuddyClose({ _transport: transport, close() {} });
    transport.start();
    const child = transport.process;
    const unrelated = spawn(
      process.execPath,
      ["-e", "process.stdin.resume();process.stdout.write('ready');"],
      { stdio: ["pipe", "pipe", "pipe"] },
    );
    try {
      await Promise.all([once(child.stdout, "data"), once(unrelated.stdout, "data")]);
      assert.equal(observer.getCloseStatus().phase, "running");
      assert.throws(() => observer.terminateOwnedProcess({ expectedPid: child.pid }), {
        code: "RUNTIME_NOT_CLOSING",
      });
      observer.close();
      assert.equal(observer.getCloseStatus().can_terminate, true);
      const registry = createRuntimeCloseRegistry();
      registry.track("pane", observer.waitForExit(), observer);
      assert.equal(registry.status("pane")[0].pid, child.pid);
      await assert.rejects(registry.recover("other-pane", { expectedPid: child.pid }), {
        code: "RUNTIME_PROCESS_CHANGED",
      });
      await assert.rejects(registry.recover("pane", { expectedPid: unrelated.pid }), {
        code: "RUNTIME_PROCESS_CHANGED",
      });
      assert.throws(
        () => observer.terminateOwnedProcess({ expectedPid: child.pid, signal: "SIGINT" }),
        { code: "RUNTIME_SIGNAL_INVALID" },
      );
      await assert.rejects(
        registry.recover("pane", { expectedPid: child.pid, signal: "SIGTERM", timeoutMs: 20 }),
        { code: "RUNTIME_EXIT_UNCONFIRMED" },
      );
      assert.equal(registry.has("pane"), true, "sending TERM must not prove exit");
      assert.equal(child.exitCode, null);
      await assert.rejects(
        registry.recover("pane", { expectedPid: child.pid, signal: "SIGKILL" }),
        {
          code: "RUNTIME_FORCE_CONFIRMATION_REQUIRED",
        },
      );
      assert.equal(registry.has("pane"), true);
      assert.deepEqual(
        await registry.recover("pane", {
          expectedPid: child.pid,
          signal: "SIGKILL",
          confirmForce: true,
        }),
        { exited: true },
      );
      const evidence = await observer.waitForExit();
      assert.equal(evidence.pid, child.pid);
      assert.equal(evidence.signal, "SIGKILL");
      assert.equal(observer.getCloseStatus().phase, "exited");
      assert.equal(registry.has("pane"), false);
      assert.equal(unrelated.exitCode, null);
      assert.equal(unrelated.signalCode, null, "unrelated child must remain untouched");
      assert.equal(observer.terminateOwnedProcess({ expectedPid: child.pid }).signal_sent, false);
    } finally {
      for (const target of [child, unrelated]) {
        if (target.exitCode === null && target.signalCode === null) {
          const exited = once(target, "exit");
          target.kill("SIGKILL");
          await exited;
        }
      }
    }
  },
);

sdkTest(
  "process never started and unobservable controllers cannot be force-recovered",
  async () => {
    const observer = observeCodeBuddyClose({
      _transport: { ...controlReplies, process: null, start() {} },
      close() {},
    });
    observer.close();
    assert.throws(
      () =>
        observer.terminateOwnedProcess({
          expectedPid: process.pid,
          signal: "SIGKILL",
          confirmForce: true,
        }),
      { code: "RUNTIME_PROCESS_CHANGED" },
    );
    const registry = createRuntimeCloseRegistry();
    registry.track("unknown", Promise.reject(new Error("cannot observe")));
    assert.deepEqual(registry.status("unknown"), [
      { phase: "exit_unconfirmed", can_terminate: false },
    ]);
    await assert.rejects(registry.recover("unknown", { expectedPid: process.pid }), {
      code: "RUNTIME_PROCESS_CHANGED",
    });
    await assert.rejects(registry.wait("unknown"), /cannot observe/);
    assert.equal(registry.has("unknown"), true);
  },
);

test("shutdown attempts all owned closes even when one controller throws", async () => {
  const registry = createRuntimeCloseRegistry();
  let otherClosed = false;
  let finishOther;
  registry.track("broken", new Promise(() => {}), {
    close() {
      throw new Error("broken");
    },
  });
  registry.track(
    "other",
    new Promise((resolve) => {
      finishOther = resolve;
    }),
    {
      close() {
        otherClosed = true;
        finishOther();
      },
    },
  );
  await assert.rejects(registry.closeAll(20), { code: "RUNTIME_EXIT_UNCONFIRMED" });
  assert.equal(otherClosed, true);
  assert.equal(registry.has("broken"), true);
  assert.equal(registry.has("other"), false);
});
