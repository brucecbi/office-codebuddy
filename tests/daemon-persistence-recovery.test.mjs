// Actual daemon functions in a synthetic environment: no server/CLI/model,
// user settings, real persistence state, or Word document is opened.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { createRuntimeCloseRegistry } from "../daemon/runtime/close-registry.mjs";
import { createSessionPersistenceRegistry } from "../daemon/session-persistence.mjs";
import { randomUUID } from "node:crypto";
import { sameTranscriptFingerprint } from "../daemon/session-transcript-state.mjs";

const tick = () => new Promise((resolve) => setImmediate(resolve));
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};

async function harness({ save = async () => {}, active = null } = {}) {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const prelude = source.slice(
    source.indexOf("const sessions = new Map();"),
    source.indexOf("// Static fallback"),
  );
  const init = source.slice(
    source.indexOf("function handleAgentMessage("),
    source.indexOf(
      "\n// ---------------------------------------------------------------------------\n// Kick off.",
    ),
  );
  const resolve = source.slice(
    source.indexOf("async function resolveReplaySessionId("),
    source.indexOf("// Reconstruct this pane's prior conversation"),
  );
  const rpcs = source.slice(
    source.indexOf("    get_session_info: async"),
    source.indexOf("    get_context: async"),
  );
  let savedId = "prior-id",
    handoff = null;
  const events = [],
    replies = [],
    order = [],
    savedCalls = [];
  const closes = createRuntimeCloseRegistry();
  const realWait = closes.wait.bind(closes);
  closes.wait = (key) => realWait(key, 15);
  const context = vm.createContext({
    Date,
    randomUUID,
    captureTranscriptFingerprint: async () => ({
      path: "/synthetic/id.jsonl",
      dev: 1,
      ino: 1,
      size: 1,
      mtime_ms: 1,
      ctime_ms: 1,
      sha256: "fixture",
    }),
    sameTranscriptFingerprint,
    Promise,
    setTimeout,
    clearTimeout,
    console: { log() {}, warn() {}, error() {} },
    createRuntimeCloseRegistry: () => closes,
    createSessionPersistenceRegistry,
    runtime: { id: "codebuddy" },
    diag() {},
    bridge: {
      activeTurn: () => active,
      sendAssistantEvent: (event, key) => events.push({ ...event, key }),
    },
    saveSessionId: async (...args) => {
      savedCalls.push(args);
      await save(...args);
      savedId = args[2];
    },
    getSessionId: async () => savedId,
    getHandoff: async () => handoff,
    setHandoff: async (_host, _cwd, _provider, _key, value) => {
      handoff = value;
      order.push(value ? "release" : "return");
    },
    resumeCommands: () => ({ resume: "synthetic-resume" }),
    startQueue: new Map(),
    archiveSession: async (info) => {
      order.push("archive:" + info.session_id);
      return { created_at: "fixture-time", archive_path: "/synthetic/archive" };
    },
    sendTranscriptReplayTo: async () => order.push("replay"),
  });
  vm.runInContext(
    prelude +
      `
    function sessionFor(key) { return sessions.get(key) || null; }
    function cwdForKey(key) { return sessionFor(key)?.cwd || "/synthetic"; }
    ` +
      resolve +
      init +
      `
    globalThis.handlers = {${rpcs}};
    globalThis.info = sessionInfo;
    globalThis.init = handleAgentMessage;
    globalThis.liveSessions = sessions;
    globalThis.handoffBusy = handoffInProgress;
  `,
    context,
  );
  const session = {
    key: "pane",
    host: "word",
    cwd: "/synthetic",
    sessionId: null,
    lifecycle: {
      active: false,
      fail(message) {
        throw new Error(message);
      },
    },
    cancel() {
      context.liveSessions.delete("pane");
      order.push("cancel");
    },
  };
  context.liveSessions.set("pane", session);
  return {
    events,
    replies,
    order,
    savedCalls,
    closes,
    session,
    init(id = "latest-id") {
      context.init(
        { type: "system", subtype: "init", model: "synthetic", session_id: id },
        session,
      );
    },
    stop() {
      context.liveSessions.delete("pane");
    },
    busy(value) {
      active = value;
    },
    handoffBusy(value) {
      if (value) context.handoffBusy.add("pane");
      else context.handoffBusy.delete("pane");
    },
    info: () => context.info("pane", "word"),
    call: (name, message = {}) =>
      context.handlers[name](
        { request_id: "test-request", ...message },
        (reply) => replies.push(reply),
        "pane",
        "word",
      ),
  };
}

test("actual init save failure emits a visible code and blocks archive even after Stop removed live session", async () => {
  const save = deferred();
  const h = await harness({ save: () => save.promise });
  h.init();
  h.stop();
  save.reject(
    Object.assign(new Error("Synthetic disk denial: private path must not be emitted"), {
      code: "EACCES",
    }),
  );
  await tick();
  assert.deepEqual(
    h.events.filter((event) => event.event === "persistence_error"),
    [{ event: "persistence_error", code: "EACCES", key: "pane" }],
  );
  assert.equal((await h.info()).persistence_status, "failed");
  await assert.rejects(h.call("release_session"), { code: "EACCES" });
  assert.deepEqual(h.order, []);
});

test("actual archive waits pending persistence after Stop, then uses the newly saved ID", async () => {
  const save = deferred();
  const h = await harness({ save: () => save.promise });
  h.init();
  h.stop();
  const releasing = h.call("release_session");
  await tick();
  assert.deepEqual(h.order, []);
  assert.equal((await h.info()).persistence_status, "saving");
  save.resolve();
  await releasing;
  assert.ok(h.order.includes("archive:latest-id"));
  assert.ok(!h.order.includes("archive:prior-id"));
  assert.equal(h.replies.at(-1).ok, true);
});

test("actual save retry reuses failed identity after Stop and restores archiving without new model session", async () => {
  let fail = true;
  const h = await harness({
    save: async () => {
      if (fail) throw Object.assign(new Error("locked"), { code: "SESSION_LOCK_BUSY" });
    },
  });
  h.init();
  h.stop();
  await tick();
  assert.equal((await h.info()).persistence_status, "failed");
  fail = false;
  await h.call("retry_session_save");
  assert.equal(h.savedCalls.length, 2);
  assert.deepEqual(h.savedCalls[0], h.savedCalls[1]);
  assert.equal((await h.info()).persistence_status, "saved");
  await h.call("release_session");
  assert.ok(h.order.includes("archive:latest-id"));
});

test("actual recovery RPC rejects missing confirmation, active turn and concurrent handoff before signaling", async () => {
  const h = await harness();
  const exit = deferred();
  let signals = 0;
  h.closes.track("pane", exit.promise, {
    getCloseStatus: () => ({ pid: 12345, phase: "closing", can_terminate: true }),
    terminateOwnedProcess() {
      signals++;
    },
  });
  await assert.rejects(h.call("recover_runtime", { expected_pid: 12345 }), /explicitly confirm/);
  h.busy("active-turn");
  await assert.rejects(h.call("recover_runtime", { expected_pid: 12345, confirm: true }), /Stop/);
  h.busy(null);
  h.handoffBusy(true);
  await assert.rejects(
    h.call("recover_runtime", { expected_pid: 12345, confirm: true }),
    /confirm/,
  );
  h.handoffBusy(false);
  assert.equal(signals, 0);
  exit.resolve();
});

test("actual recovery RPC and archive cannot treat a sent signal as confirmed process exit", async () => {
  const h = await harness();
  h.stop();
  const exit = deferred();
  let signals = 0;
  h.closes.track("pane", exit.promise, {
    getCloseStatus: () => ({ pid: 12345, phase: "closing", can_terminate: true }),
    terminateOwnedProcess({ expectedPid, signal, confirmForce }) {
      assert.equal(expectedPid, 12345);
      assert.equal(signal, "SIGTERM");
      assert.equal(confirmForce, false);
      signals++;
    },
  });
  await assert.rejects(h.call("recover_runtime", { expected_pid: 12345, confirm: true }), {
    code: "RUNTIME_EXIT_UNCONFIRMED",
  });
  assert.equal(signals, 1);
  assert.equal(h.closes.has("pane"), true);
  await assert.rejects(h.call("release_session"), { code: "RUNTIME_EXIT_UNCONFIRMED" });
  assert.ok(!h.order.some((entry) => entry.startsWith("archive:")));
  exit.resolve();
  await h.closes.wait("pane");
  await h.call("release_session");
  assert.ok(h.order.some((entry) => entry.startsWith("archive:")));
});

test("rejected observation remains a permanent fail-closed barrier even if a later status claims exit", async () => {
  const h = await harness();
  h.stop();
  let phase = "unobservable";
  h.closes.track("pane", Promise.reject(new Error("exit observation invalid")), {
    getCloseStatus: () => ({ pid: 12345, phase, can_terminate: false }),
    terminateOwnedProcess() {
      throw new Error("cannot recover unobservable process");
    },
  });
  phase = "exited";
  await assert.rejects(h.call("release_session"), /exit observation invalid/);
  assert.equal(h.closes.has("pane"), true);
  assert.ok(!h.order.some((entry) => entry.startsWith("archive:")));
});
