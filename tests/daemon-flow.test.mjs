import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile, mkdtemp, writeFile, appendFile, rename, rm, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { randomUUID } from "node:crypto";
import {
  captureTranscriptFingerprint,
  sameTranscriptFingerprint,
} from "../daemon/session-transcript-state.mjs";
import vm from "node:vm";
import { createLifecycle } from "../daemon/runtime/lifecycle.mjs";
import { createMessageRenderer } from "../daemon/runtime/render-messages.mjs";
import { createTurnSource } from "../daemon/runtime/turns.mjs";
import { classifyResult, classifyAgentError } from "../daemon/runtime/messages.mjs";
import { createRuntimeCloseRegistry } from "../daemon/runtime/close-registry.mjs";
import { createSessionPersistenceRegistry } from "../daemon/session-persistence.mjs";

// Exercise the actual daemon orchestration functions without booting its servers,
// reading user settings, starting a CLI, or contacting model providers.
async function harness(touchFolder = async () => {}) {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const start = source.slice(
    source.indexOf("async function startSessionForFolder("),
    source.indexOf("\nasync function switchFolder("),
  );
  const handler = source.slice(
    source.indexOf("function handleAgentMessage("),
    source.indexOf(
      "\n// ---------------------------------------------------------------------------\n// Kick off.",
    ),
  );
  const sessions = new Map(),
    events = [],
    models = new Map([["pane", "selected"]]);
  let active = "one",
    params,
    closes = 0,
    cleared = 0;
  const pending = [];
  const bridge = {
    activeTurn: () => active,
    nextUserMessage: async () => pending.shift(),
    clearUserMessages() {
      cleared++;
      pending.length = 0;
    },
    cancelTools() {},
    sendAssistantEvent(e) {
      events.push(e);
      if (e.event === "turn_complete") active = null;
    },
    callTaskpaneTool: async () => "tool-result",
  };
  const context = vm.createContext({
    shuttingDown: false,
    console,
    sessions,
    workspaceByKey: new Map(),
    modelByKey: models,
    createLifecycle,
    createMessageRenderer,
    createTurnSource,
    classifyResult,
    classifyAgentError,
    touchFolder,
    handoffInProgress: new Set(),
    runtimeClosures: createRuntimeCloseRegistry(),
    getHandoff: async () => null,
    bridge,
    sessionFor: (key) => sessions.get(key),
    buildSystemPromptAppend: async () => "fixture",
    modelArgFor: () => "selected",
    createOfficeBridgeMcp: (scoped) => scoped,
    userMcpServers: {},
    WORD_MCP_DISALLOWED: [],
    customPermissionHandler: async () => ({ behavior: "allow" }),
    runtime: {
      id: "codebuddy",
      mcp: {},
      runSession(p) {
        params = p;
        return {
          closed: Promise.resolve(),
          close() {
            closes++;
          },
        };
      },
    },
    diag() {},
    saveSessionId: async () => {},
    sessionPersistence: createSessionPersistenceRegistry({ save: async () => {} }),
  });
  vm.runInContext(start + "\n" + handler + "\nglobalThis.start = startSessionForFolder;", context);
  return {
    start: () => context.start("/fixture", null, { key: "pane", host: "word", replay: false }),
    sessions,
    events,
    params: () => params,
    closes: () => closes,
    cleared: () => cleared,
    async turn(id) {
      active = id;
      pending.push({ text: "fixture question", context: { host: "word" }, turn_id: id });
      return params.turnSource.next();
    },
  };
}

test("actual daemon startup can be stopped while persistence is still awaiting", async () => {
  let release;
  const h = await harness(() => new Promise((r) => (release = r)));
  const starting = h.start();
  const s = h.sessions.get("pane");
  assert.equal(typeof s.cancel, "function");
  s.cancel();
  release();
  await starting;
  assert.equal(h.params(), undefined);
  assert.equal(h.sessions.size, 0);
  assert.equal(h.events.filter((e) => e.event === "turn_complete").length, 1);
});

test("actual daemon startup failure reports one terminal and clears its queued turn", async () => {
  const h = await harness(async () => {
    throw new Error("disk unavailable");
  });
  await h.start();
  const terminal = h.events.filter((e) => e.event === "turn_complete");
  assert.equal(terminal.length, 1);
  assert.equal(terminal[0].subtype, "startup_error");
  assert.match(terminal[0].error, /disk unavailable/);
  assert.equal(h.sessions.size, 0);
  assert.equal(h.cleared(), 1);
});

test("actual daemon second-turn stream failure is terminal; stale MCP and text are fenced", async () => {
  const h = await harness();
  await h.start();
  const p = h.params();
  await h.turn("one");
  p.onMessage({
    type: "assistant",
    message: { id: "m1", content: [{ type: "text", text: "answer" }] },
  });
  p.onMessage({ type: "result", subtype: "success" });
  await h.turn("two");
  p.onEnd({ reason: "ended" });
  const count = h.events.length;
  p.onMessage({
    type: "assistant",
    message: { id: "late", content: [{ type: "text", text: "late" }] },
  });
  p.onTurnEnd({ sawResult: false });
  assert.equal(h.events.length, count);
  assert.deepEqual(
    h.events.filter((e) => e.event === "turn_complete").map((e) => e.subtype),
    ["success", "stream_ended"],
  );
  assert.throws(
    () => p.mcpServers.office.callTaskpaneTool("office_replace_text", {}, "pane"),
    /Expired session/,
  );
  assert.equal(h.closes(), 1);
  assert.equal(h.sessions.size, 0);
});

test("actual daemon fails closed if SDK init reports a different model", async () => {
  const h = await harness();
  await h.start();
  await h.turn("one");
  h.params().onMessage({ type: "system", subtype: "init", model: "unexpected", session_id: "id" });
  assert.equal(h.events.find((e) => e.event === "turn_complete").subtype, "model_mismatch");
  assert.equal(h.closes(), 1);
});

// Actual queued-start and Stop handlers, with controllable previous-child exit.
// This reproduces races at the async teardown boundary without a provider call.
async function exitWaitStartupHarness() {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const take = (start, end) =>
    source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start)));
  const sessions = new Map(),
    events = [],
    pending = [],
    waits = [],
    starts = [],
    logs = [];
  let active = null;
  const bridge = {
    activeTurn: () => active,
    clearUserMessages: () => {
      pending.length = 0;
    },
    cancelTools() {},
    nextUserMessage: async () => pending.shift(),
    sendAssistantEvent(event) {
      events.push(event);
      if (event.event === "turn_complete" && event.turn_id === active) active = null;
    },
  };
  const context = vm.createContext({
    console: { error: (...args) => logs.push(args) },
    shuttingDown: false,
    setImmediate,
    sessions,
    workspaceByKey: new Map(),
    bridge,
    sessionFor: (key) => sessions.get(key),
    runtimeClosures: {
      has: () => true,
      wait() {
        return new Promise((resolve, reject) => {
          waits.push({ resolve, reject });
        });
      },
      track() {},
    },
    createLifecycle,
    createMessageRenderer,
    createTurnSource,
    classifyAgentError,
    touchFolder: async () => {},
    getHandoff: async () => null,
    handoffInProgress: new Set(),
    buildSystemPromptAppend: async () => "synthetic",
    modelArgFor: () => "synthetic-model",
    createOfficeBridgeMcp: (scoped) => scoped,
    userMcpServers: {},
    WORD_MCP_DISALLOWED: [],
    customPermissionHandler: async () => ({}),
    runtime: {
      id: "codebuddy",
      mcp: {},
      runSession(options) {
        starts.push(options);
        return { closed: Promise.resolve(), close() {} };
      },
    },
  });
  vm.runInContext(
    [
      take("const startQueue = new Map();", "// Called on every taskpane hello"),
      take("async function startSessionForFolder(", "\nasync function switchFolder("),
      `globalThis.handlers = {${take("    stop_agent: async", "    get_session_info: async")}};`,
      "globalThis.schedule = scheduleSessionStart; globalThis.queue = startQueue;",
    ].join("\n"),
    context,
  );
  const until = async (check) => {
    for (let n = 0; n < 50 && !check(); n++) await new Promise((resolve) => setImmediate(resolve));
    assert.ok(check(), "Synthetic start queue did not reach the expected checkpoint");
  };
  return {
    events,
    pending,
    waits,
    starts,
    logs,
    sessions,
    active: () => active,
    submit(id, cwd, resumeId) {
      assert.equal(active, null, "The synthetic bridge requires serial turns");
      active = id;
      pending.push({ text: "Synthetic " + id, context: { host: "word" }, turn_id: id });
      context.schedule(cwd, resumeId, "pane", "word", "synthetic request", { replay: false });
    },
    stop: () => context.handlers.stop_agent({ request_id: "stop" }, () => {}, "pane"),
    shuttingDown: () => {
      context.shuttingDown = true;
    },
    replaceQueue: () => context.queue.delete("pane"),
    reschedule: (cwd, resumeId) =>
      context.schedule(cwd, resumeId, "pane", "word", "synthetic replacement", { replay: false }),
    waitForCount: (count) => until(() => waits.length === count),
    idle: () => until(() => !context.queue.get("pane")?.running),
  };
}

test("actual start queue turns an exit-wait timeout into one terminal failure for its accepted turn", async () => {
  const h = await exitWaitStartupHarness();
  try {
    h.submit("resend-after-stop", "/synthetic-A", "session-A");
    await h.waitForCount(1);
    assert.equal(h.sessions.get("pane").lifecycle.turnId, "resend-after-stop");
    h.waits[0].reject(new Error("Runtime process exit not confirmed"));
    await h.idle();
    const terminal = h.events.filter((event) => event.event === "turn_complete");
    assert.equal(terminal.length, 1);
    assert.equal(terminal[0].turn_id, "resend-after-stop");
    assert.equal(terminal[0].subtype, "startup_error");
    assert.match(terminal[0].error, /exit not confirmed/);
    assert.equal(h.active(), null);
    assert.equal(h.pending.length, 0);
    assert.equal(h.sessions.size, 0);
    assert.equal(h.starts.length, 0);
    assert.equal(h.logs.length, 0, "Failure must be surfaced to the owning turn, not only logged");
  } finally {
    await h.stop();
  }
});

test("shutdown fences an asynchronous queued session before SDK spawn", async () => {
  const h = await exitWaitStartupHarness();
  h.submit("shutdown-race", "/synthetic", null);
  await h.waitForCount(1);
  h.shuttingDown();
  h.waits[0].resolve();
  await h.idle();
  assert.equal(h.starts.length, 0);
});

for (const outcome of ["exit", "timeout", "reconnected-queue"]) {
  test(`actual Stop then resend fences the old ${outcome} continuation from its successor`, async () => {
    const h = await exitWaitStartupHarness();
    try {
      h.submit("turn-A", "/synthetic-A", "session-A");
      await h.waitForCount(1);
      await h.stop();
      if (outcome === "reconnected-queue") h.replaceQueue();
      h.submit("turn-B", "/synthetic-B", "session-B");
      if (outcome === "exit") h.waits[0].resolve();
      else h.waits[0].reject(new Error("Old runtime exit not confirmed"));
      await h.waitForCount(2);
      assert.equal(h.starts.length, 0, "Old cwd/session must not start against the new turn");
      assert.equal(h.active(), "turn-B");
      assert.deepEqual(
        h.pending.map((message) => message.turn_id),
        ["turn-B"],
      );
      h.waits[1].resolve();
      await h.idle();
      assert.equal(h.starts.length, 1);
      assert.equal(h.starts[0].cwd, "/synthetic-B");
      assert.equal(h.starts[0].resumeId, "session-B");
      assert.equal(h.sessions.get("pane").lifecycle.turnId, "turn-B");
      assert.equal(h.active(), "turn-B");
      assert.deepEqual(
        h.pending.map((message) => message.turn_id),
        ["turn-B"],
      );
      const terminal = h.events.filter((event) => event.event === "turn_complete");
      assert.equal(terminal.length, 1);
      assert.equal(terminal[0].turn_id, "turn-A");
      assert.equal(terminal[0].subtype, "interrupted");
      assert.deepEqual(
        h.events.filter((event) => event.event === "cwd_changed").map((event) => event.cwd),
        ["/synthetic-B"],
      );
      assert.equal(h.logs.length, 0);
    } finally {
      await h.stop();
    }
  });
}

test("same-turn queued generation replacement preserves the prompt and starts only the latest cwd/resume", async () => {
  const h = await exitWaitStartupHarness();
  try {
    h.submit("one-turn", "/old-workspace", "old-session");
    await h.waitForCount(1);
    h.reschedule("/new-workspace", "new-session");
    h.waits[0].resolve();
    await h.waitForCount(2);
    assert.equal(h.starts.length, 0);
    assert.equal(h.active(), "one-turn");
    assert.deepEqual(
      h.pending.map((message) => message.turn_id),
      ["one-turn"],
    );
    h.waits[1].resolve();
    await h.idle();
    assert.equal(h.starts.length, 1);
    assert.equal(h.starts[0].cwd, "/new-workspace");
    assert.equal(h.starts[0].resumeId, "new-session");
    assert.equal(h.events.filter((event) => event.event === "turn_complete").length, 0);
    assert.deepEqual(
      h.pending.map((message) => message.turn_id),
      ["one-turn"],
    );
  } finally {
    await h.stop();
  }
});

async function handoffHarness({
  archiveFails = false,
  active = null,
  pendingClose = null,
  stopped = false,
  fingerprint = async () => ({
    path: "/fixture/id.jsonl",
    dev: 1,
    ino: 1,
    size: 1,
    mtime_ms: 1,
    ctime_ms: 1,
    sha256: "fixture",
  }),
} = {}) {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const code = source.slice(
    source.indexOf("    get_session_info: async"),
    source.indexOf("    get_context: async"),
  );
  let saved = null;
  const order = [],
    replies = [];
  const runtimeClosures = createRuntimeCloseRegistry();
  if (pendingClose) runtimeClosures.track("pane", pendingClose);
  const context = vm.createContext({
    Date,
    setTimeout,
    clearTimeout,
    Promise,
    randomUUID,
    captureTranscriptFingerprint: fingerprint,
    sameTranscriptFingerprint,
    handoffInProgress: new Set(),
    runtimeClosures: { wait: (key) => runtimeClosures.wait(key, 10) },
    sessionPersistence: createSessionPersistenceRegistry({ save: async () => {} }),
    cwdForKey: () => "/fixture",
    runtime: { id: "codebuddy" },
    startQueue: new Map(),
    bridge: { activeTurn: () => active },
    sessionFor: () =>
      stopped
        ? null
        : {
            cancel() {
              order.push("cancel");
            },
            done: Promise.resolve(),
          },
    sessionInfo: async () => ({
      provider: "codebuddy",
      host: "word",
      cwd: "/fixture",
      session_id: "id",
      state: saved?.state || "word",
      return_confirmation: saved?.return_confirmation?.id ?? null,
    }),
    getHandoff: async () => saved,
    setHandoff: async (host, cwd, provider, key, value) => {
      saved = value;
      order.push(value ? "release" : "return");
    },
    archiveSession: async () => {
      order.push("archive");
      if (archiveFails) throw new Error("disk full");
      return { created_at: "time", archive_path: "/fixture/archive" };
    },
    sendTranscriptReplayTo: async () => {
      order.push("replay");
    },
  });
  vm.runInContext(`globalThis.handlers = {${code}}`, context);
  return {
    call: (name, msg = {}) => context.handlers[name](msg, (r) => replies.push(r), "pane", "word"),
    order,
    replies,
    saved: () => saved,
  };
}
test("actual handoff handlers close/archive before resume and replay after explicit return", async () => {
  const h = await handoffHarness();
  await h.call("release_session");
  assert.deepEqual(h.order, ["release", "cancel", "release", "archive", "release"]);
  assert.equal(h.replies[0].info.state, "released");
  await assert.rejects(h.call("return_session", { session_id: "id" }), /exited/);
  await assert.rejects(
    h.call("return_session", { session_id: "wrong", cli_closed: true }),
    /identity/,
  );
  await h.call("return_session", { session_id: "id", cli_closed: true });
  assert.deepEqual(h.order.slice(-2), ["return", "replay"]);
  assert.equal(h.saved(), null);
});
test("external transcript writes require a new explicit confirmation and a still-matching fingerprint", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-handoff-fingerprint-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const path = join(root, "id.jsonl");
  await writeFile(path, '{"type":"user","message":"synthetic history"}\n');
  const fingerprint = (info) => captureTranscriptFingerprint(info, { locate: async () => path });
  const h = await handoffHarness({ fingerprint });
  await h.call("release_session");
  const original = h.saved().transcript_fingerprint;
  await appendFile(path, '{"type":"assistant","message":"CLI continuation"}\n');
  await h.call("return_session", { session_id: "id", cli_closed: true });
  let refused = h.replies.at(-1);
  assert.equal(refused.ok, false);
  assert.equal(refused.code, "TRANSCRIPT_CHANGED");
  assert.equal(h.saved().state, "released");
  assert.ok(!h.order.includes("replay"));
  const first = refused.info.return_confirmation;
  assert.ok(first);
  await h.call("return_session", { session_id: "id", cli_closed: true });
  assert.equal(h.replies.at(-1).ok, false, "no implicit second confirmation");

  // Continued writing invalidates even an explicitly supplied old challenge.
  await appendFile(path, '{"type":"assistant","message":"still writing"}\n');
  await h.call("return_session", {
    session_id: "id",
    cli_closed: true,
    confirm_transcript_change: first,
  });
  refused = h.replies.at(-1);
  assert.equal(refused.ok, false);
  assert.notEqual(refused.info.return_confirmation, first);
  assert.deepEqual(
    h.saved().transcript_fingerprint,
    original,
    "release baseline is not silently accepted",
  );
  assert.ok(!h.order.includes("replay"));
  await assert.rejects(
    h.call("return_session", {
      session_id: "id",
      confirm_transcript_change: refused.info.return_confirmation,
    }),
    /exited/,
  );
  await h.call("return_session", {
    session_id: "id",
    cli_closed: true,
    confirm_transcript_change: refused.info.return_confirmation,
  });
  assert.equal(h.replies.at(-1).ok, true);
  assert.equal(h.saved(), null);
  assert.equal(h.order.filter((item) => item === "replay").length, 1, "history-only replay once");
});

test("transcript fingerprint detects same-content inode replacement and refuses symlinks", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-transcript-state-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const path = join(root, "id.jsonl"),
    other = join(root, "replacement.jsonl");
  const info = { session_id: "id", provider: "codebuddy" };
  const capture = () => captureTranscriptFingerprint(info, { locate: async () => path });
  await writeFile(path, "synthetic history\n");
  const before = await capture();
  await writeFile(other, "synthetic history\n");
  await rename(other, path);
  const after = await capture();
  assert.equal(before.sha256, after.sha256);
  assert.notEqual(before.ino, after.ino);
  assert.equal(sameTranscriptFingerprint(before, after), false);
  assert.equal(sameTranscriptFingerprint(after, { ...after }), true);
  assert.equal(sameTranscriptFingerprint({}, {}), false);
  const link = join(root, "link.jsonl");
  await symlink(path, link);
  await assert.rejects(captureTranscriptFingerprint(info, { locate: async () => link }));
  await assert.rejects(
    captureTranscriptFingerprint(info, { locate: async () => null }),
    /TRANSCRIPT_UNAVAILABLE/,
  );
});

test("archive failure leaves a recoverable released session; active turns cannot be handed off", async () => {
  const h = await handoffHarness({ archiveFails: true });
  await assert.rejects(h.call("release_session"), /disk full/);
  assert.equal(h.saved().state, "released");
  await h.call("return_session", { session_id: "id", cli_closed: true });
  assert.equal(h.saved(), null);
  const busy = await handoffHarness({ active: "turn" });
  await assert.rejects(busy.call("release_session"), /current turn/);
  assert.deepEqual(busy.order, []);
});

test("actual release/return handlers retain the exit barrier after Stop removed the live session", async () => {
  let exit;
  const pendingClose = new Promise((resolve) => {
    exit = resolve;
  });
  const h = await handoffHarness({ pendingClose, stopped: true });
  await assert.rejects(h.call("release_session"), /exit not confirmed/);
  assert.equal(h.saved().state, "released");
  assert.ok(!h.order.includes("archive"));
  await assert.rejects(
    h.call("return_session", { session_id: "id", cli_closed: true }),
    /exit not confirmed/,
  );
  assert.equal(h.saved().state, "released");
  exit();
  await h.call("release_session");
  assert.ok(h.order.includes("archive"));
  await h.call("return_session", { session_id: "id", cli_closed: true });
  assert.equal(h.saved(), null);
});

test("actual session info withholds CLI commands until a complete archive is recorded", async () => {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const code = source.slice(
    source.indexOf("async function sessionInfo("),
    source.indexOf("// Static fallback"),
  );
  let handoff = { state: "released" };
  const context = vm.createContext({
    runtime: { id: "codebuddy" },
    cwdForKey: () => "/fixture",
    resolveReplaySessionId: async () => "id",
    getHandoff: async () => handoff,
    sessionFor: () => null,
    resumeCommands: () => ({ resume: "synthetic-command" }),
    runtimeClosures: createRuntimeCloseRegistry(),
    sessionPersistence: createSessionPersistenceRegistry({ save: async () => {} }),
  });
  vm.runInContext(code + "\nglobalThis.info = sessionInfo;", context);
  assert.equal((await context.info("pane", "word")).commands, null);
  handoff = { ...handoff, archive_path: "/fixture/complete-archive" };
  assert.equal((await context.info("pane", "word")).commands.resume, "synthetic-command");
});
