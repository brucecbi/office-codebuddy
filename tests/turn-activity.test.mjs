// Integration tests (vm-sliced actual daemon orchestration, same technique as
// daemon-flow.test.mjs) for the activity-based turn timeout wiring:
//  - a busy turn (steady model/tool messages) is NOT killed by the idle timer
//  - a stalled turn IS stopped without inferring unchanged document content
//    from an empty pending-tools set (bridge.hasPendingWrites)
//  - turn completion is logged with subtype and elapsed time
//  - repeated SDK init messages are counted, for reconnect-churn attribution
//
// Red/blue: on the pre-fix code these fail (fixed 180s wall, no activity(),
// no hasPendingWrites plumbing, no turn/init logging).

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";

import { createLifecycle } from "../daemon/runtime/lifecycle.mjs";
import { createMessageRenderer } from "../daemon/runtime/render-messages.mjs";
import { createTurnSource } from "../daemon/runtime/turns.mjs";
import { classifyResult, classifyAgentError } from "../daemon/runtime/messages.mjs";
import { createRuntimeCloseRegistry } from "../daemon/runtime/close-registry.mjs";
import { createSessionPersistenceRegistry } from "../daemon/session-persistence.mjs";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function harness({ now = () => Date.now() } = {}) {
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
    logs = [],
    models = new Map([["pane", "selected"]]);
  let active = "one",
    params;
  const pending = [];
  const bridge = {
    activeTurn: () => active,
    nextUserMessage: async () => pending.shift(),
    clearUserMessages() {
      pending.length = 0;
    },
    cancelTools() {},
    sendAssistantEvent(e) {
      events.push(e);
      if (e.event === "turn_complete") active = null;
    },
    callTaskpaneTool: async () => "tool-result",
    hasPendingWrites: () => false,
  };
  const context = vm.createContext({
    shuttingDown: false,
    Date: { now },
    console: {
      log: (...args) => logs.push(args.join(" ")),
      warn: (...args) => logs.push(args.join(" ")),
      error: (...args) => logs.push(args.join(" ")),
    },
    sessions,
    workspaceByKey: new Map(),
    modelByKey: models,
    createLifecycle,
    createMessageRenderer,
    createTurnSource,
    classifyResult,
    classifyAgentError,
    touchFolder: async () => {},
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
        return { closed: Promise.resolve(), close() {} };
      },
    },
    diag() {},
    saveSessionId: async () => {},
    sessionPersistence: createSessionPersistenceRegistry({ save: async () => {} }),
  });
  vm.runInContext(start + "\n" + handler + "\nglobalThis.start = startSessionForFolder;", context);
  return {
    async start() {
      await context.start("/fixture", null, { key: "pane", host: "word", replay: false });
      return params;
    },
    events,
    logs,
    params: () => params,
    async pullTurn(id) {
      active = id;
      pending.push({ turn_id: id, text: "synthetic", context: { host: "word" } });
      return params.turnSource.next();
    },
  };
}

test("second-turn duration excludes earlier turns and user idle time", async () => {
  let clock = 1000;
  const h = await harness({ now: () => clock });
  const p = await h.start();
  await h.pullTurn("one");
  clock = 2000;
  p.onMessage({ type: "result", subtype: "success" });
  clock = 100000;
  await h.pullTurn("two");
  clock = 101000;
  p.onMessage({ type: "result", subtype: "success" });
  assert.match(
    h.logs.find((line) => line.includes("[turn] two end")),
    /elapsed=1000ms$/,
  );
});

test("a busy turn is not killed by the idle timer; a clean result ends it", async () => {
  process.env.DRAFTSPECT_TURN_IDLE_MS = "130";
  process.env.DRAFTSPECT_TURN_MAX_MS = "5000";
  try {
    const h = await harness();
    const p = await h.start();
    for (let i = 0; i < 6; i++) {
      await delay(60);
      p.onMessage({
        type: "assistant",
        message: { id: "m" + i, content: [{ type: "text", text: "chunk " + i }] },
      });
    }
    assert.equal(
      h.events.filter((e) => e.event === "turn_complete").length,
      0,
      "steady model activity must keep the turn alive past the idle limit",
    );
    p.onMessage({ type: "result", subtype: "success" });
    const terminal = h.events.filter((e) => e.event === "turn_complete");
    assert.equal(terminal.length, 1);
    assert.equal(terminal[0].subtype, "success");
  } finally {
    delete process.env.DRAFTSPECT_TURN_IDLE_MS;
    delete process.env.DRAFTSPECT_TURN_MAX_MS;
  }
});

test("a stalled turn with no pending write avoids any unchanged-document guarantee", async () => {
  process.env.DRAFTSPECT_TURN_IDLE_MS = "150";
  process.env.DRAFTSPECT_TURN_MAX_MS = "5000";
  try {
    const h = await harness();
    await h.start();
    await delay(450);
    const terminal = h.events.filter((e) => e.event === "turn_complete");
    assert.equal(terminal.length, 1, "the idle timer must end a stalled turn");
    assert.equal(terminal[0].subtype, "timeout");
    assert.match(terminal[0].error, /No Office edit call is pending/);
    assert.doesNotMatch(terminal[0].error, /document is unchanged/);
    assert.doesNotMatch(terminal[0].error, /unknown outcome/);
  } finally {
    delete process.env.DRAFTSPECT_TURN_IDLE_MS;
    delete process.env.DRAFTSPECT_TURN_MAX_MS;
  }
});

test("turn completion is logged with subtype and elapsed time", async () => {
  process.env.DRAFTSPECT_TURN_MAX_MS = "5000";
  try {
    const h = await harness();
    const p = await h.start();
    p.onMessage({ type: "result", subtype: "success" });
    const line = h.logs.find((l) => l.includes("[turn]") && l.includes("success"));
    assert.ok(line, "turn end must be logged");
    assert.match(line, /elapsed=\d+ms/);
  } finally {
    delete process.env.DRAFTSPECT_TURN_MAX_MS;
  }
});

test("repeated SDK init messages are counted for churn attribution", async () => {
  process.env.DRAFTSPECT_TURN_MAX_MS = "5000";
  try {
    const h = await harness();
    const p = await h.start();
    p.onMessage({ type: "system", subtype: "init", model: "selected", session_id: "s-1" });
    p.onMessage({ type: "system", subtype: "init", model: "selected", session_id: "s-1" });
    assert.ok(h.logs.some((l) => l.includes("init session s-1") && l.includes("#1")));
    assert.ok(h.logs.some((l) => l.includes("init session s-1") && l.includes("#2")));
  } finally {
    delete process.env.DRAFTSPECT_TURN_MAX_MS;
  }
});
