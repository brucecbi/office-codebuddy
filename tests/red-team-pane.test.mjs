// Independent adversarial tests. Actual pane script; synthetic DOM/Office only.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { randomUUID } from "node:crypto";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";
import { planReplacement } from "../taskpane/shared/word-edit-plan.js";
import { createBridge } from "../daemon/bridge.mjs";
import WebSocket from "ws";
import { once } from "node:events";

class Element {
  constructor(tag = "div") {
    this.tagName = tag;
    this.children = [];
    this.dataset = {};
    this.style = {};
    this.listeners = {};
    this.queries = new Map();
    this.classList = { add() {}, remove() {}, toggle() {} };
    this.value = "";
    this.isConnected = true;
    this.scrollHeight = this.scrollTop = this.clientHeight = 0;
  }
  addEventListener(type, fn) {
    this.listeners[type] = fn;
  }
  appendChild(el) {
    this.children.push(el);
    return el;
  }
  append(...els) {
    this.children.push(...els);
  }
  prepend(el) {
    this.children.unshift(el);
  }
  remove() {
    this.isConnected = false;
  }
  setAttribute() {}
  querySelectorAll() {
    return [];
  }
  querySelector(selector) {
    if (!this.queries.has(selector)) this.queries.set(selector, new Element());
    return this.queries.get(selector);
  }
  closest() {
    return new Element();
  }
  focus() {}
  setSelectionRange() {}
  set innerHTML(value) {
    this.html = value;
    this.children = [];
  }
  get innerHTML() {
    return this.html || "";
  }
}

async function makePane({ storage = new Map(), apply = async () => ({ verified: true }) } = {}) {
  const elements = new Map();
  const adapter = { getItem: (k) => storage.get(k) ?? null, setItem: (k, v) => storage.set(k, v) };
  if (!storage.has("draftspect-preview-settings-v1"))
    storage.set(
      "draftspect-preview-settings-v1",
      JSON.stringify({
        model: "synthetic-model-a",
        modelsByRuntime: { codebuddy: "synthetic-model-a" },
      }),
    );
  const body = new Element();
  body.dataset.host = "word";
  const document = {
    body,
    documentElement: new Element(),
    visibilityState: "visible",
    hidden: false,
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, new Element());
      return elements.get(id);
    },
    querySelectorAll: () => [],
    createElement: (tag) => new Element(tag),
    addEventListener() {},
  };
  const messages = [];
  let reads = 0,
    closes = 0;
  const socket = {
    send: (raw) => messages.push(JSON.parse(raw)),
    close() {
      closes++;
    },
  };
  const office = { context: { document: { url: "synthetic.docx" } }, onReady() {} };
  const unrefTimer = (...args) => {
    const h = setTimeout(...args);
    h.unref();
    return h;
  };
  const unrefInterval = (...args) => {
    const h = setInterval(...args);
    h.unref();
    return h;
  };
  const context = vm.createContext({
    location: { protocol: "http:" },
    document,
    console,
    crypto: { randomUUID },
    localStorage: adapter,
    sessionStorage: adapter,
    Office: office,
    createOperationQueue,
    planReplacement,
    setTimeout: unrefTimer,
    clearTimeout,
    setInterval: unrefInterval,
    clearInterval,
    prepareReviewWrite: async () => ({ preview: { action: "synthetic" }, apply }),
    prepareSafeWordWrite: async () => ({ preview: { action: "synthetic" }, apply }),
    prepareTextReplacement: async () => ({ preview: { action: "synthetic" }, apply }),
    resolveWordHeadingArgs: async (_name, args) => args,
    captureWordFormattingPrecondition: async () => null,
    toolReadParagraphs: async () => {
      reads++;
      return { addressing: "uniqueLocalId", paragraphs: [{ id: "P1", text: "Fixture" }] };
    },
    socket,
    WebSocket: function () {
      return socket;
    },
  });
  context.window = context;
  context.addEventListener = () => {};
  const source = (
    await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8")
  ).replace(/^import\s+[\s\S]*?;\s*$/gm, "");
  vm.runInContext(
    source +
      `\nglobalThis.api = {
    handleServerMessage, runOfficeTool, usePreset, cancelOfficeOperations,
    connect() { ws=socket; wsReady=true; activeDocUrl='synthetic.docx'; },
    turn(id) { currentTurnId=id; beginTurn(); },
    ready() { modelConfirmed=true; replayReady=true; },
    gate(name,value) {
      if(name==='model') modelConfirmed=value;
      if(name==='replay') replayReady=value;
      if(name==='released') sessionReleased=value;
      if(name==='handoff') handoffBusy=value;
      if(name==='connected') wsReady=value;
    },
    state() { return { turnId:currentTurnId,inFlight:turnInFlight,confirmed:modelConfirmed,replay:replayReady }; },
    card(id) { return toolCards.get(id); },
    health() { return checkTurnHealth(); },
  };`,
    context,
    { filename: "actual-taskpane-red-team.js" },
  );
  context.api.connect();
  return {
    api: context.api,
    elements,
    storage,
    messages,
    socket,
    office,
    context,
    reads: () => reads,
    closes: () => closes,
  };
}
const flush = async () => {
  for (let i = 0; i < 12; i++) await Promise.resolve();
};
const sendPreset = (p) =>
  p.api.usePreset({ prompt: "Summarize synthetic document; do not edit", auto_send: true });

test("RED all auto-send admission gates are enforced even if preset UI remains clickable", async () => {
  for (const [gate, value] of [
    ["model", false],
    ["replay", false],
    ["released", true],
    ["handoff", true],
    ["connected", false],
  ]) {
    const p = await makePane();
    p.api.ready();
    p.api.gate(gate, value);
    sendPreset(p);
    assert.equal(
      p.messages.filter((m) => m.type === "user_message").length,
      0,
      `${gate} gate bypassed`,
    );
    assert.equal(p.api.state().inFlight, false, `${gate} gate left an orphan Working turn`);
  }
});

test("RED auto-send detects Save As identity change before sending", async () => {
  const p = await makePane();
  p.api.ready();
  p.office.context.document.url = "different-synthetic.docx";
  sendPreset(p);
  assert.equal(p.messages.filter((m) => m.type === "user_message").length, 0);
  assert.equal(p.api.state().inFlight, false);
  assert.equal(p.closes(), 1);
});

test("RED preset owns its turn and rejects an unrelated terminal/read tool", async () => {
  const p = await makePane();
  p.api.ready();
  sendPreset(p);
  const sent = p.messages.find((m) => m.type === "user_message");
  assert.ok(sent.turn_id, "preset must supply an owned turn ID");
  assert.ok(sent.request_id, "preset must supply a receipt correlation ID");
  assert.equal(sent.turn_id, p.api.state().turnId);
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: "foreign-old",
    subtype: "success",
  });
  assert.equal(p.api.state().inFlight, true);
  await p.api.runOfficeTool({
    id: "foreign-read",
    name: "office_read_paragraphs",
    args: {},
    turn_id: "foreign-old",
    deadline: Date.now() + 120000,
  });
  assert.equal(p.reads(), 0);
  assert.equal(p.messages.find((m) => m.id === "foreign-read").ok, false);
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: sent.turn_id,
    subtype: "success",
  });
  assert.equal(p.api.state().inFlight, false);
  assert.equal(p.messages.filter((m) => m.type === "user_message").length, 1);
});

test("RED a possibly committed edit remains blocked after terminal recovery and reload", async () => {
  const storage = new Map();
  let writes = 0;
  const apply = async ({ beforeWrite }) => {
    beforeWrite();
    writes++;
    throw new Error("Lost completion after commit");
  };
  const p = await makePane({ storage, apply });
  p.api.turn("uncertain");
  const write = {
    id: "write-attempt",
    name: "office_edit_range",
    args: { target: { query: "Fixture" }, action: "replace", text: "Changed" },
    turn_id: "uncertain",
    deadline: Date.now() + 120000,
  };
  const running = p.api.runOfficeTool(write);
  await flush();
  const approve = p.api
    .card(write.id)
    .el.children.find((c) => c.textContent === "Apply this change");
  assert.ok(approve);
  approve.listeners.click();
  await running;
  assert.equal(writes, 1);
  assert.match(
    p.messages.find((m) => m.type === "tool_result" && m.id === write.id).error,
    /RESULT_UNKNOWN/,
  );
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: "uncertain",
    subtype: "success",
  });
  const reloaded = await makePane({ storage, apply });
  reloaded.api.turn("later");
  await reloaded.api.runOfficeTool({ ...write, turn_id: "later" });
  await reloaded.api.runOfficeTool({ ...write, id: "new-write-id", turn_id: "later" });
  assert.equal(writes, 1, "neither receipt ID replay nor fresh ID may replay an uninspected edit");
  assert.ok(
    reloaded.messages
      .filter((m) => m.type === "tool_result")
      .every((m) => m.ok === false && /RESULT_UNKNOWN/.test(m.error)),
  );
});

test("RED an older health response cannot cancel a newer turn on the same connection", async () => {
  const p = await makePane();
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: "epoch-A",
    capabilities: ["turn_health_v1"],
  });
  p.api.ready();
  sendPreset(p);
  const first = p.messages.find((m) => m.type === "user_message");
  const checking = p.api.health();
  const query = p.messages.find((m) => m.type === "get_turn_state");
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: first.turn_id,
    subtype: "success",
  });
  sendPreset(p);
  const second = p.messages.filter((m) => m.type === "user_message").at(-1);
  assert.notEqual(second.turn_id, first.turn_id);
  await p.api.handleServerMessage({
    type: "get_turn_state_result",
    request_id: query.request_id,
    ok: true,
    connection_epoch: "epoch-A",
    active: null,
    terminal: { event: "turn_complete", turn_id: first.turn_id, subtype: "success" },
  });
  await checking;
  assert.equal(
    p.api.state().turnId,
    second.turn_id,
    "stale health snapshot terminated the new turn",
  );
  assert.equal(p.api.state().inFlight, true);
});

test("RED a different RPC result type cannot impersonate a user ACK", async () => {
  const p = await makePane();
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: "epoch-A",
    capabilities: ["turn_health_v1"],
  });
  p.api.ready();
  sendPreset(p);
  const request = p.messages.find((m) => m.type === "user_message");
  await p.api.handleServerMessage({
    type: "get_turn_state_result",
    request_id: request.request_id,
    turn_id: request.turn_id,
    connection_epoch: "epoch-A",
    ok: true,
  });
  await flush();
  assert.match(p.elements.get("agent-status").textContent, /提交/);
  await p.api.handleServerMessage({
    type: "user_message_result",
    request_id: request.request_id,
    turn_id: request.turn_id,
    connection_epoch: "epoch-A",
    ok: true,
  });
  await flush();
  assert.match(p.elements.get("agent-status").textContent, /已接收/);
  assert.equal(p.messages.filter((m) => m.type === "user_message").length, 1);
});

test("RED healthy RPCs do not disguise stalled model progress or trigger request replay", async () => {
  const p = await makePane();
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: "epoch-A",
    capabilities: ["turn_health_v1"],
  });
  p.api.ready();
  sendPreset(p);
  const turn = p.api.state().turnId;
  for (let i = 0; i < 2; i++) {
    const checking = p.api.health();
    const query = p.messages.filter((m) => m.type === "get_turn_state").at(-1);
    await p.api.handleServerMessage({
      type: "get_turn_state_result",
      request_id: query.request_id,
      ok: true,
      connection_epoch: "epoch-A",
      active: {
        turn_id: turn,
        phase: "waiting_model",
        started_at: Date.now() - 60000,
        last_progress_at: Date.now() - 60000,
        pending_tools: [],
      },
      terminal: null,
    });
    await checking;
    assert.match(p.elements.get("health-notice-detail").textContent, /SLOW_PROGRESS/);
    assert.match(p.elements.get("agent-status").textContent, /等待模型/);
  }
  assert.equal(p.api.state().inFlight, true);
  assert.equal(p.messages.filter((m) => m.type === "user_message").length, 1);
  assert.equal(p.reads(), 0);
});

test("RED repeated wrong-epoch health snapshots block tools without accepting foreign completion", async () => {
  const p = await makePane();
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: "epoch-A",
    capabilities: ["turn_health_v1"],
  });
  p.api.ready();
  sendPreset(p);
  const turn = p.api.state().turnId;
  for (let i = 0; i < 2; i++) {
    const checking = p.api.health();
    const query = p.messages.filter((m) => m.type === "get_turn_state").at(-1);
    await p.api.handleServerMessage({
      type: "get_turn_state_result",
      request_id: query.request_id,
      ok: true,
      connection_epoch: "OLD",
      active: null,
      terminal: { event: "turn_complete", turn_id: turn, subtype: "success" },
    });
    await checking;
  }
  assert.equal(p.api.state().inFlight, true);
  assert.equal(p.api.state().turnId, turn);
  await p.api.runOfficeTool({
    id: "epoch-blocked",
    name: "office_read_paragraphs",
    args: {},
    turn_id: turn,
    deadline: Date.now() + 120000,
  });
  assert.equal(p.reads(), 0);
  assert.match(
    p.messages.find((m) => m.type === "tool_result" && m.id === "epoch-blocked").error,
    /CONNECTION_UNAVAILABLE/,
  );
  assert.equal(p.messages.filter((m) => m.type === "user_message").length, 1);
});

const waitFor = async (predicate, label = "condition") => {
  const deadline = Date.now() + 2500;
  while (Date.now() < deadline) {
    if (predicate()) return;
    await new Promise((r) => setTimeout(r, 5));
  }
  throw new Error("Timed out waiting for " + label);
};

test("RED E2E actual bridge→actual preset pane→Office proxy recovers lost ACK and terminal without replay", async () => {
  const p = await makePane();
  let key,
    driverDone = false,
    driverError = null,
    admitted = 0,
    droppedAck = 0,
    droppedTerminal = 0;
  const rpc = (fields) => (msg, reply) =>
    reply({ type: msg.type + "_result", request_id: msg.request_id, ok: true, ...fields });
  let bridge;
  bridge = createBridge({
    port: 0,
    token: "synthetic-red-e2e",
    runtimeId: "codebuddy",
    onHello: (k) => {
      key = k;
      bridge.sendToTaskpane(
        {
          type: "transcript_replay",
          events: [],
          session_info: { provider: "codebuddy", state: "word" },
        },
        k,
      );
    },
    onUserMessage: (k) => {
      admitted++;
      setImmediate(async () => {
        try {
          const turn = await bridge.nextUserMessage(k);
          const output = await bridge.callTaskpaneTool("office_read_paragraphs", {}, k);
          assert.equal(output.paragraphs[0].text, "Fixture");
          bridge.sendAssistantEvent(
            { event: "turn_complete", turn_id: turn.turn_id, subtype: "success" },
            k,
          );
        } catch (e) {
          driverError = e;
        } finally {
          driverDone = true;
        }
      });
    },
    extraHandlers: {
      list_models: rpc({ runtime: "codebuddy", models: [{ value: "synthetic-model-a" }] }),
      set_model: rpc({ model: "synthetic-model-a" }),
      get_cwd_state: rpc({ current_cwd: "/synthetic" }),
      get_context: rpc({ entries: [] }),
    },
  });
  let client;
  try {
    await waitFor(() => bridge.address(), "bridge bind");
    client = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(client, "open");
    p.socket.send = (raw) => {
      p.messages.push(JSON.parse(raw));
      client.send(raw);
    };
    p.socket.close = () => client.close();
    const inbound = [];
    client.on("message", (raw) => {
      const m = JSON.parse(String(raw));
      inbound.push(m);
      if (m.type === "user_message_result") {
        droppedAck++;
        return;
      }
      if (m.type === "assistant_event" && m.event === "turn_complete") {
        droppedTerminal++;
        return;
      }
      void p.api.handleServerMessage(m).catch((e) => {
        driverError = e;
      });
    });
    client.send(
      JSON.stringify({
        type: "hello",
        token: "synthetic-red-e2e",
        host: "word",
        active_doc: "synthetic.docx",
        protocol_version: 2,
      }),
    );
    await waitFor(() => p.api.state().confirmed && p.api.state().replay, "pane readiness");
    sendPreset(p);
    await waitFor(() => driverDone && droppedTerminal === 1, "synthetic tool and lost terminal");
    assert.equal(driverError, null);
    assert.equal(droppedAck, 1);
    assert.equal(p.reads(), 1);
    assert.equal(p.api.state().inFlight, true, "terminal really was dropped before reconciliation");
    const serverTerminal = bridge.health(key).terminal;
    assert.ok(serverTerminal);
    assert.equal(bridge.health(key).active, null);
    await p.api.health();
    assert.equal(p.api.state().inFlight, false);
    assert.match(p.elements.get("health-notice-detail").textContent, /TERMINAL_RECONCILED/);
    assert.equal(admitted, 1);
    assert.equal(p.messages.filter((m) => m.type === "user_message").length, 1);
    assert.equal(p.messages.filter((m) => m.type === "tool_result").length, 1);
    assert.equal(p.reads(), 1, "health did not rerun the Office operation");
  } finally {
    p.api.gate("connected", false);
    client?.terminate();
    await bridge.close();
  }
});

export { makePane, flush };
