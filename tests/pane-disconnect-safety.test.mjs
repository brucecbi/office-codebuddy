// Execute production pane source in a synthetic DOM/Office VM. No native UI,
// model, daemon, sockets or application processes are started by these tests.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { randomUUID } from "node:crypto";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";

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
    this.open = false;
    this.isConnected = true;
    this.parentNode = null;
    this.scrollHeight = this.scrollTop = this.clientHeight = 0;
  }
  addEventListener(type, fn) {
    this.listeners[type] = fn;
  }
  appendChild(el) {
    el.remove();
    this.children.push(el);
    el.parentNode = this;
    el.isConnected = this.isConnected;
    return el;
  }
  append(...els) {
    for (const el of els) this.appendChild(el);
  }
  prepend(el) {
    el.remove();
    this.children.unshift(el);
    el.parentNode = this;
    el.isConnected = this.isConnected;
  }
  remove() {
    if (this.parentNode) {
      const siblings = this.parentNode.children;
      siblings.splice(siblings.indexOf(this), 1);
    }
    this.parentNode = null;
    this.isConnected = false;
  }
  querySelectorAll() {
    return [];
  }
  querySelector(selector) {
    if (!this.queries.has(selector)) this.queries.set(selector, new Element());
    return this.queries.get(selector);
  }
  setAttribute() {}
  focus() {}
  closest() {
    return new Element();
  }
  set innerHTML(value) {
    this.html = value;
    for (const child of [...this.children]) child.remove();
  }
  get innerHTML() {
    return this.html || "";
  }
}

const flush = async () => {
  for (let i = 0; i < 16; i++) await Promise.resolve();
};

async function pane({ host = "word", stored = new Map(), apply } = {}) {
  const elements = new Map();
  const storage = {
    getItem: (key) => stored.get(key) ?? null,
    setItem: (k, v) => stored.set(k, v),
  };
  if (!stored.has("draftspect-preview-settings-v1"))
    stored.set(
      "draftspect-preview-settings-v1",
      JSON.stringify({
        model: "synthetic-model-a",
        modelsByRuntime: { codebuddy: "synthetic-model-a" },
      }),
    );
  const body = new Element();
  body.dataset.host = host;
  const document = {
    body,
    documentElement: new Element(),
    visibilityState: "visible",
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, new Element());
      return elements.get(id);
    },
    createElement: (tag) => new Element(tag),
    querySelectorAll: () => [],
    addEventListener() {},
  };
  for (const id of ["health-notice", "health-notice-text", "health-notice-detail"])
    document.getElementById(id);
  elements.get("health-notice").hidden = true;
  const timers = new Map();
  let timerId = 0;
  const sent = [],
    sockets = [];
  const documentUrl = host === "excel" ? "synthetic.xlsx" : "synthetic.docx";
  const context = vm.createContext({
    document,
    location: { protocol: "https:" },
    localStorage: storage,
    sessionStorage: storage,
    Office: { context: { document: { url: documentUrl } }, onReady() {} },
    console: { warn() {} },
    crypto: { randomUUID },
    createOperationQueue,
    fetch: async () => ({ ok: true, text: async () => "synthetic-token" }),
    setTimeout(fn, ms) {
      const id = ++timerId;
      timers.set(id, { fn, ms });
      return id;
    },
    clearTimeout: (id) => timers.delete(id),
    resolveWordHeadingArgs: async (_name, args) => args,
    prepareReviewWrite: async () => ({ preview: { action: "synthetic" }, apply }),
    WebSocket: function () {
      const socket = {
        send(raw) {
          sent.push({ ...JSON.parse(raw), socket: sockets.indexOf(socket) });
        },
        close() {
          socket.onclose?.();
        },
      };
      sockets.push(socket);
      return socket;
    },
  });
  const source = (
    await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8")
  ).replace(/^import\s+[\s\S]*?;\s*$/gm, "");
  vm.runInContext(
    source +
      `\nglobalThis.api = {
        handleServerMessage, runOfficeTool, appendToolUse, updateToolResult,
        sendUserMessage, wsConnect, checkTurnHealth, cancelOfficeOperations,
        turn(id) { currentTurnId = id; beginTurn(); },
        requestTimer(type) { return [...pendingRequests.values()].find((p) => p.type === type)?.timer; },
        state() { return { inFlight: turnInFlight, turn: currentTurnId,
          ready: wsReady, healthBlocked, modelConfirmed, replayReady }; },
        card(id) { return toolCards.get(id); },
      };`,
    context,
    { filename: "actual-taskpane-disconnect-safety.js" },
  );
  context.api.wsConnect();
  sockets[0].onopen();
  return {
    api: context.api,
    elements,
    stored,
    sent,
    sockets,
    documentUrl,
    async timer(ms) {
      const entry = [...timers].find(([, timer]) => timer.ms === ms);
      assert.ok(entry, `expected timer ${ms}`);
      timers.delete(entry[0]);
      entry[1].fn();
      await flush();
    },
    async expireRequest(type) {
      const id = context.api.requestTimer(type);
      assert.ok(id, `expected pending ${type}`);
      const timer = timers.get(id);
      assert.ok(timer);
      timers.delete(id);
      timer.fn();
      await flush();
    },
  };
}

async function handshake(p, { epoch = "epoch-A", replay = true, model = true } = {}) {
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: epoch,
    capabilities: ["turn_health_v1", "confirm_window_v1"],
  });
  const catalog = p.sent.findLast((m) => m.type === "list_models");
  await p.api.handleServerMessage({
    type: "list_models_result",
    request_id: catalog.request_id,
    ok: true,
    runtime: "codebuddy",
    models: [{ value: "synthetic-model-a" }],
  });
  await flush();
  if (replay) await p.api.handleServerMessage({ type: "transcript_replay", events: [] });
  if (model) await acknowledgeModel(p);
}
async function acknowledgeModel(p) {
  const choice = p.sent.findLast((m) => m.type === "set_model");
  assert.ok(choice);
  await p.api.handleServerMessage({
    type: "set_model_result",
    request_id: choice.request_id,
    ok: true,
    model: choice.model,
  });
  await flush();
}
async function acknowledgeUser(p) {
  const request = p.sent.findLast((m) => m.type === "user_message");
  await p.api.handleServerMessage({
    type: "user_message_result",
    request_id: request.request_id,
    turn_id: request.turn_id,
    connection_epoch: "epoch-A",
    ok: true,
  });
  await flush();
  return request;
}

for (const host of ["word", "excel"]) {
  test(`${host} disconnect exposes a neutral fault card and settles only unconfirmed records`, async () => {
    const p = await pane({ host });
    await handshake(p);
    p.api.turn("interrupted-read");
    const completed = p.api.appendToolUse("read completed", {}, "completed");
    p.api.updateToolResult("completed", "verified before loss", "completed");
    const pending = p.api.appendToolUse("read pending", {}, "pending");
    p.sockets[0].onclose({ reason: "PRIVATE_SECRET is not a diagnosis" });
    assert.equal(p.api.state().inFlight, false);
    assert.equal(p.elements.get("health-notice").hidden, false);
    assert.match(p.elements.get("health-notice-detail").textContent, /^CONNECTION_UNAVAILABLE/);
    assert.match(p.elements.get("health-notice-text").textContent, /连接.*不可用/);
    assert.doesNotMatch(
      p.elements.get("health-notice-text").textContent,
      /已崩溃|确认崩溃|PRIVATE_SECRET/,
    );
    assert.equal(pending.state, "interrupted / outcome unknown");
    assert.equal(pending.group.details.open, false);
    assert.equal(pending.group.alert.hidden, false);
    assert.doesNotMatch(pending.group.summary.textContent, /进行中/);
    assert.equal(completed.state, "completed");
    assert.equal(p.elements.get("send").disabled, true);
    assert.equal(p.elements.get("stop-agent").hidden, true);
    assert.equal(p.elements.get("agent-status").className, "agent-status err");
    assert.doesNotMatch(p.elements.get("agent-status").textContent, /Ready|Working|可发送/);
    assert.equal(p.stored.get(`draftspect-inspect-before-write:${p.documentUrl}`), undefined);
  });

  test(`${host} recovery waits for fresh model acknowledgement and history without replaying a request`, async () => {
    const p = await pane({ host });
    await handshake(p);
    assert.equal(p.api.sendUserMessage("Synthetic read; do not edit"), true);
    const original = await acknowledgeUser(p);
    p.sockets[0].onclose();
    await p.timer(1500); // The real production reconnect callback fetches a new token.
    assert.equal(p.sockets.length, 2);
    assert.equal(p.elements.get("send").disabled, true);
    p.sockets[1].onopen();
    await handshake(p, { epoch: "epoch-B", replay: false, model: false });
    assert.equal(p.elements.get("send").disabled, true);
    await acknowledgeModel(p);
    assert.equal(p.elements.get("send").disabled, true);
    assert.doesNotMatch(
      p.elements.get("health-notice-detail").textContent,
      /^CONNECTION_RECOVERED/,
    );
    await p.api.handleServerMessage({ type: "transcript_replay", events: [] });
    assert.equal(p.elements.get("send").disabled, false);
    assert.match(p.elements.get("agent-status").textContent, /可发送/);
    assert.match(p.elements.get("health-notice-detail").textContent, /^CONNECTION_RECOVERED/);
    assert.match(p.elements.get("health-notice-text").textContent, /未被自动重发|不会自动重发/);
    assert.match(p.elements.get("health-notice-text").textContent, /手动|重新提出/);
    assert.equal(p.sent.filter((m) => m.type === "user_message").length, 1);
    assert.ok(!p.sent.some((m) => m.type === "tool_result"));
    await p.api.handleServerMessage({
      type: "assistant_event",
      event: "turn_complete",
      turn_id: original.turn_id,
    });
    assert.equal(p.api.state().turn, null);
    p.sockets[0].onclose(); // A stale socket cannot destroy the recovered connection.
    assert.equal(p.api.state().ready, true);
  });

  test(`${host} idle connection loss recovers without leaving a stale failure card`, async () => {
    const p = await pane({ host });
    await handshake(p);
    p.sockets[0].onclose();
    assert.equal(p.elements.get("health-notice").hidden, false);
    await p.timer(1500);
    p.sockets[1].onopen();
    await handshake(p, { epoch: "epoch-B" });
    assert.equal(p.elements.get("health-notice").hidden, true);
    assert.equal(p.sent.filter((m) => m.type === "user_message").length, 0);
  });
}

test("unresponsive service is an error state while its unknown request remains locked and stoppable", async () => {
  const p = await pane();
  await handshake(p);
  p.api.turn("waiting-on-service");
  for (let i = 0; i < 2; i++) {
    const check = p.api.checkTurnHealth();
    await p.expireRequest("get_turn_state");
    await check;
  }
  assert.equal(p.api.state().inFlight, true, "no terminal result or confirmed stop was received");
  assert.equal(p.api.state().healthBlocked, true);
  assert.equal(p.elements.get("agent-status").className, "agent-status err");
  assert.match(p.elements.get("agent-status").textContent, /未响应/);
  assert.equal(p.elements.get("stop-agent").hidden, false);
  assert.equal(p.elements.get("send").disabled, true);
  assert.match(p.elements.get("health-notice-detail").textContent, /^SERVICE_UNRESPONSIVE/);
  assert.doesNotMatch(p.elements.get("health-notice-text").textContent, /已崩溃|确认崩溃/);
  assert.equal(p.sent.filter((m) => m.type === "user_message").length, 0);
});

test("a model-catalog RPC rejected by disconnect cannot overwrite the connection fault diagnosis", async () => {
  const p = await pane();
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    connection_epoch: "epoch-A",
    capabilities: ["turn_health_v1"],
  });
  assert.ok(p.sent.some((m) => m.type === "list_models"));
  p.sockets[0].onclose();
  await flush();
  assert.match(p.elements.get("health-notice-detail").textContent, /^CONNECTION_UNAVAILABLE/);
  assert.match(p.elements.get("health-notice-text").textContent, /原因尚未确认/);
  assert.equal(p.elements.get("send").disabled, true);
  assert.equal(p.sent.filter((m) => m.type === "user_message").length, 0);
});

test("disconnect before approval cancels the proposal without claiming a write occurred", async () => {
  let writes = 0;
  const p = await pane({
    apply: async ({ beforeWrite }) => {
      beforeWrite();
      writes++;
    },
  });
  await handshake(p);
  p.api.turn("proposal-only");
  const running = p.api.runOfficeTool({
    id: "unapproved-write",
    name: "office_edit_range",
    args: { target: { query: "Fixture" }, action: "replace", text: "Changed" },
    turn_id: "proposal-only",
    deadline: Date.now() + 120000,
  });
  await flush();
  assert.match(p.api.card("unapproved-write").state, /awaiting your confirmation/);
  p.sockets[0].onclose();
  await running;
  assert.equal(writes, 0);
  assert.equal(p.api.card("unapproved-write").state, "not applied / error");
  assert.equal(p.stored.get("draftspect-inspect-before-write:synthetic.docx"), undefined);
  assert.ok(!p.sent.some((m) => m.type === "tool_result" && m.ok));
});

test("disconnect during an Office commit preserves unknown outcome through late return and reload", async () => {
  const stored = new Map();
  let writes = 0,
    release;
  const p = await pane({
    stored,
    apply: async ({ beforeWrite }) => {
      beforeWrite();
      writes++;
      await new Promise((resolve) => {
        release = resolve;
      });
      return { verified: true };
    },
  });
  await handshake(p);
  p.api.turn("in-write");
  const msg = {
    id: "attempted-write",
    name: "office_edit_range",
    args: { target: { query: "Fixture" }, action: "replace", text: "Changed" },
    turn_id: "in-write",
    deadline: Date.now() + 120000,
  };
  const running = p.api.runOfficeTool(msg);
  await flush();
  p.api
    .card(msg.id)
    .el.children.find((el) => el.textContent === "Apply this change")
    .listeners.click();
  await flush();
  assert.equal(writes, 1);
  p.sockets[0].onclose();
  assert.equal(stored.get("draftspect-inspect-before-write:synthetic.docx"), "unknown");
  assert.match(p.elements.get("agent-status").textContent, /修改结果待核实/);
  release();
  await running;
  assert.equal(p.api.card(msg.id).state, "outcome unknown — inspect Word");
  assert.ok(!p.sent.some((m) => m.type === "tool_result" && m.ok));
  const reloaded = await pane({
    stored,
    apply: async () => {
      writes++;
    },
  });
  await handshake(reloaded);
  assert.match(reloaded.elements.get("agent-status").textContent, /修改结果待核实/);
  reloaded.api.turn("fresh-proposal");
  await reloaded.api.runOfficeTool({ ...msg, turn_id: "fresh-proposal" });
  await reloaded.api.runOfficeTool({ ...msg, id: "fresh-id", turn_id: "fresh-proposal" });
  assert.equal(writes, 1);
  assert.ok(
    reloaded.sent
      .filter((m) => m.type === "tool_result")
      .every((m) => !m.ok && /RESULT_UNKNOWN/.test(m.error)),
  );
  const inspection = reloaded.elements
    .get("messages")
    .children.find((el) => /prior edit/.test(el.textContent || ""));
  assert.match(inspection.textContent, /reread|re-read/i);
  assert.equal(stored.get("draftspect-operation:attempted-write"), "attempted");
});

test("fault disclosures are closed by default in both native-host entry points", async () => {
  for (const host of ["word", "excel"]) {
    const html = await readFile(new URL(`../taskpane/${host}/index.html`, import.meta.url), "utf8");
    const notice = html.match(/<section id="health-notice"[\s\S]*?<\/section>/)?.[0];
    assert.ok(notice);
    assert.match(notice, /<details>\s*<summary>详情<\/summary>/);
    assert.doesNotMatch(notice, /<details[^>]*\bopen\b/);
    assert.match(notice, /id="health-reconnect"/);
  }
});
