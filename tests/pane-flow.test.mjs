import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { randomUUID } from "node:crypto";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";
import { planReplacement } from "../taskpane/shared/word-edit-plan.js";
import { toolSetFont } from "../taskpane/shared/tools-word.js";

// Execute the actual pane script against a tiny synthetic DOM. No webview/Word.
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
    this.parentNode = null;
    this.open = false;
    this.scrollHeight = 0;
    this.scrollTop = 0;
    this.clientHeight = 0;
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
  set innerHTML(value) {
    this.html = value;
    for (const child of [...this.children]) child.remove();
  }
  get innerHTML() {
    return this.html || "";
  }
}
async function pane({
  host = "word",
  reviewApply = async () => ({ verified: true }),
  globals = {},
} = {}) {
  const elements = new Map(),
    stored = new Map();
  stored.set(
    "draftspect-preview-settings-v1",
    JSON.stringify({
      model: "synthetic-model-a",
      modelsByRuntime: { codebuddy: "synthetic-model-a" },
    }),
  );
  const storage = { getItem: (k) => stored.get(k) ?? null, setItem: (k, v) => stored.set(k, v) };
  const body = new Element();
  body.dataset.host = host;
  const document = {
    body,
    documentElement: new Element(),
    getElementById: (id) => {
      if (!elements.has(id)) elements.set(id, new Element());
      return elements.get(id);
    },
    querySelectorAll: () => [],
    createElement: (tag) => new Element(tag),
    addEventListener() {},
  };
  const sent = [];
  let headerWrites = 0;
  const socket = { send: (s) => sent.push(JSON.parse(s)), close() {} };
  const context = vm.createContext({
    location: { protocol: "http:" },
    document,
    console,
    crypto: { randomUUID },
    localStorage: storage,
    sessionStorage: storage,
    Office: { context: { document: { url: "fixture.docx" } }, onReady() {} },
    createOperationQueue,
    planReplacement,
    prepareReviewWrite: async () => ({ preview: { action: "synthetic" }, apply: reviewApply }),
    prepareSafeWordWrite: async (name) =>
      name === "office_header_footer"
        ? {
            preview: { action: "synthetic header" },
            apply: async () => {
              headerWrites++;
            },
          }
        : null,
    prepareTextReplacement: async () => [],
    captureWordFormattingPrecondition: async () => null,
    resolveWordHeadingArgs: async (_name, args) => args,
    setTimeout: (...args) => {
      const h = setTimeout(...args);
      h.unref();
      return h;
    },
    clearTimeout,
    toolReadParagraphs: async () => ({
      addressing: "uniqueLocalId",
      paragraphs: [{ id: "P1", text: "Fixture" }],
    }),
    toolHeaderFooter: async () => {
      headerWrites++;
    },
    socket,
    WebSocket: function () {
      return socket;
    },
    ...globals,
  });
  const source = (
    await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8")
  ).replace(/^import\s+[\s\S]*?;\s*$/gm, "");
  vm.runInContext(
    source +
      `\nglobalThis.api = {
    handleServerMessage, runOfficeTool, cancelOfficeOperations, appendToolUse, updateToolResult,
    appendUserMessage, appendAssistantDelta, upsertAssistantMessage,
    startConnection: wsConnect,
    connect() { ws = socket; wsReady = true; activeDocUrl = 'fixture.docx'; },
    turn(id) { currentTurnId = id; beginTurn(); },
    state() { return { model: settings.model, confirmed: modelConfirmed, replay: replayReady, inFlight: turnInFlight, released: sessionReleased, sessionInfo: currentSessionInfo, settingsKey: SETTINGS_KEY }; },
    card(id) { return toolCards.get(id); },
  };`,
    context,
    { filename: "taskpane.js" },
  );
  context.api.connect();
  return { api: context.api, elements, stored, sent, socket, headerWrites: () => headerWrites };
}
const flush = async () => {
  for (let i = 0; i < 8; i++) await Promise.resolve();
};

async function confirmReady(p, welcome = {}) {
  await p.api.handleServerMessage({
    type: "welcome",
    protocol_version: 2,
    runtime: "codebuddy",
    ...welcome,
  });
  await p.api.handleServerMessage({ type: "transcript_replay", events: [] });
  await p.api.handleServerMessage({
    type: "list_models_result",
    ok: true,
    runtime: "codebuddy",
    models: [{ value: "synthetic-model-a" }],
  });
  const request = p.sent.findLast((m) => m.type === "set_model");
  await p.api.handleServerMessage({
    type: "set_model_result",
    request_id: request.request_id,
    ok: true,
    model: request.model,
  });
  await flush();
}

for (const host of ["word", "excel"]) {
  test(`${host} footer loses readiness on disconnect and waits for a fresh model acknowledgement`, async () => {
    const p = await pane({ host });
    p.api.startConnection();
    p.socket.onopen();
    await confirmReady(p);
    const status = p.elements.get("agent-status");
    assert.match(status.textContent, /可发送/);
    assert.equal(p.elements.get("send").disabled, false);

    p.socket.onclose();
    assert.match(status.textContent, /连接已断开/);
    assert.equal(p.elements.get("send").disabled, true);
    assert.equal(p.elements.get("stop-agent").hidden, true);
    await flush();
    assert.doesNotMatch(status.textContent, /可发送|Working/);

    p.api.startConnection();
    assert.match(status.textContent, /正在连接/);
    p.socket.onopen();
    assert.match(status.textContent, /等待模型确认/);
    assert.equal(p.elements.get("send").disabled, true);
    await confirmReady(p);
    assert.match(status.textContent, /可发送/);
    assert.equal(p.elements.get("send").disabled, false);
  });

  test(`${host} disconnect clears stale Working but preserves an unverified edit warning`, async () => {
    const p = await pane({ host });
    p.api.startConnection();
    p.socket.onopen();
    await confirmReady(p);
    p.api.turn("working-disconnect");
    await p.api.handleServerMessage({
      type: "assistant_event",
      event: "tool_use_announce",
      turn_id: "working-disconnect",
      tool_id: "read",
      tool: "office_read_paragraphs",
      input: {},
    });
    assert.equal(p.elements.get("agent-status").className, "agent-status working");
    p.socket.onclose();
    assert.equal(p.api.state().inFlight, false);
    assert.match(p.elements.get("agent-status").textContent, /连接已断开/);
    assert.equal(p.elements.get("stop-agent").hidden, true);

    p.stored.set("draftspect-inspect-before-write:fixture.docx", "unknown");
    p.api.startConnection();
    p.socket.onopen();
    p.socket.onclose();
    assert.match(p.elements.get("agent-status").textContent, /修改结果待核实/);
    assert.match(
      p.elements.get("agent-status").textContent,
      host === "excel" ? /请检查 Excel/ : /请检查 Word/,
    );
    assert.equal(p.elements.get("agent-status").className, "agent-status err");
    assert.equal(p.elements.get("send").disabled, true);
  });

  test(`${host} composer copy follows the daemon runtime and resets on disconnect`, async () => {
    const p = await pane({ host });
    const input = p.elements.get("input");
    const model = p.elements.get("composer-model");
    input.value = "Keep my unsent draft";
    assert.equal(input.placeholder, "Ask your agent…");
    for (const [runtime, label] of [
      ["codebuddy", "CodeBuddy"],
      ["claude", "Claude"],
    ]) {
      await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime });
      assert.equal(input.placeholder, `Ask ${label}…`);
      assert.equal(model.title, `Model — provided by ${label}`);
    }
    await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime: "unknown" });
    assert.equal(input.placeholder, "Ask your agent…");
    p.api.startConnection();
    p.socket.onopen();
    await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime: "codebuddy" });
    p.socket.onclose();
    assert.equal(input.placeholder, "Ask your agent…");
    assert.equal(model.title, "Model — provided by the connected runtime");
    assert.equal(input.value, "Keep my unsent draft");
    assert.equal(p.api.state().model, "synthetic-model-a");
    assert.ok(!p.sent.some((m) => m.type === "user_message"));
  });

  test(`${host} static editor and setup copy is provider-neutral`, async () => {
    const html = await readFile(new URL(`../taskpane/${host}/index.html`, import.meta.url), "utf8");
    assert.match(html, /id="input" placeholder="Ask your agent…"/);
    assert.match(html, /placeholder="Instructions for the agent\. Can be multi-line\."/);
    assert.doesNotMatch(
      html,
      /Ask Claude|send to Claude|folder Claude|Claude can|Point Claude|served by your legacy gateway/,
    );
    // Real model IDs and the context-storage filename are not branding copy.
    assert.match(html, /Connect to load available models/);
    assert.doesNotMatch(html, /value="sonnet"/);
    assert.match(html, /CLAUDE\.md/);
  });
}

test("auth guidance follows the current runtime, including an existing banner", async () => {
  const p = await pane();
  for (const runtime of ["claude", "codebuddy", "unknown"]) {
    await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime });
    await p.api.handleServerMessage({
      type: "assistant_event",
      event: "auth_error",
      error: "Synthetic sign-in failure",
    });
    const banner = p.elements.get("auth-error-banner");
    const body = banner.querySelector(".auth-error-body").innerHTML;
    assert.equal(
      banner.querySelector(".auth-error-raw pre").textContent,
      "Synthetic sign-in failure",
    );
    if (runtime === "claude") assert.match(body, /<code>claude<\/code>/);
    else {
      assert.doesNotMatch(body, /Anthropic|ANTHROPIC_API_KEY|Claude Code|<code>claude<\/code>/);
      if (runtime === "codebuddy") assert.match(body, /<code>codebuddy<\/code>/);
    }
  }
});

test("pane preserves CodeBuddy choice and gates sending on replay plus model acknowledgement", async () => {
  const p = await pane();
  assert.equal(p.api.state().model, "synthetic-model-a");
  await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime: "codebuddy" });
  await p.api.handleServerMessage({ type: "transcript_replay", events: [] });
  await p.api.handleServerMessage({
    type: "list_models_result",
    ok: true,
    runtime: "codebuddy",
    models: [{ value: "synthetic-model-a" }],
  });
  assert.equal(p.api.state().confirmed, false);
  assert.equal(p.elements.get("send").disabled, true);
  const request = p.sent.find((m) => m.type === "set_model");
  assert.equal(request.model, "synthetic-model-a");
  await p.api.handleServerMessage({
    type: "set_model_result",
    request_id: request.request_id,
    ok: true,
    model: request.model,
  });
  await flush();
  assert.equal(p.api.state().confirmed, true);
  assert.equal(p.elements.get("send").disabled, false);
  assert.ok(p.stored.has("draftspect-preview-settings-v1"));
  assert.ok(!p.stored.has("claude-code-office-settings-v1"));
});

test("foreign model requires a choice; dropdown stays usable while Send is blocked", async () => {
  const p = await pane();
  await p.api.handleServerMessage({ type: "welcome", protocol_version: 2, runtime: "claude" });
  await p.api.handleServerMessage({
    type: "list_models_result",
    ok: true,
    runtime: "claude",
    models: [{ value: "sonnet" }],
  });
  assert.equal(p.api.state().confirmed, false);
  assert.match(p.elements.get("health-notice-text").textContent, /重新选择可用模型/);
  assert.match(p.elements.get("agent-status").textContent, /请选择可用模型/);
  assert.equal(p.elements.get("composer-model").disabled, false);
  assert.equal(p.elements.get("composer-model").value, "");
  assert.ok(!p.sent.some((m) => m.type === "set_model"));
});

test("tool details are expandable and untruncated; a terminal error cannot become Ready", async () => {
  const p = await pane();
  p.api.turn("t");
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "tool_use_announce",
    turn_id: "t",
    tool_id: "tool",
    tool: "Bash",
    input: { command: "x".repeat(1000) },
  });
  const card = p.api.card("tool");
  assert.equal(card.el.tagName, "details");
  assert.ok(card.el.children[1].textContent.length > 1000);
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "tool_result_announce",
    turn_id: "t",
    tool_id: "tool",
    output: "ok",
    status: "completed",
  });
  assert.equal(card.state, "completed");
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: "t",
    subtype: "error_max_turns",
    error: "limit details",
  });
  assert.equal(p.api.state().inFlight, false);
  assert.equal(p.elements.get("agent-status").textContent, "Stopped — see message");
  assert.ok(p.elements.get("messages").children.some((c) => c.textContent === "limit details"));
  const css = await readFile(new URL("../taskpane/shared/styles.css", import.meta.url), "utf8");
  assert.ok(!/data-show-diagnostics[^}]+\.msg\.tool/.test(css));
});

for (const host of ["word", "excel"]) {
  test(`${host} groups SDK and Office records once per turn, closed by default`, async () => {
    const p = await pane({ host });
    p.api.appendUserMessage("Synthetic question");
    p.api.turn("grouped");
    for (let i = 0; i < 24; i++) {
      const id = `tool-${i}`;
      p.api.appendToolUse(
        i % 2 ? "Office execution: office_find" : "mcp__office__office_find",
        { token: "secret", query: "x".repeat(1000) },
        id,
      );
      p.api.updateToolResult(id, "y".repeat(1200), "completed");
    }
    const group = p.api.card("tool-0").group;
    assert.equal(group.details.open, false);
    assert.equal(group.cards.size, 24);
    assert.equal(group.body.children.length, 24);
    assert.match(group.summary.textContent, /24 条记录/);
    assert.equal(group.alert.hidden, true);
    assert.equal(
      p.elements.get("messages").children.filter((el) => el.className === "process-turn").length,
      1,
    );
    assert.equal(
      p.elements.get("messages").children.filter((el) => el.className === "msg tool").length,
      0,
    );
    assert.match(p.api.card("tool-0").el.children[1].textContent, /\[redacted\]/);
    assert.ok(p.api.card("tool-0").output.textContent.length > 1200);
    p.api.appendToolUse("duplicate", {}, "tool-0");
    assert.equal(group.cards.size, 24);
    p.api.upsertAssistantMessage("answer", "Synthetic answer remains visible");
    const answer = p.elements.get("messages").children.at(-1);
    assert.equal(answer.className, "msg assistant");
    assert.equal(answer.hidden, false);
    group.details.open = true;
    p.api.appendToolUse("next", {}, "next");
    p.api.updateToolResult("next", "ok", "completed");
    assert.equal(group.details.open, true);
    await p.api.handleServerMessage({
      type: "assistant_event",
      event: "turn_complete",
      turn_id: "grouped",
    });
    assert.equal(group.details.open, true);
    p.api.appendUserMessage("Next question");
    const next = p.api.appendToolUse("next turn", {}, "next-turn");
    assert.notEqual(next.group, group);
    assert.equal(next.group.details.open, false);
  });
}

test("late results update their owning group without affecting a newer turn", async () => {
  const p = await pane();
  const old = p.api.appendToolUse("old", {}, "old");
  p.api.appendUserMessage("New question");
  const fresh = p.api.appendToolUse("fresh", {}, "fresh");
  p.api.updateToolResult("old", "late failure", "error");
  assert.match(old.group.summary.textContent, /1 项异常/);
  assert.equal(old.group.alert.hidden, false);
  assert.equal(fresh.group.alert.hidden, true);
  assert.match(fresh.group.summary.textContent, /1 项进行中/);
  assert.equal(old.el.parentNode, old.group.body);
});

test("provider errors and unconfirmed results have one visible alert outside closed details", async () => {
  const p = await pane();
  for (const [id, status] of [
    ["error", "error"],
    ["unknown", "outcome unknown — inspect Word"],
  ]) {
    p.api.appendToolUse("synthetic", {}, id);
    await p.api.handleServerMessage({
      type: "assistant_event",
      event: "tool_result_announce",
      tool_id: id,
      output: "not successful",
      status,
    });
  }
  const group = p.api.card("error").group;
  assert.equal(group.details.open, false);
  assert.equal(group.alert.parentNode, group.el);
  assert.equal(group.alert.hidden, false);
  assert.match(group.alertText.textContent, /异常.*未确认/);
  assert.equal(group.el.children.filter((el) => el.className.includes("process-alert")).length, 1);
  group.alert.children[1].listeners.click();
  assert.equal(group.details.open, true);
});

test("stopping a running tool preserves a visible unknown-result warning", async () => {
  const p = await pane();
  p.api.turn("stop");
  const card = p.api.appendToolUse("pending", {}, "pending");
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: "stop",
    interrupted: true,
  });
  assert.equal(card.group.details.open, false);
  assert.equal(card.group.alert.hidden, false);
  assert.match(card.group.summary.textContent, /结果未确认/);
  assert.doesNotMatch(card.group.summary.textContent, /记录已完成/);
});

test("missing tool-result status stays explicitly unconfirmed instead of breaking replay", async () => {
  const p = await pane();
  p.api.updateToolResult("missing", "legacy output", undefined);
  const card = p.api.card("missing");
  assert.equal(card.state, "result not reported");
  assert.equal(card.group.alert.hidden, false);
  assert.match(card.group.summary.textContent, /结果未确认/);
});

test("replay groups by user turn, resets prior groups and handles an orphan result", async () => {
  const p = await pane();
  p.api.appendToolUse("obsolete", {}, "obsolete");
  const events = [
    { kind: "tool_result", id: "orphan", output: "ok", status: "completed" },
    { kind: "user", text: "Q1" },
    { kind: "tool", id: "first", name: "read", input: {} },
    { kind: "tool_result", id: "first", output: "ok", status: "completed" },
    { kind: "assistant", text: "A1" },
    { kind: "user", text: "Q2" },
    { kind: "tool", id: "second", name: "read", input: {} },
    { kind: "assistant", text: "A2" },
  ];
  for (let i = 0; i < 2; i++) {
    await p.api.handleServerMessage({ type: "transcript_replay", events, truncated: true });
    const children = p.elements.get("messages").children;
    assert.equal(children.filter((el) => el.className === "process-turn").length, 3);
    assert.equal(children.filter((el) => el.className === "msg assistant").length, 2);
    for (const id of ["orphan", "first", "second"])
      assert.equal(p.api.card(id).group.details.open, false);
    assert.equal(p.api.card("obsolete"), undefined);
    assert.equal(p.api.card("second").group.alert.hidden, false);
    assert.match(p.api.card("second").group.summary.textContent, /结果未确认/);
  }
  const live = p.api.appendToolUse("live", {}, "live");
  assert.notEqual(live.group, p.api.card("second").group);
});

test("empty assistant updates create no blank bubbles and substantive text stays outside logs", async () => {
  const p = await pane();
  p.api.upsertAssistantMessage("empty", "");
  p.api.appendAssistantDelta("");
  assert.equal(p.elements.get("messages").children.length, 0);
  p.api.upsertAssistantMessage("answer", "Visible text");
  const el = p.elements.get("messages").children[0];
  p.api.upsertAssistantMessage("answer", "");
  assert.equal(el.hidden, true);
  p.api.upsertAssistantMessage("answer", "Restored text");
  assert.equal(el.hidden, false);
  p.api.appendToolUse("read", {}, "read");
  assert.equal(el.parentNode, p.elements.get("messages"));
});

test("approval remains outside the collapsed log, then returns after a verified write", async () => {
  let applies = 0;
  const p = await pane({
    reviewApply: async () => {
      applies++;
      return { verified: true };
    },
  });
  p.api.turn("approve");
  p.api.appendToolUse("earlier read", {}, "earlier");
  p.api.updateToolResult("earlier", "ok", "completed");
  p.api.upsertAssistantMessage("analysis", "Long analysis\n".repeat(100));
  const running = p.api.runOfficeTool({
    id: "approve-format",
    name: "office_edit_range",
    args: { target: { query: "Fixture" }, action: "format", font: { size: 13 } },
    turn_id: "approve",
    deadline: Date.now() + 120000,
  });
  await flush();
  const card = p.api.card("approve-format");
  assert.match(card.state, /awaiting/);
  assert.equal(card.group.details.open, false);
  assert.equal(card.el.open, true);
  assert.equal(card.el.parentNode, p.elements.get("messages"));
  assert.equal(p.elements.get("messages").children.at(-1), card.el);
  card.el.children.find((el) => el.textContent === "Apply this change").listeners.click();
  await running;
  assert.equal(applies, 1);
  assert.equal(card.state, "completed");
  assert.equal(card.el.parentNode, card.group.body);
  assert.equal(card.group.details.open, false);
  assert.equal(
    card.el.children.some((el) => el.tagName === "button"),
    false,
  );
});

test("header/footer tools are writes even without text; cancelling confirmation never clears headers", async () => {
  const p = await pane();
  p.api.turn("t");
  const running = p.api.runOfficeTool({
    id: "header",
    name: "office_header_footer",
    args: { which: "header" },
    turn_id: "t",
    deadline: Date.now() + 120000,
  });
  await flush();
  assert.match(p.api.card("header").state, /awaiting/);
  const card = p.api.card("header");
  assert.equal(card.el.parentNode, p.elements.get("messages"));
  assert.equal(card.group.details.open, false);
  p.api.cancelOfficeOperations();
  await running;
  assert.equal(card.el.parentNode, card.group.body);
  assert.equal(card.group.alert.hidden, false);
  assert.equal(p.headerWrites(), 0);
  assert.equal(p.sent.at(-1).ok, false);
});

test("pane shows the full session identity and release state blocks composing after replay", async () => {
  const p = await pane();
  const id = "complete-session-identity-123456789";
  await p.api.handleServerMessage({
    type: "transcript_replay",
    events: [],
    session_info: {
      session_id: id,
      provider: "codebuddy",
      state: "released",
      commands: { resume: "resume fixture", reuse_as_new_session: "fork fixture" },
    },
  });
  assert.equal(p.api.state().released, true);
  assert.ok(p.elements.get("session-identity").textContent.includes(id));
  assert.equal(p.elements.get("send").disabled, true);
  assert.equal(p.elements.get("session-return").hidden, false);
  assert.ok(p.elements.get("session-commands").textContent.includes("resume fixture"));
  await p.api.handleServerMessage({
    type: "transcript_replay",
    events: [],
    session_info: { session_id: id, provider: "codebuddy", state: "word" },
  });
  assert.equal(p.api.state().released, false);
  assert.equal(p.elements.get("session-return").hidden, true);
});
test("changed CLI history stays released until a second deliberate return click sends the challenge", async () => {
  const p = await pane();
  const info = { session_id: "id", provider: "codebuddy", state: "released" };
  await p.api.handleServerMessage({ type: "transcript_replay", events: [], session_info: info });
  const button = p.elements.get("session-return");
  const first = button.listeners.click();
  await flush();
  const request = p.sent.findLast((m) => m.type === "return_session");
  assert.equal(request.cli_closed, true);
  assert.equal(request.confirm_transcript_change, undefined);
  await p.api.handleServerMessage({
    type: "return_session_result",
    request_id: request.request_id,
    ok: false,
    code: "TRANSCRIPT_CHANGED",
    error: "changed",
    info: { ...info, return_confirmation: "fresh-challenge" },
  });
  await first;
  await flush();
  assert.equal(p.api.state().released, true);
  assert.equal(p.elements.get("send").disabled, true);
  assert.match(button.textContent, /再次确认 CLI 已退出/);
  assert.equal(p.sent.filter((m) => m.type === "return_session").length, 1);
  const second = button.listeners.click();
  await flush();
  const confirmed = p.sent.findLast((m) => m.type === "return_session");
  assert.equal(confirmed.confirm_transcript_change, "fresh-challenge");
  await p.api.handleServerMessage({
    type: "return_session_result",
    request_id: confirmed.request_id,
    ok: true,
    info: { ...info, state: "word" },
  });
  await second;
  assert.equal(p.api.state().released, false);
});

test("new review writes remain behind the real pane confirmation and cancellation gate", async () => {
  let applies = 0;
  const p = await pane({
    reviewApply: async () => {
      applies++;
      return { verified: true };
    },
  });
  p.api.turn("review");
  const running = p.api.runOfficeTool({
    id: "format",
    name: "office_edit_range",
    args: { target: { query: "Fixture" }, action: "format", font: { bold: true } },
    turn_id: "review",
    deadline: Date.now() + 120000,
  });
  await flush();
  assert.match(p.api.card("format").state, /awaiting/);
  assert.equal(applies, 0);
  p.api.cancelOfficeOperations();
  await running;
  assert.equal(applies, 0);
  assert.equal(p.sent.at(-1).ok, false);
});

test("unconfirmed Stop keeps the turn until the daemon reports its state", async () => {
  async function stoppedPane({ health }) {
    const p = await pane();
    await confirmReady(
      p,
      health ? { connection_epoch: "e1", capabilities: ["turn_health_v1"] } : {},
    );
    p.api.turn("t1");
    p.elements.get("stop-agent").listeners.click();
    await flush();
    const stop = p.sent.findLast((m) => m.type === "stop_agent");
    await p.api.handleServerMessage({
      type: "stop_agent_result",
      request_id: stop.request_id,
      ok: false,
      error: "synthetic stop failure",
    });
    await flush();
    return p;
  }
  const reply = async (p, result) => {
    const poll = p.sent.findLast((m) => m.type === "get_turn_state");
    assert.ok(poll, "turn state is queried immediately");
    await p.api.handleServerMessage({
      type: "get_turn_state_result",
      request_id: poll.request_id,
      ok: true,
      connection_epoch: "e1",
      ...result,
    });
    await flush();
  };

  const running = await stoppedPane({ health: true });
  const now = Date.now();
  await reply(running, {
    active: { turn_id: "t1", phase: "streaming", started_at: now, last_progress_at: now },
  });
  assert.equal(running.api.state().inFlight, true);
  assert.equal(running.elements.get("stop-agent").hidden, false);
  assert.equal(running.elements.get("send").disabled, true);

  const finished = await stoppedPane({ health: true });
  assert.equal(finished.api.state().inFlight, true);
  await reply(finished, {
    active: null,
    terminal: { event: "turn_complete", turn_id: "t1", interrupted: true },
  });
  assert.equal(finished.api.state().inFlight, false);
  assert.equal(finished.elements.get("send").disabled, false);

  const legacy = await stoppedPane({ health: false });
  assert.ok(!legacy.sent.some((m) => m.type === "get_turn_state"));
  assert.equal(legacy.api.state().inFlight, false);
  assert.match(legacy.elements.get("agent-status").textContent, /Stop not confirmed/);
});

test("an auth-flagged turn_complete ends the turn and shows the sign-in banner", async () => {
  const p = await pane();
  await confirmReady(p);
  p.api.turn("t-auth");
  await p.api.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    turn_id: "t-auth",
    subtype: "runtime_error",
    error: "Authentication required",
    auth: true,
  });
  assert.equal(p.api.state().inFlight, false);
  const banner = p.elements.get("auth-error-banner");
  assert.equal(banner.querySelector(".auth-error-raw pre").textContent, "Authentication required");
  assert.match(banner.querySelector(".auth-error-body").innerHTML, /<code>codebuddy<\/code>/);
});

test("Excel writes pin the sheet and refuse when target cells change during confirmation", async () => {
  const cells = { value: [["old"]] };
  const writes = [];
  const p = await pane({
    host: "excel",
    globals: {
      captureExcelPrecondition: async (_name, args) => ({
        args: { ...args, sheet: args.sheet ?? "Sheet1" },
        state: { used: "Sheet1!A1:A1", formulas: structuredClone(cells.value) },
        preview: {
          sheet: "Sheet1",
          target_address: "Sheet1!A1:A1",
          current_contents: {
            values: structuredClone(cells.value),
            formulas: structuredClone(cells.value),
          },
        },
      }),
      toolExcelWriteRange: async (args) => {
        writes.push(args);
        return { written: 1 };
      },
    },
  });
  const run = async (id, change) => {
    p.api.turn(id);
    const running = p.api.runOfficeTool({
      id,
      name: "excel_write_range",
      args: { address: "A1", values: [["new"]] },
      turn_id: id,
      deadline: Date.now() + 120000,
    });
    await flush();
    const card = p.api.card(id);
    assert.match(card.state, /awaiting/);
    assert.match(card.output.textContent, /current_contents/);
    assert.ok(card.output.textContent.includes(cells.value[0][0]));
    change?.();
    card.el.children.find((el) => el.textContent === "Apply this change").listeners.click();
    await running;
    await flush();
    return p.sent.findLast((m) => m.type === "tool_result" && m.id === id);
  };

  const stale = await run("stale", () => {
    cells.value = [["typed by user"]];
  });
  assert.equal(stale.ok, false);
  assert.match(stale.error, /STALE_PROPOSAL/);
  assert.equal(writes.length, 0);
  assert.match(p.elements.get("health-notice-text").textContent, /^Excel 操作未完成/);

  const fresh = await run("fresh");
  assert.equal(fresh.ok, true);
  assert.deepEqual(writes, [{ address: "A1", values: [["new"]], sheet: "Sheet1" }]);
});

test("generic Word approval rejects a format-only change even when paragraph text and style are unchanged", async () => {
  let formatting = "plain",
    writes = 0;
  const p = await pane({
    globals: {
      captureWordFormattingPrecondition: async () => ({
        targets: [{ key: "P1", ooxml: formatting }],
      }),
      toolSetFont: async () => {
        writes++;
        return { formatted_count: 1 };
      },
    },
  });
  p.api.turn("format-stale");
  const running = p.api.runOfficeTool({
    id: "format-stale",
    name: "office_set_font",
    args: { ids: ["P1"], bold: true },
    turn_id: "format-stale",
    deadline: Date.now() + 120000,
  });
  await flush();
  await flush();
  const card = p.api.card("format-stale");
  assert.match(card.state, /awaiting/);
  formatting = "italic changed externally";
  card.el.children.find((el) => el.textContent === "Apply this change").listeners.click();
  await running;
  const result = p.sent.findLast((m) => m.type === "tool_result" && m.id === "format-stale");
  assert.equal(result.ok, false);
  assert.match(result.error, /STALE_PROPOSAL/);
  assert.equal(writes, 0);
});

test("with confirm_window_v1 slow preview time does not consume the approval window", async () => {
  const run = async (capabilities) => {
    const p = await pane();
    await confirmReady(p, { connection_epoch: "e1", capabilities });
    p.api.turn("slow");
    const running = p.api.runOfficeTool({
      id: "slow-header",
      name: "office_header_footer",
      args: { which: "header" },
      turn_id: "slow",
      // Preview preparation already used almost the whole fixed budget.
      deadline: Date.now() + 40,
    });
    await flush();
    return { p, running };
  };

  const current = await run(["turn_health_v1", "confirm_window_v1"]);
  const card = current.p.api.card("slow-header");
  assert.match(card.state, /awaiting your confirmation \(45 seconds\)/);
  await new Promise((r) => setTimeout(r, 80));
  card.el.children.find((el) => el.textContent === "Apply this change").listeners.click();
  await current.running;
  await flush();
  assert.equal(current.p.headerWrites(), 1);
  assert.equal(current.p.sent.findLast((m) => m.type === "tool_result").ok, true);

  const legacy = await run(["turn_health_v1"]);
  await legacy.running;
  await flush();
  const result = legacy.p.sent.findLast((m) => m.type === "tool_result");
  assert.equal(result.ok, false);
  assert.match(result.error, /CONFIRMATION_EXPIRED/);
  assert.equal(legacy.p.headerWrites(), 0);
});

test("cancelling during a generic font load does not write or report success", async () => {
  let paints = 0;
  const paragraph = {
    text: "Synthetic",
    style: "Normal",
    styleBuiltIn: "Normal",
    uniqueLocalId: "ID0",
    listItemOrNullObject: { isNullObject: true },
    load() {},
    getRange() {
      return {
        font: {
          set bold(value) {
            paints++;
            void value;
          },
        },
        load() {},
      };
    },
  };
  const document = {
    changeTrackingMode: "Off",
    load() {},
    body: { paragraphs: { items: [paragraph], load() {} } },
    getSelection: () => ({ select() {} }),
  };
  let release;
  const gate = new Promise((resolve) => {
    release = resolve;
  });
  let started;
  const entered = new Promise((resolve) => {
    started = resolve;
  });
  let syncs = 0;
  const previous = globalThis.Word;
  globalThis.Word = {
    run: (fn) =>
      fn({
        document,
        sync: async () => {
          if (++syncs === 1) {
            started();
            await gate;
          }
        },
      }),
    ChangeTrackingMode: { trackAll: "TrackAll" },
    ChangeTrackingVersion: { current: "Current" },
  };
  try {
    const p = await pane({ globals: { toolSetFont } });
    p.api.turn("review-turn");
    const running = p.api.runOfficeTool({
      id: "cancel-gap",
      name: "office_set_font",
      args: { ids: ["ID0"], bold: true },
      turn_id: "review-turn",
      deadline: Date.now() + 120000,
    });
    let button;
    for (let attempt = 0; attempt < 20 && !button; attempt++) {
      await flush();
      button = p.api
        .card("cancel-gap")
        ?.el.children.find((el) => el.textContent === "Apply this change");
      if (!button) await new Promise((resolve) => setImmediate(resolve));
    }
    assert.ok(button);
    button.listeners.click();
    await entered;
    p.api.cancelOfficeOperations();
    release();
    await running;
    assert.equal(paints, 0);
    const result = p.sent.find(
      (message) => message.type === "tool_result" && message.id === "cancel-gap",
    );
    assert.equal(result.ok, false);
    assert.match(result.error, /CANCELLED_OR_EXPIRED/);
    assert.doesNotMatch(result.error, /RESULT_UNKNOWN/);
  } finally {
    globalThis.Word = previous;
  }
});
