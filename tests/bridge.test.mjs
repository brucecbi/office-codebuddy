// Unit test for daemon/bridge.mjs network binding.
//
// Regression guard for the WS loopback-bind security fix: the `ws`
// library binds 0.0.0.0 when constructed with only `port`, exposing the
// agent bridge to the local network. createBridge must pass
// host: "127.0.0.1" so the bridge is loopback-only.

import { test } from "node:test";
import assert from "node:assert/strict";

import { createBridge } from "../daemon/bridge.mjs";
import WebSocket from "ws";
import { once } from "node:events";

// Wait for the ws server to finish binding (it binds asynchronously).
async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening within timeout");
}

test("WS bridge binds to loopback (127.0.0.1), not all interfaces", async () => {
  // port: 0 → ephemeral free port, so the test never collides with a
  // running daemon or another test.
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  try {
    await waitForListening(bridge);
    const addr = bridge.address();
    assert.ok(addr && typeof addr === "object", "address() should return the bound socket info");
    assert.equal(addr.address, "127.0.0.1", "WS bridge must bind loopback, not 0.0.0.0");
  } finally {
    await bridge.close();
  }
});

test("createBridge requires a token", () => {
  assert.throws(() => createBridge({ port: 0 }), /requires a token/);
});

test("diagnostic host counts include authenticated panes only, expose no document paths, and clear on close", async () => {
  let notifyClose;
  const bridge = createBridge({
    port: 0,
    token: "host-count-fixture",
    onClose: () => notifyClose?.(),
  });
  const clients = [];
  try {
    await bridge.ready;
    const connect = async (host, token = "host-count-fixture") => {
      const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
      clients.push(ws);
      await once(ws, "open");
      const response = once(ws, token === "host-count-fixture" ? "message" : "close");
      ws.send(
        JSON.stringify({
          type: "hello",
          token,
          host,
          active_doc: `/secret/${host}-${clients.length}.file`,
        }),
      );
      await response;
      return ws;
    };
    await connect("excel", "wrong-token");
    assert.deepEqual(bridge.diagnosticSummary().connected_panes_by_host, { word: 0, excel: 0 });
    await connect("word");
    await connect("excel");
    const second = await connect("excel");
    assert.deepEqual(bridge.diagnosticSummary().connected_panes_by_host, { word: 1, excel: 2 });
    assert.doesNotMatch(JSON.stringify(bridge.diagnosticSummary()), /secret|file|token/);
    const closed = once(second, "close");
    const serverClosed = new Promise((resolve) => {
      notifyClose = resolve;
    });
    second.close();
    await closed;
    await serverClosed;
    assert.deepEqual(bridge.diagnosticSummary().connected_panes_by_host, { word: 1, excel: 1 });
  } finally {
    clients.forEach((ws) => ws.terminate());
    await bridge.close();
  }
});

test("async admission refuses a released session before an idle SDK waiter receives text", async () => {
  let key, rejectGate, entered;
  const gateEntered = new Promise((resolve) => {
    entered = resolve;
  });
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    onHello: (k) => {
      key = k;
    },
    onUserMessage: () =>
      new Promise((_, reject) => {
        rejectGate = reject;
        entered();
      }),
  });
  let ws;
  try {
    await waitForListening(bridge);
    ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(ws, "open");
    const hello = once(ws, "message");
    ws.send(
      JSON.stringify({
        type: "hello",
        token: "fixture",
        host: "word",
        active_doc: "/fixture.docx",
      }),
    );
    await hello;
    let received = false;
    const idle = bridge.nextUserMessage(key).then(
      () => {
        received = true;
      },
      (e) => e.message,
    );
    ws.send(
      JSON.stringify({ type: "user_message", text: "must not reach SDK", turn_id: "denied" }),
    );
    await gateEntered;
    assert.equal(received, false);
    const terminal = once(ws, "message");
    rejectGate(new Error("Session released for CLI"));
    const result = JSON.parse(String((await terminal)[0]));
    assert.equal(result.event, "turn_complete");
    assert.match(result.error, /released/);
    assert.equal(await idle, "Session aborted");
    assert.equal(received, false);
    assert.equal(bridge.activeTurn(key), null);
  } finally {
    ws?.terminate();
    await bridge.close();
  }
});

test("bridge scopes turns/tools to their document and refuses another pane's tool result", async () => {
  const keys = new Map();
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    runtimeId: "codebuddy",
    onHello: (key, host, doc) => keys.set(doc, key),
  });
  const clients = [];
  async function connect(doc) {
    const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    clients.push(ws);
    await once(ws, "open");
    const received = [];
    ws.on("message", (data) => received.push(JSON.parse(String(data))));
    const send = (data) => ws.send(JSON.stringify(data));
    const wait = async (predicate) => {
      const deadline = Date.now() + 2000;
      while (Date.now() < deadline) {
        const hit = received.find(predicate);
        if (hit) return hit;
        await new Promise((r) => setTimeout(r, 5));
      }
      throw new Error("fixture message timed out");
    };
    send({ type: "hello", token: "fixture", host: "word", active_doc: doc });
    const welcome = await wait((m) => m.type === "welcome");
    assert.equal(welcome.runtime, "codebuddy");
    assert.equal(welcome.protocol_version, 2);
    return { ws, send, wait };
  }
  try {
    await waitForListening(bridge);
    const a = await connect("/matter/a.docx"),
      b = await connect("/matter/b.docx");
    const ka = keys.get("/matter/a.docx"),
      kb = keys.get("/matter/b.docx");
    assert.notEqual(ka, kb);
    a.send({ type: "user_message", text: "a", turn_id: "ta" });
    b.send({ type: "user_message", text: "b", turn_id: "tb" });
    assert.equal((await bridge.nextUserMessage(ka)).text, "a");
    assert.equal((await bridge.nextUserMessage(kb)).text, "b");
    a.send({ type: "user_message", text: "duplicate", turn_id: "bad" });
    await a.wait((m) => m.event === "info");
    assert.equal(bridge.activeTurn(ka), "ta");
    let settled = false;
    const result = bridge.callTaskpaneTool("office_read_paragraphs", {}, ka).then((r) => {
      settled = true;
      return r;
    });
    const call = await a.wait((m) => m.type === "tool_call");
    b.send({ type: "tool_result", id: call.id, ok: true, result: "forged" });
    b.send({ type: "ping" });
    await b.wait((m) => m.type === "pong");
    assert.equal(settled, false);
    a.send({ type: "tool_result", id: call.id, ok: true, result: "owned" });
    assert.equal(await result, "owned");
    bridge.sendAssistantEvent({ event: "turn_complete", turn_id: "old", subtype: "success" }, ka);
    assert.equal(bridge.activeTurn(ka), "ta");
    bridge.sendAssistantEvent({ event: "turn_complete", turn_id: "ta", subtype: "success" }, ka);
    assert.equal(bridge.activeTurn(ka), null);
    assert.equal(bridge.activeTurn(kb), "tb");
    const closed = once(a.ws, "close");
    a.send({ type: "context_update", active_doc: "/matter/saved-as.docx" });
    assert.equal((await closed)[0], 4006);
  } finally {
    clients.forEach((ws) => ws.terminate());
    await bridge.close();
  }
});
