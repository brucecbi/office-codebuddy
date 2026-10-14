// Loopback integration tests for the bridge's write-tracking and activity
// surface used by the activity-based turn timeout:
//  - hasPendingWrites(key) distinguishes read-only from mutating tool calls
//  - onActivity fires on tool dispatch/result (keeps busy turns alive)
//  - tool calls are logged with their name and elapsed time (no args/results)

import { test } from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import WebSocket from "ws";

import { createBridge } from "../daemon/bridge.mjs";
import { timeoutCopy } from "../daemon/runtime/lifecycle.mjs";

async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening within timeout");
}

async function fakePane(bridge, { doc = "/tmp/fixture.docx" } = {}) {
  await waitForListening(bridge);
  const ws = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
  await once(ws, "open");
  const welcome = once(ws, "message");
  ws.send(
    JSON.stringify({
      type: "hello",
      token: "fixture",
      host: "word",
      active_doc: doc,
      pane_id: "pane-1",
      protocol_version: 2,
    }),
  );
  ws.welcome = JSON.parse(String((await welcome)[0]));
  return ws;
}

function toolResult(ws, id, result = { ok: true }) {
  ws.send(JSON.stringify({ type: "tool_result", id, ok: true, result }));
}

for (const outcome of ["verified_write", "RESULT_UNKNOWN"]) {
  test(`timeout after settled ${outcome} never claims the document is unchanged`, async (t) => {
    let key;
    const bridge = createBridge({
      port: 0,
      token: "fixture",
      onHello: (k) => {
        key = k;
      },
    });
    const ws = await fakePane(bridge);
    t.after(async () => {
      ws.terminate();
      await bridge.close();
    });
    ws.on("message", (raw) => {
      const msg = JSON.parse(raw);
      if (msg.type === "tool_call")
        ws.send(
          JSON.stringify({
            type: "tool_result",
            id: msg.id,
            ok: outcome === "verified_write",
            result: { applied: true, verified: true },
            error:
              outcome === "RESULT_UNKNOWN"
                ? "RESULT_UNKNOWN: post-write verification failed"
                : undefined,
          }),
        );
    });
    const call = bridge.callTaskpaneTool("office_replace_text", {}, key);
    if (outcome === "RESULT_UNKNOWN") await assert.rejects(call, /RESULT_UNKNOWN/);
    else await call;
    assert.equal(bridge.hasPendingWrites(key), false);
    const text = timeoutCopy("idle", bridge.hasPendingWrites(key), {
      idleMs: 120000,
      maxMs: 900000,
    });
    assert.doesNotMatch(text, /document is unchanged|ask the agent to continue/);
    assert.match(text, /earlier tool results/i);
    assert.match(text, /RESULT_UNKNOWN/);
  });
}

test("hasPendingWrites tracks mutating vs read-only tool calls per pane", async (t) => {
  let key = null;
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    allowedOrigins: [],
    onHello: (k) => {
      key = k;
    },
  });
  const ws = await fakePane(bridge);
  t.after(async () => {
    ws.close();
    await bridge.close();
  });
  // Answer every tool_call as soon as it arrives, before any dispatch below.
  ws.on("message", (data) => {
    const msg = JSON.parse(data);
    if (msg.type === "tool_call") toolResult(ws, msg.id);
  });

  assert.equal(bridge.hasPendingWrites(key), false);

  const read = bridge.callTaskpaneTool("office_read_paragraphs", {}, key);
  read.catch(() => {});
  assert.equal(bridge.hasPendingWrites(key), false, "read-only tools are not pending writes");

  const write = bridge.callTaskpaneTool("office_replace_text", { find: "a" }, key);
  write.catch(() => {});
  assert.equal(bridge.hasPendingWrites(key), true, "mutating tools count as pending writes");

  await Promise.allSettled([read, write]);
  assert.equal(bridge.hasPendingWrites(key), false, "settled calls clear the flag");
});

test("onActivity fires on tool dispatch and on tool_result", async (t) => {
  let key = null;
  const activity = [];
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    allowedOrigins: [],
    onHello: (k) => {
      key = k;
    },
    onActivity: (k) => activity.push(k),
  });
  const ws = await fakePane(bridge);
  t.after(async () => {
    ws.close();
    await bridge.close();
  });
  ws.on("message", (data) => {
    const msg = JSON.parse(data);
    if (msg.type === "tool_call") toolResult(ws, msg.id);
  });

  await bridge.callTaskpaneTool("office_get_outline", {}, key);
  assert.ok(activity.length >= 2, `expected dispatch+result activity, got ${activity.length}`);
  assert.ok(activity.every((k) => k === key));
});

test("tool calls are logged with name and elapsed time, never args", async (t) => {
  let key = null;
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    allowedOrigins: [],
    onHello: (k) => {
      key = k;
    },
  });
  const ws = await fakePane(bridge);
  t.after(async () => {
    ws.close();
    await bridge.close();
  });
  ws.on("message", (data) => {
    const msg = JSON.parse(data);
    if (msg.type === "tool_call") toolResult(ws, msg.id);
  });

  const lines = [];
  const origLog = console.log;
  console.log = (...args) => {
    lines.push(args.join(" "));
    origLog(...args);
  };
  try {
    await bridge.callTaskpaneTool("office_replace_text", { find: "SECRET-ARG" }, key);
  } finally {
    console.log = origLog;
  }
  const callLine = lines.find((l) => l.includes("tool_call office_replace_text"));
  const resultLine = lines.find((l) => l.includes("tool_result office_replace_text"));
  assert.ok(callLine, "dispatch must be logged");
  assert.ok(resultLine, "settlement must be logged");
  assert.match(resultLine, /elapsed=\d+ms/);
  assert.ok(!lines.some((l) => l.includes("SECRET-ARG")), "tool args must not be logged");
});

test("a confirmation card re-arms the tool deadline; unconfirmed calls still expire", async (t) => {
  let key;
  const bridge = createBridge({
    port: 0,
    token: "fixture",
    toolTimeoutMs: 60,
    onHello: (k) => {
      key = k;
    },
  });
  const ws = await fakePane(bridge);
  t.after(async () => {
    ws.terminate();
    await bridge.close();
  });
  assert.ok(ws.welcome.capabilities.includes("confirm_window_v1"));
  const calls = [];
  ws.on("message", (raw) => {
    const msg = JSON.parse(raw);
    if (msg.type === "tool_call") calls.push(msg);
  });

  const expired = bridge.callTaskpaneTool("office_replace_text", {}, key);
  await assert.rejects(expired, /timed out/);

  const confirmed = bridge.callTaskpaneTool("office_replace_text", {}, key);
  while (calls.length < 2) await new Promise((r) => setTimeout(r, 5));
  const { id, turn_id } = calls[1];
  ws.send(JSON.stringify({ type: "tool_phase", id, turn_id, phase: "awaiting_confirmation" }));
  await new Promise((r) => setTimeout(r, 200));
  toolResult(ws, id, { applied: true });
  assert.deepEqual(await confirmed, { applied: true });
});
