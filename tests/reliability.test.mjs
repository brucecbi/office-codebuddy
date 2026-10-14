import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import { createSessionStore } from "../daemon/sessions.mjs";
import { createLifecycle } from "../daemon/runtime/lifecycle.mjs";
import { createMessageRenderer } from "../daemon/runtime/render-messages.mjs";
import { runSession as codebuddy } from "../daemon/runtime/runtime-codebuddy.mjs";
import { runSession as claude } from "../daemon/runtime/runtime-claude.mjs";
import { eventsFromLine } from "../daemon/transcript.mjs";
// All transports in this file are injected synthetic mocks; no authentication occurs.
process.env.CODEBUDDY_INTERNET_ENVIRONMENT = "external";

test("independent stores serialize concurrent writes and scope documents/providers/hosts", async () => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-v4-"));
  try {
    const file = join(dir, "sessions.json");
    const stores = [0, 1].map(() => createSessionStore({ file, legacyFile: null }));
    await Promise.all(
      Array.from({ length: 40 }, (_, i) =>
        stores[i % 2].saveSessionId(
          i % 3 ? "word" : "excel",
          "/matter",
          "sid-" + i,
          i % 2 ? "claude" : "codebuddy",
          "doc-" + i,
        ),
      ),
    );
    for (let i = 0; i < 40; i++)
      assert.equal(
        await stores[0].getSessionId(
          i % 3 ? "word" : "excel",
          "/matter",
          i % 2 ? "claude" : "codebuddy",
          "doc-" + i,
        ),
        "sid-" + i,
      );
    assert.equal(await stores[0].getSessionId("word", "/other", "codebuddy", "doc-2"), null);
    assert.equal(await stores[0].getSessionId("word", "/matter", "claude", "doc-2"), null);
    assert.equal(await stores[0].getSessionId("excel", "/matter", "codebuddy", "doc-2"), null);
  } finally {
    await rm(dir, { recursive: true });
  }
});

test("legacy v3 is preserved byte-for-byte, backed up, and not auto-assigned to a document", async () => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-migrate-"));
  try {
    const legacyFile = join(dir, "old.json"),
      file = join(dir, "new.json");
    const raw = JSON.stringify({
      version: 3,
      folders: {
        "/matter": { sessions: { claude: { word: "old-c" }, codebuddy: { word: "old-b" } } },
      },
    });
    await writeFile(legacyFile, raw);
    const s = createSessionStore({ file, legacyFile });
    assert.equal(await s.getSessionId("word", "/matter", "codebuddy", "new-doc"), null);
    await s.touchFolder("/matter");
    assert.equal(await readFile(legacyFile, "utf8"), raw);
    const backup = (await readdir(dir)).find((n) => n.endsWith(".bak"));
    assert.equal(await readFile(join(dir, backup), "utf8"), raw);
    assert.equal(await s.getSessionId("word", "/matter", "codebuddy"), "old-b");
    assert.equal(JSON.parse(await readFile(file, "utf8")).version, 4);
    const future = '{"version":99,"folders":{}}';
    await writeFile(file, future);
    await assert.rejects(s.touchFolder("/matter"), /Unsupported/);
    assert.equal(await readFile(file, "utf8"), future);
    await writeFile(file, "{bad");
    await assert.rejects(s.touchFolder("/matter"));
    assert.equal(await readFile(file, "utf8"), "{bad");
  } finally {
    await rm(dir, { recursive: true });
  }
});

test("Stop before transport attach is latched; terminal and cleanup occur once", () => {
  const events = [];
  let closes = 0,
    cleanups = 0;
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup: () => cleanups++ });
  s.begin("first");
  s.cancel();
  s.attach({ close: () => closes++ });
  s.fail("late error");
  s.cancel();
  assert.equal(events.length, 1);
  assert.equal(events[0].interrupted, true);
  assert.equal(events[0].turn_id, "first");
  assert.equal(closes, 1);
  assert.equal(cleanups, 1);
  assert.equal(s.begin("next"), false);
});

test("a successful earlier turn does not hide a later stream-end failure", () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {} });
  s.begin("one");
  s.finish({ subtype: "success" });
  s.begin("two");
  s.fail("stream ended", "stream_ended");
  assert.deepEqual(
    events.map((e) => [e.turn_id, e.subtype]),
    [
      ["one", "success"],
      ["two", "stream_ended"],
    ],
  );
});

test("a hung startup/turn has a bounded terminal and cleanup", async () => {
  const events = [];
  const s = createLifecycle({ emit: (e) => events.push(e), cleanup() {}, timeoutMs: 10 });
  s.begin("hung");
  await delay(25);
  assert.equal(events.length, 1);
  assert.equal(events[0].subtype, "timeout");
  assert.equal(s.closed, true);
});

test("renderer reconciles standard message_start and divergent multi-block deltas", () => {
  const events = [];
  const render = createMessageRenderer((e) => events.push(e), "turn");
  const stream = (event) => render({ type: "stream_event", event });
  stream({ type: "message_start", message: { id: "m1" } });
  stream({ type: "content_block_delta", index: 0, delta: { type: "text_delta", text: "ab" } });
  stream({
    type: "content_block_delta",
    index: 1,
    delta: { type: "thinking_delta", thinking: "PRIVATE" },
  });
  stream({ type: "content_block_delta", index: 2, delta: { type: "text_delta", text: "XYZ" } });
  const complete = {
    type: "assistant",
    message: {
      id: "m1",
      content: [
        { type: "text", text: "abc" },
        { type: "thinking", thinking: "PRIVATE" },
        { type: "text", text: "XYZ" },
      ],
    },
  };
  render(complete);
  const n = events.length;
  render(complete);
  assert.equal(events.length, n);
  assert.equal(events.at(-1).text, "abcXYZ");
  assert.equal(new Set(events.map((e) => e.message_id)).size, 1);
  assert.ok(!JSON.stringify(events).includes("PRIVATE"));
});

test("complete-only output, unknown-id deltas, and tool result IDs remain visible", () => {
  const events = [];
  const r = createMessageRenderer((e) => events.push(e), "t");
  r({ type: "stream_event", event: { delta: { type: "text_delta", text: "Hello" } } });
  assert.equal(events.length, 0);
  r({
    type: "assistant",
    message: {
      id: "m",
      content: [
        { type: "text", text: "Hello" },
        { type: "tool_use", id: "u", name: "Bash", input: { command: "pwd" } },
      ],
    },
  });
  r({
    type: "user",
    message: { content: [{ type: "tool_result", tool_use_id: "u", content: "ok" }] },
  });
  assert.equal(events.find((e) => e.event === "assistant_message").text, "Hello");
  assert.equal(events.at(-1).tool_id, "u");
  assert.equal(events.at(-1).status, "completed");
});

test("CodeBuddy connect/resume failure closes resources and does not silently discard history", async () => {
  let closes = 0,
    fresh = 0;
  const endings = [];
  const h = codebuddy(
    {
      cwd: "/tmp",
      resumeId: "locked",
      turnSource: { next: async () => null },
      onMessage() {},
      onEnd: (e) => endings.push(e),
    },
    {
      create() {
        fresh++;
      },
      resume() {
        return {
          connect: async () => {
            throw new Error("Session locked");
          },
          close() {
            closes++;
          },
        };
      },
      observeClose: (session) => ({ close: () => session.close(), waitForExit: async () => {} }),
    },
  );
  await h.done;
  assert.equal(closes, 1);
  assert.equal(fresh, 0);
  assert.equal(endings[0].reason, "error");
});

test("CodeBuddy Stop interrupts a hung connect and releases the lock exactly once", async () => {
  let closes = 0;
  const events = [];
  const h = codebuddy(
    {
      cwd: "/tmp",
      turnSource: { next: async () => null },
      onMessage() {},
      onEnd: (e) => events.push(e),
    },
    {
      create: () => ({
        connect: () => new Promise(() => {}),
        close() {
          closes++;
        },
        interrupt: async () => {},
      }),
      observeClose: (session) => ({ close: () => session.close(), waitForExit: async () => {} }),
    },
  );
  h.cancel();
  await h.done;
  h.close();
  assert.equal(closes, 1);
  assert.deepEqual(events, [{ reason: "aborted" }]);
});

test("CodeBuddy processes two turns, passes disallowedTools, and reports resultless stream once", async () => {
  let sends = 0,
    streamReturns = 0,
    closes = 0,
    options;
  const end = [];
  const notes = [];
  const h = codebuddy(
    {
      cwd: "/tmp",
      model: "chosen",
      disallowedTools: ["foreign_word_tool"],
      turnSource: { next: async () => "next" },
      onMessage: (m) => notes.push(m.type),
      onTurnEnd: (e) => notes.push(e.sawResult),
      onEnd: (e) => end.push(e),
    },
    {
      create(o) {
        options = o;
        return {
          connect: async () => {},
          send: async () => sends++,
          async *stream() {
            try {
              if (sends === 1) yield { type: "result", subtype: "success" };
            } finally {
              streamReturns++;
            }
          },
          close() {
            closes++;
          },
        };
      },
      observeClose: (session) => ({ close: () => session.close(), waitForExit: async () => {} }),
    },
  );
  await h.done;
  assert.equal(options.model, "chosen");
  assert.deepEqual(options.disallowedTools, ["foreign_word_tool"]);
  assert.equal(sends, 2);
  assert.equal(streamReturns, 2);
  assert.equal(closes, 1);
  assert.deepEqual(notes, ["result", false]);
  assert.equal(end.length, 1);
});

test("Claude transport tears down its iterator on an exception", async () => {
  let closes = 0;
  const endings = [];
  const h = claude(
    { turnSource: { next: async () => null }, onMessage() {}, onEnd: (e) => endings.push(e) },
    {
      querySession: () => ({
        async *[Symbol.asyncIterator]() {
          throw new Error("offline failure");
        },
        close() {
          closes++;
        },
      }),
    },
  );
  await h.done;
  assert.equal(closes, 1);
  assert.equal(endings[0].reason, "error");
});

test("CodeBuddy top-level history schema supports input_text/output_text and omits thinking", () => {
  assert.deepEqual(
    eventsFromLine(
      { role: "user", content: [{ type: "input_text", text: "question" }] },
      "codebuddy",
    ),
    [{ kind: "user", text: "question" }],
  );
  assert.deepEqual(
    eventsFromLine(
      {
        role: "assistant",
        content: [
          { type: "thinking", thinking: "hidden" },
          { type: "output_text", text: "answer" },
        ],
      },
      "codebuddy",
    ),
    [{ kind: "assistant", text: "answer" }],
  );
  assert.deepEqual(eventsFromLine({ role: "assistant", content: "answer" }, "codebuddy"), [
    { kind: "assistant", text: "answer" },
  ]);
});
