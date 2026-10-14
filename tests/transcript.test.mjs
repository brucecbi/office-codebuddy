// Unit tests for daemon/transcript.mjs — the SDK-session .jsonl → chat
// bubble reconstruction. Pure transform + a tmpdir round-trip; no SDK,
// no Office.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { stripContextHeader, eventsFromLine, readTranscript } from "../daemon/transcript.mjs";

test("stripContextHeader removes the injected per-turn header", () => {
  assert.equal(
    stripContextHeader('[Host: Word · Doc: /x/y.docx · Selection: "hi"]\n\nFix the typo.'),
    "Fix the typo.",
  );
  assert.equal(stripContextHeader("No header here."), "No header here.");
  // Only the leading bracketed line is stripped, not later bracketed text.
  assert.equal(stripContextHeader("[Doc: a]\n\nsee [Note: keep this]"), "see [Note: keep this]");
});

test("eventsFromLine: user string → user bubble (header stripped)", () => {
  const evs = eventsFromLine({
    type: "user",
    message: { role: "user", content: "[Doc: a]\n\nhello" },
  });
  assert.deepEqual(evs, [{ kind: "user", text: "hello" }]);
});

test("eventsFromLine: user array keeps tool result identity and status", () => {
  const evs = eventsFromLine({
    type: "user",
    message: {
      role: "user",
      content: [{ type: "tool_result", tool_use_id: "tool-1", content: "ok" }],
    },
  });
  assert.deepEqual(evs, [{ kind: "tool_result", id: "tool-1", output: "ok", status: "completed" }]);
});

test("eventsFromLine: assistant text + tool_use, thinking skipped", () => {
  const evs = eventsFromLine({
    type: "assistant",
    message: {
      role: "assistant",
      content: [
        { type: "thinking", thinking: "hmm" },
        { type: "text", text: "Done." },
        { type: "tool_use", name: "mcp__office__office_replace_text", input: { a: 1 } },
      ],
    },
  });
  assert.deepEqual(evs, [
    { kind: "assistant", text: "Done." },
    { kind: "tool", name: "mcp__office__office_replace_text", input: { a: 1 } },
  ]);
});

test("eventsFromLine: internal lines (queue-operation/ai-title) → skipped", () => {
  assert.deepEqual(eventsFromLine({ type: "queue-operation", foo: 1 }), []);
  assert.deepEqual(eventsFromLine({ type: "ai-title", title: "x" }), []);
  assert.deepEqual(eventsFromLine({}), []);
});

// CodeBuddy 2.149.0 persists native history items, not its Claude-compatible
// stream-json envelopes. A nested message may contain usage only.
test("CodeBuddy native messages retain top-level content when nested usage is present", () => {
  assert.deepEqual(
    eventsFromLine(
      {
        type: "message",
        role: "assistant",
        content: [{ type: "output_text", text: "Synthetic summary." }],
        message: { usage: { output_tokens: 12 } },
      },
      "codebuddy",
    ),
    [{ kind: "assistant", text: "Synthetic summary." }],
  );
  assert.deepEqual(
    eventsFromLine(
      {
        type: "message",
        role: "user",
        content: [{ type: "input_text", text: "Read the synthetic document." }],
      },
      "codebuddy",
    ),
    [{ kind: "user", text: "Read the synthetic document." }],
  );
});

test("CodeBuddy native function calls/results retain call identity, arguments and full output", () => {
  const call = {
    type: "function_call",
    id: "history-item-1",
    callId: "office-call-1",
    name: "mcp__office__office_read_story",
    arguments: '{"target":{"story":"body"}}',
    message: { usage: { output_tokens: 8 } },
  };
  assert.deepEqual(eventsFromLine(call, "codebuddy"), [
    {
      kind: "tool",
      id: "office-call-1",
      name: "mcp__office__office_read_story",
      input: { target: { story: "body" } },
    },
  ]);
  const output = { type: "text", text: "Synthetic complete document contents." };
  assert.deepEqual(
    eventsFromLine(
      {
        type: "function_call_result",
        id: "history-item-2",
        callId: "office-call-1",
        status: "completed",
        output,
      },
      "codebuddy",
    ),
    [{ kind: "tool_result", id: "office-call-1", output, status: "completed" }],
  );
  assert.deepEqual(
    eventsFromLine(
      { type: "function_call_output", call_id: "office-call-1", output: [output] },
      "codebuddy",
    ),
    [{ kind: "tool_result", id: "office-call-1", output: [output], status: "completed" }],
  );
  assert.deepEqual(eventsFromLine(call, "claude"), []);
});

test("CodeBuddy native tool failures never replay as success", () => {
  const base = { type: "function_call_result", callId: "failed-call", output: "Refused" };
  for (const failure of [
    { status: "incomplete" },
    { status: "failed" },
    { status: "cancelled" },
    { is_error: true },
    { status: "completed", providerData: { error: "denied" } },
    { status: "completed", providerData: { toolResult: { error: "denied" } } },
  ]) {
    assert.deepEqual(eventsFromLine({ ...base, ...failure }, "codebuddy"), [
      { kind: "tool_result", id: "failed-call", output: "Refused", status: "error" },
    ]);
  }
  assert.equal(
    eventsFromLine({ ...base, status: "in_progress" }, "codebuddy")[0].status,
    "in_progress",
  );
});

test("CodeBuddy malformed call arguments remain visible and legacy envelopes still work", () => {
  assert.deepEqual(
    eventsFromLine(
      { type: "function_call", call_id: "call", name: "Read", arguments: '{"unfinished":' },
      "codebuddy",
    ),
    [{ kind: "tool", id: "call", name: "Read", input: '{"unfinished":' }],
  );
  assert.deepEqual(
    eventsFromLine(
      { type: "assistant", message: { role: "assistant", content: "Older envelope" } },
      "codebuddy",
    ),
    [{ kind: "assistant", text: "Older envelope" }],
  );
});

test("CodeBuddy completed Office result preserves nested native approval failure", () => {
  const output = {
    type: "text",
    text: JSON.stringify([
      { type: "text", text: "Error: USER_DECLINED_OR_EXPIRED: no edit was authorized." },
    ]),
  };
  const line = {
    type: "function_call_result",
    name: "mcp__office__office_insert_paragraphs",
    callId: "native-expired-approval",
    status: "completed",
    output,
    providerData: { toolResult: { content: output.text } },
  };
  assert.deepEqual(eventsFromLine(line, "codebuddy"), [
    { kind: "tool_result", id: line.callId, output, status: "error" },
  ]);
  const { output: ignored, ...fallback } = line;
  assert.deepEqual(eventsFromLine(fallback, "codebuddy"), [
    { kind: "tool_result", id: line.callId, output: output.text, status: "error" },
  ]);
});

test("CodeBuddy native uncertain Office write does not replay as successful", () => {
  const output = JSON.stringify([
    { type: "text", text: "Error: RESULT_UNKNOWN: CANCELLED_OR_EXPIRED: inspect the document." },
  ]);
  const event = eventsFromLine(
    {
      type: "function_call_output",
      name: "office_replace_section",
      call_id: "native-partial-write",
      status: "completed",
      output,
    },
    "codebuddy",
  )[0];
  assert.equal(event.status, "error");
  assert.equal(event.output, output);
});

test("CodeBuddy successful readback with Error text in document data remains successful", () => {
  for (const payload of [
    { paragraphs: [{ text: "Error: USER_DECLINED_OR_EXPIRED: this is document text." }] },
    { text: "Error: an example quoted by a document", coverage: { authoritative: true } },
    { inserted_count: 1, readback: { text: "No error. Error: a quoted label." } },
  ]) {
    const output = {
      type: "text",
      text: JSON.stringify([{ type: "text", text: JSON.stringify(payload) }]),
    };
    assert.deepEqual(
      eventsFromLine(
        {
          type: "function_call_result",
          name: "mcp__office__office_read_paragraphs",
          callId: "success",
          status: "completed",
          output,
        },
        "codebuddy",
      ),
      [{ kind: "tool_result", id: "success", output, status: "completed" }],
    );
  }
});

test("readTranscript returns [] for a missing session", async () => {
  const r = await readTranscript("definitely-not-a-real-session-id-xyz");
  assert.deepEqual(r, { events: [], truncated: false });
  assert.deepEqual(await readTranscript(null), { events: [], truncated: false });
});

test("readTranscript round-trips a real .jsonl under ~/.claude/projects + truncates", async () => {
  // transcript.mjs reads $HOME/.claude/projects; override HOME to a tmpdir.
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-transcript-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    // transcript.mjs caches PROJECTS_DIR at import time from the original
    // HOME, so re-import a fresh copy bound to the fake HOME.
    const mod = await import("../daemon/transcript.mjs?t=" + Date.now() + Math.random());
    const projDir = join(fakeHome, ".claude", "projects", "proj-1");
    await mkdir(projDir, { recursive: true });
    const sid = "session-abc";
    const lines = [
      JSON.stringify({ type: "queue-operation", x: 1 }),
      JSON.stringify({ type: "user", message: { role: "user", content: "[Doc: a]\n\nfirst" } }),
      JSON.stringify({
        type: "assistant",
        message: { role: "assistant", content: [{ type: "text", text: "reply one" }] },
      }),
      "not json — should be skipped",
      JSON.stringify({ type: "user", message: { role: "user", content: "second" } }),
      JSON.stringify({
        type: "assistant",
        message: { role: "assistant", content: [{ type: "tool_use", name: "Bash", input: {} }] },
      }),
    ];
    await writeFile(join(projDir, `${sid}.jsonl`), lines.join("\n") + "\n");

    const full = await mod.readTranscript(sid);
    assert.equal(full.truncated, false);
    assert.deepEqual(full.events, [
      { kind: "user", text: "first" },
      { kind: "assistant", text: "reply one" },
      { kind: "user", text: "second" },
      { kind: "tool", name: "Bash", input: {} },
    ]);

    // maxEvents ring buffer: keep only the last 2, flag truncated.
    const clipped = await mod.readTranscript(sid, { maxEvents: 2 });
    assert.equal(clipped.truncated, true);
    assert.equal(clipped.events.length, 2);
    assert.deepEqual(
      clipped.events.map((e) => e.kind),
      ["user", "tool"],
    );
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
});
