// Unit tests for daemon/runtime/turns.mjs — the shared per-turn message
// preparation extracted from daemon/index.mjs during the runtime-adapter
// refactor (behavior must match the old userMessageStream +
// renderContextHeader exactly).

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  renderContextHeader,
  prepareTurnContent,
  createTurnSource,
} from "../daemon/runtime/turns.mjs";

test("renderContextHeader: empty when nothing to say; combines host/doc/selection", () => {
  assert.equal(renderContextHeader({}), "");
  assert.equal(
    renderContextHeader({ host: "word", activeDoc: "/x/a.docx" }),
    "[Host: Word · Doc: /x/a.docx]",
  );
  assert.equal(
    renderContextHeader({ host: "excel", selection: { text: "hello" } }),
    '[Host: Excel · Selection: "hello"]',
  );
});

test("renderContextHeader: selection text is truncated to 80 chars; cursor-only uses para id", () => {
  const long = "x".repeat(120);
  const header = renderContextHeader({ host: "word", selection: { text: long } });
  assert.equal(header, `[Host: Word · Selection: "${"x".repeat(77)}..."]`);
  assert.equal(
    renderContextHeader({ host: "word", selection: { para_id: "0000002A" } }),
    "[Host: Word · Cursor in paragraph 0000002A]",
  );
});

test("renderContextHeader: track-changes surfaces only non-default modes", () => {
  assert.equal(renderContextHeader({ host: "word", trackChangesMode: "always" }), "[Host: Word]");
  assert.equal(
    renderContextHeader({ host: "word", trackChangesMode: "off" }),
    "[Host: Word · Track changes: off]",
  );
  assert.equal(
    renderContextHeader({ host: "word", trackChangesMode: "never" }),
    "[Host: Word · Track changes: do not turn tracking on; existing document tracking still records edits; word and phrase replacements are still tracked]",
  );
});

test("prepareTurnContent: slash commands stay bare; others get the header; flags reset", () => {
  const session = {
    slashCommandPending: false,
    turnProducedOutput: true,
    interrupted: true,
    streamedText: { id: "msg-1", text: "half a sentence" },
    retryNotified: true,
  };

  const slash = prepareTurnContent({ text: "  /help", context: { host: "word" } }, session);
  assert.equal(slash, "/help");
  assert.equal(session.slashCommandPending, true);
  assert.equal(session.turnProducedOutput, false);
  // A Stop that landed while idle must not leak into this turn.
  assert.equal(session.interrupted, false);
  // Per-turn message-layer tracking resets.
  assert.equal(session.streamedText, null);
  assert.equal(session.retryNotified, false);

  const normal = prepareTurnContent(
    { text: "summarize", context: { host: "word", activeDoc: "/x/a.docx" } },
    session,
  );
  assert.equal(normal, "[Host: Word · Doc: /x/a.docx]\n\nsummarize");
  assert.equal(session.slashCommandPending, false);

  // No session (defensive) — content still builds.
  assert.equal(prepareTurnContent({ text: "hi", context: {} }, null), "hi");
});

test("createTurnSource: a rejected bridge waiter yields null (clean shutdown)", async () => {
  let rejectNow;
  const pull = () =>
    new Promise((_, reject) => {
      rejectNow = reject;
    });
  const src = createTurnSource({ pull, session: null });
  const pending = src.next();
  rejectNow(new Error("Session aborted"));
  assert.equal(await pending, null);
});

test("createTurnSource: resolves prepared content for a normal message", async () => {
  const pull = async () => ({ text: "do it", context: { host: "excel" } });
  const src = createTurnSource({ pull, session: null });
  assert.equal(await src.next(), "[Host: Excel]\n\ndo it");
});
