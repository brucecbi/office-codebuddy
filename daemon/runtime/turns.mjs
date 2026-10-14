// Shared per-turn plumbing for every runtime.
//
// A "turn" is one user message with its context header prepended. The
// header ([Host: Word · Doc: … · Selection: …]) tells the agent which
// surface and document it is operating on without an extra tool call.
//
// Slash commands must stay bare: the SDKs only treat a message as a
// command when it STARTS with "/", and the header would push it off the
// front. The command template is self-contained; it can call office_*
// tools if it needs doc context.
//
// Runtime loops pull turns through createTurnSource(); a null return
// means the bridge waiter was rejected (that pane's session was
// swapped/aborted), which the loop treats as a clean shutdown — exactly
// the old infinite-stdin generator's exit path.

export function renderContextHeader(ctx) {
  const parts = [];
  // Host first so the agent immediately knows which tool family to use.
  // Both office_* (Word) and excel_* tools are registered simultaneously;
  // without this hint, the agent can pick the wrong family.
  if (ctx.host === "word" || ctx.host === "excel") {
    parts.push(`Host: ${ctx.host === "word" ? "Word" : "Excel"}`);
  }
  if (ctx.activeDoc) parts.push(`Doc: ${ctx.activeDoc}`);
  if (ctx.selection) {
    const s = ctx.selection;
    if (s.text) {
      const preview = s.text.length > 80 ? s.text.slice(0, 77) + "..." : s.text;
      parts.push(`Selection: "${preview}"`);
    } else if (s.para_id) {
      parts.push(`Cursor in paragraph ${s.para_id}`);
    }
  }
  // Only surface track-changes mode when it's NOT the default ("always"). The
  // system prompt says "always" by default; only deviation needs signaling.
  // "never" means do not enable tracking. Spell that out: the bare word reads
  // as "edits will be untracked", which is false when Word already has it on.
  if (ctx.trackChangesMode === "never") {
    parts.push(
      "Track changes: do not turn tracking on; existing document tracking still records edits; word and phrase replacements are still tracked",
    );
  } else if (ctx.trackChangesMode && ctx.trackChangesMode !== "always") {
    parts.push(`Track changes: ${ctx.trackChangesMode}`);
  }
  return parts.length ? `[${parts.join(" · ")}]` : "";
}

// Build the content string for one user message and reset the session's
// per-turn tracking flags (used to detect slash commands that produce no
// chat output — terminal-only built-ins like /help).
export function prepareTurnContent(msg, session) {
  const { text, context } = msg;
  const trimmed = typeof text === "string" ? text.trimStart() : text;
  const isSlashCommand = typeof trimmed === "string" && trimmed.startsWith("/");
  const header = renderContextHeader(context);
  const content = isSlashCommand ? trimmed : header ? `${header}\n\n${text}` : text;
  if (session) {
    session.slashCommandPending = isSlashCommand;
    session.turnProducedOutput = false;
    // Per-turn message-layer tracking (phase 3):
    //  - streamedText: dedup bookkeeping for the complete-assistant-message
    //    fallback ({ id, text } for the latest streamed message).
    //  - retryNotified: one "retrying" pane event per turn, not per attempt.
    session.streamedText = null;
    session.retryNotified = false;
    // A Stop that landed while no turn was in flight must not leak into
    // this turn's terminal state: on the CodeBuddy transport interrupt()
    // on an idle session is a no-op, so the flag would otherwise survive
    // until the next result and mislabel a completed answer as cancelled.
    session.interrupted = false;
  }
  return content;
}

// Pull-based turn source over the bridge's per-pane queue.
export function createTurnSource({ pull, session, onTurnStart = () => {} }) {
  return {
    async next() {
      let msg;
      try {
        msg = await pull();
      } catch {
        // Bridge rejected the waiter — session was aborted. Exit cleanly.
        return null;
      }
      if (session?.lifecycle?.closed) return null;
      const content = prepareTurnContent(msg, session);
      onTurnStart(msg);
      return content;
    },
  };
}
