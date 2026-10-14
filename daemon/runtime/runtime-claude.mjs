// Claude runtime — the original Draftspect transport, wrapping
// @anthropic-ai/claude-agent-sdk's query(). Select with DRAFTSPECT_RUNTIME=claude
// only after reviewing provider authentication and usage terms.
// CodeBuddy is the default. This transport's behavior must stay identical to
// the inline loop it replaced:
//   - one long-lived query() fed by an INFINITE prompt stream
//     (the old userMessageStream), aborted via AbortController
//   - systemPrompt { type:"preset", preset:"claude_code", append }
//   - stderr sniffing (onStderr) for usage-limit hints
//   - resume via options.resume

import { query, createSdkMcpServer, tool } from "@anthropic-ai/claude-agent-sdk";
import { previewSdkOptions } from "../preview-policy.mjs";

export const id = "claude";
export const supportsStderr = true;

// MCP registration primitives for the office tools. MUST come from the
// same SDK as the transport — handing a session a server instance built by
// the other SDK makes the provider reject the request (400).
export const mcp = { createSdkMcpServer, tool };

// Standard aliases for the retained experimental adapter. Account availability
// is provider-controlled; this list is not an entitlement check.
export const STATIC_MODELS = ["haiku", "sonnet", "opus"];

// Synchronous catalog for validation (null = discovery-based runtime).
export const knownModels = STATIC_MODELS;

export async function probe() {
  try {
    await import("@anthropic-ai/claude-agent-sdk");
    return { ok: true, detail: "claude-agent-sdk importable" };
  } catch (err) {
    return { ok: false, detail: String(err?.message ?? err) };
  }
}

// The fixed catalog, in discovery shape (Phase 4: the daemon serves this
// to the taskpane's model dropdown via list_models).
export async function listModels() {
  return STATIC_MODELS.map((value) => ({ value }));
}

// Start one pane's agent loop. Returns a handle with cancel()/close();
// the loop reports terminal states through onEnd — it never throws to
// the caller.
export function runSession(
  {
    cwd,
    resumeId = null,
    model,
    systemPromptAppend,
    mcpServers,
    disallowedTools,
    canUseTool,
    includePartialMessages = true,
    turnSource,
    isStale = () => false,
    onMessage,
    onStderr,
    onEnd,
  },
  { querySession = query } = {},
) {
  const abortController = new AbortController();
  let connection;
  let closed = false;
  const close = () => {
    if (closed) return;
    closed = true;
    abortController.abort();
    try {
      connection?.close();
    } catch {}
  };

  // Infinite input stream — the shape query() requires. Ends when the
  // turn source reports the pane's queue was cleared (session swapped).
  const prompt = (async function* () {
    while (true) {
      const content = await turnSource.next();
      if (content === null) return;
      yield {
        type: "user",
        message: { role: "user", content },
        parent_tool_use_id: null,
      };
    }
  })();

  const makeOptions = () =>
    previewSdkOptions({
      cwd,
      systemPrompt: { type: "preset", preset: "claude_code", append: systemPromptAppend },
      mcpServers,
      disallowedTools,
      canUseTool,
      includePartialMessages,
      model,
      abortController,
      stderr: onStderr,
    });

  const done = (async () => {
    try {
      const options = makeOptions();
      if (resumeId) options.resume = resumeId;
      connection = querySession({ prompt, options });
      for await (const msg of connection) {
        if (isStale()) break; // this pane's loop was restarted
        onMessage(msg);
      }
      onEnd({ reason: "ended" });
    } catch (err) {
      if (err?.name === "AbortError" || /aborted/i.test(err?.message ?? "")) {
        onEnd({ reason: "aborted", error: err });
      } else {
        onEnd({ reason: "error", error: err });
      }
    } finally {
      close();
    }
  })();

  return {
    cancel: close,
    close,
    done,
  };
}
