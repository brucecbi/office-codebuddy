// Read a resumed SDK session's transcript back into the chat panel.
//
// The Agent SDK persists every session as newline-delimited JSON at
// ~/.claude/projects/<encoded-cwd>/<session_id>.jsonl. We reconstruct the
// user-visible bubble list from it so reopening the add-in (or restarting
// the daemon / Office) shows prior conversation evidence. This is not the
// model's exact active context: compaction, truncation, hidden blocks and
// provider-side processing can differ from the visible history.
//
// We locate the file by globbing for `<session_id>.jsonl` under every
// project dir, so we never depend on Claude Code's (undocumented) cwd ->
// directory-name encoding. Schema observed empirically (2026-05-16):
//
//   { type:"user",      message:{ role:"user",      content:<str|array> } }
//   { type:"assistant", message:{ role:"assistant", content:[ blocks ] } }
//   { type:"queue-operation"|"ai-title"|"attachment"|"last-prompt", ... }  -> skip
//
//   user.content as a string  -> a user-typed message (carries our injected
//                                 [Doc: … · Selection: …] header to strip)
//   user.content as an array  -> tool_result blocks retain identity/output;
//                                genuine text blocks become user bubbles
//   assistant blocks: text -> assistant bubble; tool_use -> tool announce;
//                     thinking -> skip (not shown live either)
//
// The .jsonl schema is the SDK's internal format. The daemon already
// depends on it for live messages, so this is the same coupling, not new.
// Every per-line parse is defensive (bad lines skipped).
// CodeBuddy native JSONL additionally uses top-level message role/content and
// independent function_call/result items; those are normalized below.

import { readdir, stat } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { createInterface } from "node:readline";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

export function projectsDirectory(provider, { env = process.env, home = homedir() } = {}) {
  if (!["claude", "codebuddy"].includes(provider))
    throw new Error(`Unknown transcript provider: ${provider}`);
  const config = provider === "codebuddy" ? env.CODEBUDDY_CONFIG_DIR?.trim() : null;
  return join(
    config ? resolve(config) : join(home, provider === "codebuddy" ? ".codebuddy" : ".claude"),
    "projects",
  );
}

// Strip the per-turn context header the daemon prepends to user messages
// (renderContextHeader: a single bracketed line containing Host:/Doc:/
// Selection:/Track changes:, then a blank line, then the user's actual
// text).
function stripContextHeader(text) {
  return text.replace(
    /^\[(?:Host:|Doc:|Selection:|Cursor in paragraph|Track changes:)[^\n]*\]\n\n?/,
    "",
  );
}

// Find <sessionId>.jsonl under any project dir. Returns the newest match by
// mtime, or null. No dependency on the cwd->dirname encoding scheme.
async function locateSessionFile(
  sessionId,
  { provider = "claude", projectsDir = projectsDirectory(provider) } = {},
) {
  if (!/^[a-zA-Z0-9][a-zA-Z0-9_:-]*$/.test(sessionId)) return null;
  const PROJECTS_DIR = projectsDir;
  let dirs;
  try {
    dirs = await readdir(PROJECTS_DIR, { withFileTypes: true });
  } catch {
    return null;
  }
  const candidates = [];
  for (const d of dirs) {
    if (!d.isDirectory()) continue;
    const p = join(PROJECTS_DIR, d.name, `${sessionId}.jsonl`);
    try {
      const s = await stat(p);
      if (s.isFile()) candidates.push({ path: p, mtime: s.mtimeMs });
    } catch {
      /* not in this dir */
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.mtime - a.mtime);
  return candidates[0].path;
}

// CodeBuddy 2.162 can persist a failed Office MCP call as status:"completed"
// while keeping our error in JSON-encoded text carriers and dropping isError.
// Inspect only result envelopes, never recursively search document payloads:
// a paragraph containing "Error:" is data, not a failed read operation.
function officeResultContainsError(value, depth = 0) {
  if (depth > 6 || value == null) return false;
  if (typeof value === "string") {
    const text = value.trim();
    if (/^Error:\s*\S/.test(text)) return true;
    if (text.length > 1_000_000 || !/^[\[{]/.test(text)) return false;
    try {
      return officeResultContainsError(JSON.parse(text), depth + 1);
    } catch {
      return false;
    }
  }
  if (Array.isArray(value)) return value.some((item) => officeResultContainsError(item, depth + 1));
  if (typeof value !== "object") return false;
  if (value.isError === true || value.is_error === true) return true;
  return value.type === "text" && typeof value.text === "string"
    ? officeResultContainsError(value.text, depth + 1)
    : false;
}

function eventsFromLine(obj, provider = "claude") {
  const out = [];
  if (!obj || typeof obj !== "object") return out;
  if (provider === "codebuddy") {
    // Native JSONL items differ from the SDK's Claude-compatible live stream.
    // callId identifies the tool; id identifies the persisted history item.
    if (obj.type === "function_call") {
      let input = obj.arguments ?? obj.input ?? {};
      if (typeof input === "string") {
        try {
          input = JSON.parse(input);
        } catch {
          // Keep malformed/truncated arguments visible for historical review.
        }
      }
      return [
        {
          kind: "tool",
          id: obj.callId ?? obj.call_id ?? obj.id,
          name: obj.name || "",
          input,
        },
      ];
    }
    if (obj.type === "function_call_output" || obj.type === "function_call_result") {
      const output = obj.output ?? obj.providerData?.toolResult?.content ?? "";
      const failed =
        ["incomplete", "failed", "error", "cancelled", "canceled"].includes(obj.status) ||
        obj.is_error === true ||
        obj.isError === true ||
        !!obj.providerData?.error ||
        !!obj.providerData?.toolResult?.error ||
        obj.providerData?.toolResult?.isError === true ||
        (/^(?:mcp__office__)?office_[a-z0-9_]+$/.test(obj.name || "") &&
          officeResultContainsError(output));
      return [
        {
          kind: "tool_result",
          id: obj.callId ?? obj.call_id ?? obj.id,
          output,
          status: failed ? "error" : (obj.status ?? "completed"),
        },
      ];
    }
  }
  // CodeBuddy adds message:{usage:...} to some native items. Prefer their
  // top-level role/content, while retaining support for older SDK envelopes.
  const nativeMessage =
    provider === "codebuddy" &&
    (obj.role === "user" || obj.role === "assistant") &&
    Object.hasOwn(obj, "content");
  const m = nativeMessage ? obj : (obj.message ?? (provider === "codebuddy" ? obj : null));
  if (!m || typeof m !== "object") return out; // queue-operation/ai-title/etc.
  const role = m.role;
  const content = Array.isArray(m.content)
    ? m.content.map((b) =>
        provider === "codebuddy" && (b?.type === "input_text" || b?.type === "output_text")
          ? { ...b, type: "text" }
          : b,
      )
    : m.content;

  if (obj.type === "user" || role === "user") {
    if (typeof content === "string") {
      const text = stripContextHeader(content).trim();
      if (text) out.push({ kind: "user", text });
    } else if (Array.isArray(content)) {
      for (const b of content) {
        if (b?.type === "tool_result")
          out.push({
            kind: "tool_result",
            id: b.tool_use_id,
            output: b.content,
            status: b.is_error ? "error" : "completed",
          });
      }
      // Tool results above are not user-typed; surface genuine text separately.
      const text = content
        .filter((b) => b && b.type === "text" && typeof b.text === "string")
        .map((b) => b.text)
        .join("")
        .trim();
      if (text) out.push({ kind: "user", text });
    }
    return out;
  }

  if (obj.type === "assistant" || role === "assistant") {
    const blocks = typeof content === "string" ? [{ type: "text", text: content }] : content;
    if (!Array.isArray(blocks)) return out;
    for (const b of blocks) {
      if (!b || typeof b !== "object") continue;
      if (b.type === "text" && typeof b.text === "string") {
        const t = b.text.trim();
        if (t) out.push({ kind: "assistant", text: t });
      } else if (b.type === "tool_use") {
        out.push({
          kind: "tool",
          ...(b.id ? { id: b.id } : {}),
          name: b.name || "",
          input: b.input ?? {},
        });
      }
      // b.type === "thinking" -> skip (not shown in the live UI either)
    }
  }
  return out;
}

/**
 * Reconstruct the chat transcript for a session.
 * @param {string} sessionId
 * @param {{ maxEvents?: number }} opts
 * @returns {Promise<{ events: Array, truncated: boolean }>}
 *   events kinds: {kind:"user",text} | {kind:"assistant",text} |
 *                 {kind:"tool",name,input}
 *   truncated: true if older events were dropped to honor maxEvents.
 */
export async function readTranscript(
  sessionId,
  { maxEvents = 200, provider = "claude", projectsDir } = {},
) {
  if (!sessionId) return { events: [], truncated: false };
  const file = await locateSessionFile(sessionId, {
    provider,
    ...(projectsDir ? { projectsDir } : {}),
  });
  if (!file) return { events: [], truncated: false };

  // Ring buffer of the last maxEvents render events — bounds memory and
  // render time for long sessions.
  const ring = [];
  let total = 0;

  const rl = createInterface({
    input: createReadStream(file, { encoding: "utf8" }),
    crlfDelay: Infinity,
  });
  try {
    for await (const line of rl) {
      if (!line.trim()) continue;
      let obj;
      try {
        obj = JSON.parse(line);
      } catch {
        continue;
      }
      for (const ev of eventsFromLine(obj, provider)) {
        ring.push(ev);
        total++;
        if (ring.length > maxEvents) ring.shift();
      }
    }
  } finally {
    rl.close();
  }

  return { events: ring, truncated: total > ring.length };
}

// Exposed for testing / reuse.
export { stripContextHeader, locateSessionFile, eventsFromLine };
