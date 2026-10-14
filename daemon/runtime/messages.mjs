// Message-layer logic shared by every runtime: turns SDK messages into
// pane-facing decisions and classifies terminal states. Pure functions —
// index.mjs owns the side effects (bridge sends, session flags).
//
// Terminal-state vocabulary (report §4.4 / migration plan phase 3):
//   success       — result message, subtype "success"
//   interrupted   — user pressed Stop (flag consumed at result time)
//   stream_ended  — loop/stream ended with NO result message
//   retrying      — transport is retrying (system/api_retry), not terminal
//   error         — result carries an error subtype, or the iterator threw
//
// Rate-limit classification is layered (migration plan Q9 decision):
//   1. deterministic subtypes: error_max_turns / error_max_budget_usd
//   2. keyword scan of errors[] strings
//   3. ALWAYS surface the raw error text to the pane
// The keyword table is version-noted; re-check it on SDK upgrades — if it
// goes stale the user still sees the raw text (never worse than today).

/** Keyword table for the keyword-matching layer. Re-check on SDK upgrades. */
export const RATE_LIMIT_KEYWORDS =
  /(usage limit|rate limit|daily limit|weekly limit|quota|too many requests|429|limit reached|limit will reset|resets? at|overloaded)/i;

// Deterministic user-facing texts for the structured subtypes.
const MAX_TURNS_TEXT =
  "Turn limit reached for this request (`error_max_turns`). Send a new message to continue.";
const MAX_BUDGET_TEXT =
  "Budget limit reached for this request (`error_max_budget_usd`). Raise the limit or start a new request.";

/**
 * Classify a `result` message.
 * @returns {{
 *   terminal: "success"|"error",
 *   kind: "success"|"max_turns"|"max_budget"|"rate_limit"|"auth"|"error",
 *   errorText: string|null,   // ready to show in the pane (null on success)
 *   raw: string,              // raw errors[] joined ("" when none)
 *   subtype: string,
 * }}
 */
const AUTH_KEYWORDS = /\b(authentication|unauthorized|credential|api[- ]?key|sign[- ]?in|401)\b/i;
const isAuthFailure = (text) => AUTH_KEYWORDS.test(text) || /OAUTH/i.test(text);

export function classifyResult(msg) {
  const subtype = String(msg?.subtype ?? "");
  const rawErrors = Array.isArray(msg?.errors)
    ? msg.errors.filter((e) => typeof e === "string" && e.trim())
    : [];
  const raw = rawErrors.join(" — ");
  const isError =
    msg?.is_error === true ||
    (subtype !== "" && subtype !== "success" && subtype !== "interrupted");

  if (!isError) {
    return { terminal: "success", kind: "success", errorText: null, raw, subtype };
  }

  if (subtype === "error_max_turns") {
    return {
      terminal: "error",
      kind: "max_turns",
      errorText: [MAX_TURNS_TEXT, raw].filter(Boolean).join(" "),
      raw,
      subtype,
    };
  }
  if (subtype === "error_max_budget_usd") {
    return {
      terminal: "error",
      kind: "max_budget",
      errorText: [MAX_BUDGET_TEXT, raw].filter(Boolean).join(" "),
      raw,
      subtype,
    };
  }
  if (raw && RATE_LIMIT_KEYWORDS.test(raw)) {
    return {
      terminal: "error",
      kind: "rate_limit",
      errorText: `Possible usage limit. ${raw}`,
      raw,
      subtype,
    };
  }
  const fallback = typeof msg?.result === "string" && msg.result.trim() ? msg.result.trim() : "";
  const errorText = raw || fallback || "The agent reported an error without details.";
  return {
    terminal: "error",
    kind: isAuthFailure(errorText) ? "auth" : "error",
    errorText,
    raw,
    subtype,
  };
}

/**
 * Classify an exception thrown by the runtime loop (CodeBuddy surfaces SDK
 * errors as thrown ExecutionError before yielding; Claude surfaces aborts
 * this way too — the caller decides what to do per reason).
 * @returns {{ raw: string, isAuth: boolean, isRateLimit: boolean, text: string }}
 */
export function classifyAgentError(error) {
  const parts = [];
  if (error?.message) parts.push(String(error.message));
  if (Array.isArray(error?.errors)) {
    for (const e of error.errors) {
      if (typeof e === "string" && e.trim() && !parts.includes(e)) parts.push(e);
    }
  }
  const raw = parts.join(" — ") || "The agent loop stopped unexpectedly.";
  const isAuth = isAuthFailure(raw);
  const isRateLimit = RATE_LIMIT_KEYWORDS.test(raw);
  const text = isRateLimit ? `Possible usage limit. ${raw}` : raw;
  return { raw, isAuth, isRateLimit, text };
}

/**
 * Dedup helper for the "complete assistant message" fallback: given the
 * final text of an assistant message and the text already shown from
 * stream deltas for the SAME message id, return only the not-yet-shown
 * remainder ("" when fully shown).
 */
export function unstreamedRemainder(fullText, streamedText) {
  const full = String(fullText ?? "");
  const seen = String(streamedText ?? "");
  if (!full) return "";
  if (!seen) return full;
  if (full === seen) return "";
  if (full.startsWith(seen)) return full.slice(seen.length);
  // Divergence (shouldn't normally happen for concatenated deltas): treat
  // the already-shown length as a prefix budget and forward the tail.
  return full.length > seen.length ? full.slice(seen.length) : "";
}
