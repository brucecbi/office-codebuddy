// Unit tests for daemon/runtime/messages.mjs — message-layer logic added
// in phase 3: complete-answer fallback dedup, result/error classification,
// and the layered rate-limit signals (migration plan Q9 decision).

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  classifyResult,
  classifyAgentError,
  unstreamedRemainder,
} from "../daemon/runtime/messages.mjs";

test("classifyResult: success carries no error text", () => {
  const r = classifyResult({ type: "result", subtype: "success", is_error: false });
  assert.equal(r.terminal, "success");
  assert.equal(r.kind, "success");
  assert.equal(r.errorText, null);
});

test("classifyResult: deterministic subtype texts (max turns / max budget)", () => {
  const turns = classifyResult({
    subtype: "error_max_turns",
    is_error: true,
    errors: ["hit limit"],
  });
  assert.equal(turns.terminal, "error");
  assert.equal(turns.kind, "max_turns");
  assert.match(turns.errorText, /Turn limit/);
  assert.equal(turns.raw, "hit limit");

  const budget = classifyResult({ subtype: "error_max_budget_usd", is_error: true });
  assert.equal(budget.kind, "max_budget");
  assert.match(budget.errorText, /Budget limit/);
});

test("classifyResult: keyword layer flags rate-limit-shaped errors and keeps the raw text", () => {
  const r = classifyResult({
    subtype: "error_during_execution",
    is_error: true,
    errors: ["Usage limit reached; resets at 18:00"],
  });
  assert.equal(r.kind, "rate_limit");
  assert.match(r.errorText, /Possible usage limit/);
  assert.match(r.errorText, /resets at 18:00/);
});

test("classifyResult: generic errors surface the raw text; empty errors get a fallback", () => {
  const r = classifyResult({
    subtype: "error_during_execution",
    is_error: true,
    errors: ["boom"],
  });
  assert.equal(r.kind, "error");
  assert.equal(r.errorText, "boom");

  const empty = classifyResult({ subtype: "error_during_execution", is_error: true });
  assert.equal(empty.kind, "error");
  assert.match(empty.errorText, /without details/);

  // A result string used as the error carrier (older shape).
  const viaResult = classifyResult({
    subtype: "error_during_execution",
    is_error: true,
    result: "carried text",
  });
  assert.equal(viaResult.errorText, "carried text");
});

test("classifyAgentError: joins message + structured errors[]; flags auth and rate limit", () => {
  const auth = classifyAgentError(
    Object.assign(new Error("Authentication required"), {
      errors: ["Authentication required. Please use /login command to sign in to your account"],
    }),
  );
  assert.equal(auth.isAuth, true);
  assert.match(auth.raw, /Authentication required/);

  const rl = classifyAgentError(
    Object.assign(new Error("Execution failed"), { errors: ["too many requests"] }),
  );
  assert.equal(rl.isRateLimit, true);
  assert.match(rl.text, /Possible usage limit/);

  const plain = classifyAgentError(new Error("socket hang up"));
  assert.equal(plain.isAuth, false);
  assert.equal(plain.isRateLimit, false);
  assert.equal(plain.text, "socket hang up");
});

test("unstreamedRemainder: full fallback / dedup / tail / divergence", () => {
  assert.equal(unstreamedRemainder("hello", ""), "hello"); // nothing streamed
  assert.equal(unstreamedRemainder("hello", "hello"), ""); // fully streamed
  assert.equal(unstreamedRemainder("hello world", "hello "), "world"); // partial stream
  assert.equal(unstreamedRemainder("hello", "hello world"), ""); // streamed longer (pathological)
  assert.equal(unstreamedRemainder("", "abc"), ""); // empty full text
  assert.equal(unstreamedRemainder("abcdef", "abXY"), "ef"); // divergence: prefix budget
});

test("error results that report an authentication failure are classified as auth", () => {
  for (const errors of [["401 Unauthorized"], ["Authentication required"], ["Invalid API key"]]) {
    const c = classifyResult({ type: "result", subtype: "error_during_execution", errors });
    assert.equal(c.kind, "auth", errors[0]);
    assert.equal(c.terminal, "error");
  }
  assert.equal(
    classifyResult({ type: "result", subtype: "error_during_execution", errors: ["boom"] }).kind,
    "error",
  );
  assert.equal(classifyResult({ type: "result", subtype: "success" }).kind, "success");
});
