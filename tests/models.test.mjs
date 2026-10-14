// Unit tests for daemon/runtime/models.mjs — the runtime-aware model
// resolution added in phase 4 (stale-id dropping, catalog validation).

import { test } from "node:test";
import assert from "node:assert/strict";
import { normalizeModelList, pickModel } from "../daemon/runtime/models.mjs";

test("normalizeModelList: strings, SDK objects, and junk", () => {
  assert.deepEqual(normalizeModelList(["a", " b "]), [{ value: "a" }, { value: "b" }]);
  assert.deepEqual(normalizeModelList([{ modelId: "m1", name: "M One" }, { value: "m2" }]), [
    { value: "m1", label: "M One" },
    { value: "m2" },
  ]);
  // displayName / id fallbacks
  assert.deepEqual(normalizeModelList([{ id: "m3", displayName: "M3" }]), [
    { value: "m3", label: "M3" },
  ]);
  // junk entries skipped; empty results → null
  assert.equal(normalizeModelList([]), null);
  assert.equal(normalizeModelList(null), null);
  assert.equal(normalizeModelList([{}, 42, ""]), null);
});

test("pickModel: claude runtime — static catalog with fallback", () => {
  const known = ["sonnet", "opus"];
  assert.deepEqual(
    pickModel({ requested: "sonnet", runtimeId: "claude", knownModels: known, fallback: "F" }),
    { model: "sonnet", dropped: false },
  );
  assert.deepEqual(
    pickModel({ requested: "gone", runtimeId: "claude", knownModels: known, fallback: "F" }),
    { model: "F", dropped: true },
  );
  assert.deepEqual(
    pickModel({ requested: null, runtimeId: "claude", knownModels: known, fallback: "F" }),
    { model: "F", dropped: false },
  );
});

test("pickModel: discovery runtime — no request means the runtime default", () => {
  assert.deepEqual(
    pickModel({ requested: "", runtimeId: "codebuddy", dynamicModels: [{ value: "x" }] }),
    { model: undefined, dropped: false },
  );
});

test("pickModel: discovery runtime — unverified choices rejected before catalog", () => {
  // Cold or failed discovery must not let foreign/stale model IDs through.
  assert.deepEqual(
    pickModel({
      requested: "synthetic-unverified-model",
      runtimeId: "codebuddy",
      dynamicModels: null,
    }),
    { model: undefined, dropped: true },
  );
  // catalog present and includes the id → keep
  assert.deepEqual(
    pickModel({
      requested: "synthetic-model-a",
      runtimeId: "codebuddy",
      dynamicModels: [{ value: "synthetic-model-a" }, { value: "synthetic-model-c" }],
    }),
    { model: "synthetic-model-a", dropped: false },
  );
  // stale id (e.g. a legacy provider id after switching runtime) → dropped
  assert.deepEqual(
    pickModel({
      requested: "sonnet",
      runtimeId: "codebuddy",
      dynamicModels: [{ value: "synthetic-model-a" }],
    }),
    { model: undefined, dropped: true },
  );
});
