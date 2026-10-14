import { test } from "node:test";
import assert from "node:assert/strict";
import { createSessionPersistenceRegistry } from "../daemon/session-persistence.mjs";

const identity = { host: "word", cwd: "/synthetic", provider: "codebuddy", session_id: "latest" };
const tick = () => new Promise((resolve) => setImmediate(resolve));
const deferred = () => {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
};

test("persistence failure is visible, remains rejecting after Stop, and explicit retry preserves ID", async () => {
  const events = [];
  const calls = [];
  const first = deferred();
  let succeeds = false;
  const registry = createSessionPersistenceRegistry({
    save(...args) {
      calls.push(args);
      return succeeds ? Promise.resolve() : first.promise;
    },
    notify(event, key) {
      events.push({ ...event, key });
    },
  });
  const saving = registry.persist("pane", identity);
  assert.equal(registry.status("pane", identity).status, "saving");
  const failure = Object.assign(new Error("private path intentionally not emitted"), {
    code: "SESSION_LOCK_BUSY",
  });
  first.reject(failure);
  await assert.rejects(saving, { code: "SESSION_LOCK_BUSY" });
  assert.deepEqual(events, [
    { event: "persistence_error", code: "SESSION_LOCK_BUSY", key: "pane" },
  ]);
  // The registry has no live-session pointer, so removal cannot clear this gate.
  await assert.rejects(registry.awaitSaved("pane", identity), { code: "SESSION_LOCK_BUSY" });
  assert.equal(registry.status("pane", identity).status, "failed");
  succeeds = true;
  await registry.retry("pane", identity);
  assert.deepEqual(await registry.awaitSaved("pane", identity), { session_id: "latest" });
  assert.equal(registry.status("pane", identity).status, "saved");
  assert.equal(calls.length, 2);
  assert.deepEqual(calls[0], calls[1]);
});

test("pending persistence blocks handoff and notifications cannot redefine disk success", async () => {
  const waiting = deferred();
  const registry = createSessionPersistenceRegistry({
    save: () => waiting.promise,
    notify() {
      throw new Error("pane disconnected");
    },
  });
  const saving = registry.persist("pane", identity);
  let passed = false;
  const barrier = registry.awaitSaved("pane", identity).then(() => {
    passed = true;
  });
  await tick();
  assert.equal(passed, false);
  waiting.resolve();
  await Promise.all([saving, barrier]);
  assert.equal(passed, true);
  assert.equal(registry.status("pane", identity).status, "saved");
});

test("newer identity serializes after old write, suppresses late old notification and controls the barrier", async () => {
  const old = deferred(),
    latest = deferred();
  const calls = [],
    events = [];
  const registry = createSessionPersistenceRegistry({
    save(_host, _cwd, id) {
      calls.push(id);
      return id === "old" ? old.promise : latest.promise;
    },
    notify(event) {
      events.push(event);
    },
  });
  const oldWrite = registry.persist("pane", { ...identity, session_id: "old" });
  await tick();
  let passed = false;
  const barrier = registry.awaitSaved("pane", identity).then(() => {
    passed = true;
  });
  const newWrite = registry.persist("pane", identity);
  await tick();
  assert.deepEqual(calls, ["old"]);
  old.resolve();
  await oldWrite;
  await tick();
  assert.deepEqual(calls, ["old", "latest"]);
  assert.equal(passed, false);
  assert.equal(registry.status("pane", identity).session_id, "latest");
  assert.deepEqual(events, []);
  latest.resolve();
  await Promise.all([newWrite, barrier]);
  assert.equal(passed, true);
  assert.deepEqual(events, [{ event: "persistence_saved" }]);
});

test("workspace/provider mismatches do not leak IDs and invalidate an in-flight handoff scope", async () => {
  const old = deferred();
  const registry = createSessionPersistenceRegistry({
    save: (_host, cwd) => (cwd === identity.cwd ? old.promise : Promise.resolve()),
  });
  const saving = registry.persist("pane", identity);
  const barrier = registry.awaitSaved("pane", identity);
  assert.equal(registry.status("pane", { ...identity, cwd: "/other" }), null);
  assert.equal(registry.status("pane", { ...identity, provider: "claude" }), null);
  assert.throws(() => registry.retry("pane", { ...identity, cwd: "/other" }), {
    code: "SESSION_PERSISTENCE_SCOPE_CHANGED",
  });
  const changed = registry.persist("pane", { ...identity, cwd: "/other", session_id: "other-id" });
  old.resolve();
  await saving;
  await assert.rejects(barrier, { code: "SESSION_PERSISTENCE_SCOPE_CHANGED" });
  await changed;
  assert.equal(registry.status("pane", identity), null);
});

test("duplicate init does not silently retry failed persistence; retries deduplicate while pending", async () => {
  let calls = 0;
  const retrying = deferred();
  const registry = createSessionPersistenceRegistry({
    save() {
      calls++;
      return calls === 1 ? Promise.reject(new Error("failed")) : retrying.promise;
    },
  });
  const original = registry.persist("pane", identity);
  await assert.rejects(original, /failed/);
  assert.equal(registry.persist("pane", identity), original);
  const retry = registry.retry("pane", identity);
  assert.equal(registry.retry("pane", identity), retry);
  await tick();
  assert.equal(calls, 2);
  retrying.resolve();
  await retry;
});

test("queued persistence freezes its identity and never emits raw error details as a code", async () => {
  const calls = [],
    events = [];
  const registry = createSessionPersistenceRegistry({
    save(...args) {
      calls.push(args);
      throw Object.assign(new Error("private"), { code: "/private/synthetic/path" });
    },
    notify: (event) => events.push(event),
  });
  const mutable = { ...identity };
  const saving = registry.persist("pane", mutable);
  mutable.cwd = "/wrong-target";
  mutable.session_id = "wrong-id";
  await assert.rejects(saving, /private/);
  assert.deepEqual(calls, [["word", "/synthetic", "latest", "codebuddy", "pane"]]);
  assert.deepEqual(events, [{ event: "persistence_error", code: "SESSION_PERSISTENCE_FAILED" }]);
  assert.equal(registry.status("pane", identity).session_id, "latest");
});
