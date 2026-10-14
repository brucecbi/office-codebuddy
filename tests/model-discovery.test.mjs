import { test } from "node:test";
import assert from "node:assert/strict";
import { createModelDiscovery } from "../daemon/runtime/model-discovery.mjs";

function fixture({ delayedExit = false, connect = async () => {}, observeFailure = false } = {}) {
  let created = 0,
    closed = 0,
    settings,
    resolveExit;
  const exit = new Promise((resolve) => {
    resolveExit = resolve;
  });
  const discovery = createModelDiscovery({
    options: () => ({ environment: "fixture", hooks: { NEVER: "SECRET" }, tools: ["Bash"] }),
    create(options) {
      created++;
      settings = options;
      return {
        connect,
        getAvailableModels: async () => ["fixture-model"],
        close() {
          closed++;
        },
      };
    },
    observeClose(session) {
      if (observeFailure) throw new Error("Unsupported shape");
      return {
        close() {
          session.close();
          if (!delayedExit) resolveExit();
        },
        waitForExit: () => exit,
      };
    },
  });
  return { discovery, resolveExit, status: () => ({ created, closed, settings }) };
}

test("catalog is metadata-only, deduplicates, caches and observes actual child exit", async () => {
  const f = fixture({ delayedExit: true });
  assert.deepEqual(await Promise.all([f.discovery.list(), f.discovery.list()]), [
    ["fixture-model"],
    ["fixture-model"],
  ]);
  assert.equal(f.status().created, 1);
  assert.equal(f.status().closed, 1);
  assert.deepEqual(f.status().settings.hooks, {});
  assert.deepEqual(f.status().settings.mcpServers, {});
  assert.deepEqual(f.status().settings.settingSources, []);
  assert.deepEqual(f.status().settings.tools, []);
  await assert.rejects(f.discovery.shutdown(5), { code: "RUNTIME_EXIT_UNCONFIRMED" });
  f.resolveExit();
  await f.discovery.shutdown();
  assert.equal(await f.discovery.list({ force: true }), null);
  assert.equal(f.status().created, 1, "shutdown fences new discovery children");
});

test("shutdown during catalog connection closes the child without waiting for the request", async () => {
  let finish;
  const f = fixture({
    connect: () =>
      new Promise((resolve) => {
        finish = resolve;
      }),
  });
  const query = f.discovery.list();
  await f.discovery.shutdown();
  assert.equal(f.status().closed, 1);
  finish();
  assert.equal(await query, null);
});

test("unobservable catalog child never becomes successful shutdown", async () => {
  const f = fixture({ observeFailure: true });
  assert.equal(await f.discovery.list(), null);
  assert.equal(f.status().closed, 1);
  await assert.rejects(f.discovery.shutdown(), /unobservable/);
});
