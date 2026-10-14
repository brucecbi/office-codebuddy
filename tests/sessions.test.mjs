// Unit tests for daemon/sessions.mjs — per-(host, workspace) session-id
// bookkeeping. Tests override HOME and DRAFTSPECT_STATE_DIR so an inherited
// caller state directory cannot be read or written. Set both before the fresh
// dynamic import and restore them even when a test callback fails.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

async function withFakeHome(fn) {
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-sess-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  const prevState = process.env.DRAFTSPECT_STATE_DIR;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  process.env.DRAFTSPECT_STATE_DIR = join(fakeHome, ".draftspect-preview");
  try {
    // Cache-bust the dynamic import so each test gets a freshly-evaluated
    // module reading the new $HOME.
    const cacheBuster = "?t=" + Date.now() + Math.random();
    const mod = await import("../daemon/sessions.mjs" + cacheBuster);
    await fn(mod, fakeHome);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    if (prevState === undefined) delete process.env.DRAFTSPECT_STATE_DIR;
    else process.env.DRAFTSPECT_STATE_DIR = prevState;
    await rm(fakeHome, { recursive: true, force: true });
  }
}

async function readStateFile(fakeHome) {
  try {
    const raw = await readFile(join(fakeHome, ".draftspect-preview", "sessions.json"), "utf8");
    return JSON.parse(raw);
  } catch (e) {
    if (e.code === "ENOENT") return null;
    throw e;
  }
}

test("fake HOME isolates inherited state and restores environment even on failure", async () => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-inherited-state-"));
  const inheritedState = join(root, "caller-state");
  const previousState = process.env.DRAFTSPECT_STATE_DIR;
  const sentinel = JSON.stringify({ version: 3, folders: {}, sentinel: "caller-owned" });
  await mkdir(inheritedState);
  await writeFile(join(inheritedState, "sessions.json"), sentinel);
  try {
    for (const inherited of [inheritedState, undefined]) {
      if (inherited === undefined) delete process.env.DRAFTSPECT_STATE_DIR;
      else process.env.DRAFTSPECT_STATE_DIR = inherited;
      await withFakeHome(async ({ saveSessionId }, fakeHome) => {
        assert.equal(process.env.DRAFTSPECT_STATE_DIR, join(fakeHome, ".draftspect-preview"));
        await saveSessionId("word", "/tmp/isolation-fixture", "synthetic-session");
        assert.ok(await readStateFile(fakeHome), "state must be saved inside the fake HOME");
      });
      assert.equal(process.env.DRAFTSPECT_STATE_DIR, inherited);
      await assert.rejects(
        withFakeHome(async () => {
          throw new Error("synthetic callback failure");
        }),
        /synthetic callback failure/,
      );
      assert.equal(process.env.DRAFTSPECT_STATE_DIR, inherited);
    }
    assert.equal(await readFile(join(inheritedState, "sessions.json"), "utf8"), sentinel);
  } finally {
    if (previousState === undefined) delete process.env.DRAFTSPECT_STATE_DIR;
    else process.env.DRAFTSPECT_STATE_DIR = previousState;
    await rm(root, { recursive: true, force: true });
  }
});

test("getSessionId returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionId }) => {
    assert.equal(await getSessionId("word", "/tmp/whatever"), null);
  });
});

test("save → get round-trip is keyed by host AND cwd", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("word", "/tmp/folderA", "word-sid");
    await saveSessionId("excel", "/tmp/folderA", "excel-sid");

    // Same folder, different host → independent session ids.
    assert.equal(await getSessionId("word", "/tmp/folderA"), "word-sid");
    assert.equal(await getSessionId("excel", "/tmp/folderA"), "excel-sid");
    // A host with no saved session here → null.
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);
  });
});

test("normalizeHost: a bad host is a no-op / null", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("powerpoint", "/tmp/x", "nope");
    assert.equal(await getSessionId("powerpoint", "/tmp/x"), null);
    assert.equal(await getSessionId("word", "/tmp/x"), null);
  });
});

test("touchFolder does not clobber a saved session id", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getSessionId }) => {
    await touchFolder("/tmp/folderB");
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);

    await saveSessionId("word", "/tmp/folderB", "w1");
    await touchFolder("/tmp/folderB");
    assert.equal(
      await getSessionId("word", "/tmp/folderB"),
      "w1",
      "touch must not blow away a saved session id",
    );
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder }, fakeHome) => {
    await touchFolder(join(fakeHome, "Library"));
    const state = await readStateFile(fakeHome);
    // touchFolder returns early for $HOME children — nothing persisted.
    assert.ok(
      state === null || !state.folders || Object.keys(state.folders).length === 0,
      "an OS-managed $HOME child must never be persisted",
    );
  });
});

test("migrates a v1 sessions.json: the un-hostable old id is dropped", async () => {
  await withFakeHome(async ({ getSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const v1 = {
      version: 1,
      folders: {
        "/tmp/legacy": {
          session_id: "old-uuid",
          last_used: "2026-05-01T00:00:00.000Z",
          display_name: "legacy",
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(v1));
    // The v1 id can't be attributed to a host, so it's intentionally not
    // resumable post-migration.
    assert.equal(await getSessionId("word", "/tmp/legacy"), null);
    assert.equal(await getSessionId("excel", "/tmp/legacy"), null);
  });
});

test("legacy v2 is not imported automatically; explicit migration keeps providers independent", async () => {
  await withFakeHome(async ({ getSessionId: previewGet, createSessionStore }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const v2 = {
      version: 2,
      folders: {
        "/tmp/v2folder": {
          last_used: "2026-06-01T00:00:00.000Z",
          display_name: "v2folder",
          sessions: { word: "claude-word-sid", excel: "claude-excel-sid" },
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(v2));
    assert.equal(await previewGet("word", "/tmp/v2folder"), null);
    const { getSessionId, saveSessionId } = createSessionStore({
      file: join(fakeHome, ".draftspect-preview", "sessions.json"),
      legacyFile: join(dir, "sessions.json"),
    });

    // v2 ids are Claude-era — they resurface under the default provider…
    assert.equal(await getSessionId("word", "/tmp/v2folder"), "claude-word-sid");
    assert.equal(await getSessionId("word", "/tmp/v2folder", "claude"), "claude-word-sid");
    assert.equal(await getSessionId("excel", "/tmp/v2folder"), "claude-excel-sid");
    // …while another provider has its own, empty namespace.
    assert.equal(await getSessionId("word", "/tmp/v2folder", "codebuddy"), null);

    await saveSessionId("word", "/tmp/v2folder", "cb-word-sid", "codebuddy");
    assert.equal(await getSessionId("word", "/tmp/v2folder", "codebuddy"), "cb-word-sid");
    // …and does not clobber the Claude id.
    assert.equal(await getSessionId("word", "/tmp/v2folder"), "claude-word-sid");
  });
});
