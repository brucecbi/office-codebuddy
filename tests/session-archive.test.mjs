import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, stat, mkdir, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
import { archiveSession, resumeCommands, shellQuote } from "../daemon/session-archive.mjs";
import { createSessionStore } from "../daemon/sessions.mjs";

test("archive preserves raw transcript bytes/hash and creates distinct immutable packages", async () => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-archive-"));
  const source = join(root, "source.jsonl");
  const bytes = Buffer.from(
    '{"role":"user","content":"合成 fixture"}\n{"role":"assistant","content":"OK"}\n',
  );
  await writeFile(source, bytes);
  const info = { session_id: "fixture-id", cwd: "/fixture", host: "word", provider: "codebuddy" };
  const options = { root: join(root, "archives"), locate: async () => source };
  const a = await archiveSession(info, options),
    b = await archiveSession(info, options);
  assert.notEqual(a.archive_path, b.archive_path);
  assert.deepEqual(await readFile(join(a.archive_path, "session.jsonl")), bytes);
  assert.deepEqual(await readFile(source), bytes);
  assert.equal(a.transcript_sha256, createHash("sha256").update(bytes).digest("hex"));
  assert.equal((await stat(join(a.archive_path, "session.jsonl"))).mode & 0o777, 0o600);
  assert.match(
    await readFile(join(a.archive_path, "README.md"), "utf8"),
    /not a Word document backup/,
  );
  assert.equal(
    JSON.parse(await readFile(join(a.archive_path, "manifest.json"))).session_id,
    "fixture-id",
  );
});
test("resume commands quote paths safely, use an exact ID and offer explicit fork", () => {
  const cmd = resumeCommands({
    provider: "codebuddy",
    session_id: "id_123",
    cwd: "/fixture/it's $(not-a-command)",
  });
  assert.equal(shellQuote("a'b"), "'a'\\''b'");
  assert.match(cmd.resume, /--resume 'id_123'/);
  assert.ok(cmd.resume.includes("'/fixture/it'\\''s $(not-a-command)'"));
  assert.ok(!cmd.resume.includes("--continue"));
  assert.ok(cmd.reuse_as_new_session.includes("--fork-session"));
  assert.throws(
    () => resumeCommands({ provider: "codebuddy", session_id: "../../evil", cwd: "/fixture" }),
    /Invalid/,
  );
});
test("missing transcript is an explicit archive failure", async () => {
  await assert.rejects(
    archiveSession({ session_id: "missing", provider: "codebuddy" }, { locate: async () => null }),
    /No persisted transcript/,
  );
});
test("archive preserves session-local attachments with hashes and refuses symlinks", async () => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-attachments-"));
  const source = join(root, "fixture-id.jsonl");
  await writeFile(source, "{}\n");
  await mkdir(join(root, "fixture-id", "tool-results"), { recursive: true });
  await writeFile(
    join(root, "fixture-id", "tool-results", "read.txt"),
    "Synthetic full tool result",
  );
  const info = { session_id: "fixture-id", provider: "codebuddy", cwd: "/fixture", host: "word" };
  const a = await archiveSession(info, {
    root: join(root, "archives"),
    locate: async () => source,
  });
  assert.equal(a.session_files.length, 1);
  const file = a.session_files[0];
  assert.equal(
    await readFile(join(a.archive_path, file.path), "utf8"),
    "Synthetic full tool result",
  );
  assert.equal(
    file.sha256,
    createHash("sha256").update("Synthetic full tool result").digest("hex"),
  );
  await symlink(source, join(root, "fixture-id", "external-link"));
  await assert.rejects(
    archiveSession(info, { root: join(root, "archives"), locate: async () => source }),
    /refuses links/,
  );
});
test("serial release survives reopening the store and stays document/provider scoped", async () => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-handoff-"));
  const options = { file: join(dir, "sessions.json"), legacyFile: null };
  const a = createSessionStore(options);
  await a.saveSessionId("word", "/fixture", "id", "codebuddy", "doc-A");
  await a.setHandoff("word", "/fixture", "codebuddy", "doc-A", {
    state: "released",
    session_id: "id",
  });
  const b = createSessionStore(options);
  assert.equal((await b.getHandoff("word", "/fixture", "codebuddy", "doc-A")).state, "released");
  assert.equal(await b.getHandoff("word", "/fixture", "codebuddy", "doc-B"), null);
  assert.equal(await b.getHandoff("word", "/fixture", "claude", "doc-A"), null);
  await b.setHandoff("word", "/fixture", "codebuddy", "doc-A", null);
  assert.equal(await a.getHandoff("word", "/fixture", "codebuddy", "doc-A"), null);
  assert.equal(await a.getSessionId("word", "/fixture", "codebuddy", "doc-A"), "id");
});
