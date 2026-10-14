import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, symlink, lstat, chmod, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import { randomBytes } from "node:crypto";
import vm from "node:vm";
import { writeRegularFile } from "../daemon/regular-file.mjs";
import { setContextEntries } from "../daemon/context.mjs";

test("regular writer creates and repairs a private token's permissions without trailing bytes", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-regular-file-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const path = join(root, "bridge-token");
  await writeRegularFile(path, "synthetic-long-token", { mode: 0o600 });
  assert.equal((await lstat(path)).mode & 0o777, 0o600);
  await chmod(path, 0o644);
  await writeRegularFile(path, "short", { mode: 0o600 });
  assert.equal(await readFile(path, "utf8"), "short");
  assert.equal((await lstat(path)).mode & 0o777, 0o600);
});

test("token and context writes refuse existing and dangling symlinks without changing their targets", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-no-follow-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const target = join(root, "protected.txt");
  await writeFile(target, "do not change");
  for (const dangling of [false, true]) {
    const workspace = join(root, dangling ? "dangling" : "existing");
    await mkdir(workspace);
    const destination = dangling ? join(root, "absent-target") : target;
    const token = join(workspace, "bridge-token");
    await symlink(destination, token);
    await assert.rejects(writeRegularFile(token, "synthetic-token", { mode: 0o600 }), /symlink/);
    await symlink(destination, join(workspace, "CLAUDE.md"));
    await assert.rejects(setContextEntries(workspace, []), /symlink/);
    assert.equal(await readFile(target, "utf8"), "do not change");
    if (dangling) await assert.rejects(lstat(destination), { code: "ENOENT" });
  }
});

test("actual daemon token startup block refuses a linked token or immediate state directory", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-token-block-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const block = source.slice(source.indexOf("const BRIDGE_TOKEN ="), source.indexOf("// CLI args"));
  const target = join(root, "protected.txt");
  await writeFile(target, "preserved");
  const state = join(root, "real-state");
  await mkdir(state);
  const execute = (token) =>
    vm.runInNewContext(`(async () => {${block}})()`, {
      TOKEN_FILE: token,
      dirname,
      mkdir,
      lstat,
      randomBytes,
      writeRegularFile,
      console: { log() {} },
    });
  const token = join(state, "bridge-token");
  await symlink(target, token);
  await assert.rejects(execute(token), /symlink/);
  const link = join(root, "linked-state");
  await symlink(state, link);
  await assert.rejects(execute(join(link, "bridge-token")), /real directory/);
  assert.equal(await readFile(target, "utf8"), "preserved");
  await rm(token);
  await execute(token);
  assert.equal((await readFile(token, "utf8")).length, 48);
  assert.equal((await lstat(token)).mode & 0o777, 0o600);
});

test("regular writer refuses directory targets and normal context writes preserve existing text", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-regular-context-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await assert.rejects(writeRegularFile(root, "not a file"), /non-regular/);
  await writeFile(join(root, "CLAUDE.md"), "# Existing instructions\n", { mode: 0o644 });
  await setContextEntries(root, []);
  const content = await readFile(join(root, "CLAUDE.md"), "utf8");
  assert.match(content, /^# Existing instructions/);
  assert.equal((await lstat(join(root, "CLAUDE.md"))).mode & 0o777, 0o644);
});
