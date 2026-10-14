import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm, readFile, stat, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createPreviewTls } from "../scripts/setup-preview-tls.mjs";

test("explicit TLS setup creates host-only private material without trust or listeners; refuses overwrite", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "preview-tls-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const env = { PATH: process.env.PATH, DRAFTSPECT_STATE_DIR: root };
  const result = await createPreviewTls({ env });
  assert.equal(result.trustInstalled, false);
  assert.equal(result.listenersStarted, false);
  assert.equal(result.certificate.native_trust, "NOT_CHECKED");
  const bytes = await readFile(result.configPath);
  const config = JSON.parse(bytes);
  assert.equal((await stat(config.keyPath)).mode & 0o077, 0);
  assert.equal((await stat(result.configPath)).mode & 0o077, 0);
  assert.ok(!JSON.stringify(result).includes("PRIVATE KEY"));
  await assert.rejects(createPreviewTls({ env }), /PREVIEW_TLS_EXISTS/);
  assert.deepEqual(await readFile(result.configPath), bytes);
});
test("TLS setup refuses a state-directory symlink without modifying target", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "preview-tls-link-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const link = join(root, "link");
  await symlink(root, link);
  await assert.rejects(
    createPreviewTls({ env: { PATH: process.env.PATH, DRAFTSPECT_STATE_DIR: link } }),
    /PREVIEW_TLS_DIRECTORY/,
  );
});
