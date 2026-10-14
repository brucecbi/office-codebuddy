import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile, readdir, rm, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import {
  planPreviewRegistration,
  installPreviewRegistration,
  uninstallPreviewRegistration,
  previewManifestName,
} from "../app/preview-registration.mjs";
const projectRoot = fileURLToPath(new URL("../", import.meta.url));
const transport = { origin: "https://127.0.0.1:47834" };
async function fixture(t) {
  const home = await mkdtemp(join(tmpdir(), "preview-registration-"));
  t.after(() => rm(home, { recursive: true, force: true }));
  return { home, projectRoot, transport };
}
const dir = (f, host) =>
  join(f.home, "Library", "Containers", `com.microsoft.${host}`, "Data", "Documents", "wef");
test("registration directory symlinks are refused before any write", async (t) => {
  const f = await fixture(t);
  const directory = dir(f, "Excel");
  await mkdir(join(directory, ".."), { recursive: true });
  const target = join(f.home, "unrelated-directory");
  await mkdir(target);
  await symlink(target, directory);
  await assert.rejects(installPreviewRegistration(f), /CONFLICT/);
  assert.deepEqual(await readdir(target), []);
  await assert.rejects(readdir(dir(f, "Word")), { code: "ENOENT" });
});
test("both hosts are checked before any write; legacy file and arbitrary duplicate GUID block registration", async (t) => {
  const f = await fixture(t);
  await mkdir(dir(f, "Excel"), { recursive: true });
  const old = join(dir(f, "Excel"), "excel.xml");
  await writeFile(old, "legacy sentinel");
  await assert.rejects(installPreviewRegistration(f), /PREVIEW_REGISTRATION_CONFLICT/);
  assert.equal(await readFile(old, "utf8"), "legacy sentinel");
  await assert.rejects(readdir(dir(f, "Word")), { code: "ENOENT" });
  await rm(old);
  await writeFile(
    join(dir(f, "Excel"), "another-plugin.xml"),
    await readFile(join(projectRoot, "manifests/excel.xml")),
  );
  await assert.rejects(planPreviewRegistration(f), /PREVIEW_REGISTRATION_CONFLICT/);
});
test("synthetic installation is idempotent and uninstall preserves unrelated manifests", async (t) => {
  const f = await fixture(t);
  await mkdir(dir(f, "Word"), { recursive: true });
  const unrelated = join(dir(f, "Word"), "other.xml");
  await writeFile(unrelated, "<Id>unrelated</Id>");
  const first = await installPreviewRegistration(f);
  assert.equal(first.installed.length, 2);
  assert.deepEqual(await installPreviewRegistration(f), first);
  assert.equal((await uninstallPreviewRegistration(f)).removed.length, 2);
  assert.equal(await readFile(unrelated, "utf8"), "<Id>unrelated</Id>");
});
test("symlink destinations and modified preview registrations are preserved, not replaced", async (t) => {
  const f = await fixture(t);
  await mkdir(dir(f, "Word"), { recursive: true });
  const sentinel = join(f.home, "do-not-touch");
  await writeFile(sentinel, "sentinel");
  const path = join(dir(f, "Word"), previewManifestName("Word"));
  await symlink(sentinel, path);
  await assert.rejects(installPreviewRegistration(f), /CONFLICT/);
  assert.equal(await readFile(sentinel, "utf8"), "sentinel");
  await rm(path);
  await writeFile(path, "modified registration");
  await assert.rejects(uninstallPreviewRegistration(f), /CONFLICT/);
  assert.equal(await readFile(path, "utf8"), "modified registration");
});
