import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import {
  EXCEL_BOOTSTRAP,
  verifyExcelBootstrap,
  openExcelBootstrap,
} from "../app/excel-bootstrap.mjs";
import { transportSettings, renderManifestTransport } from "../daemon/local-transport.mjs";
const transport = transportSettings({});
const xml = () => readFile(new URL("../manifests/excel.xml", import.meta.url), "utf8");
const asset = () => readFile(new URL("../" + EXCEL_BOOTSTRAP.asset, import.meta.url));
async function fixture(t) {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-excel-bootstrap-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const root = join(dir, "project"),
    installedManifestPath = join(dir, "excel.xml");
  await mkdir(join(root, "manifests"), { recursive: true });
  await mkdir(dirname(join(root, EXCEL_BOOTSTRAP.asset)), { recursive: true });
  await writeFile(join(root, "manifests", "excel.xml"), await xml());
  await writeFile(join(root, EXCEL_BOOTSTRAP.asset), await asset());
  await writeFile(installedManifestPath, await xml());
  return { projectRoot: root, installedManifestPath, tempRoot: dir, platform: "darwin", transport };
}
test("Excel shipped bootstrap hash, manifest identity and rendered URLs match", async () => {
  const manifestXml = await xml(),
    secure = { ...transport, origin: "https://127.0.0.1:47834" };
  assert.deepEqual(
    verifyExcelBootstrap({
      manifestXml,
      installedManifestXml: renderManifestTransport(manifestXml, secure),
      assetBytes: await asset(),
      transport: secure,
    }),
    { id: EXCEL_BOOTSTRAP.id, version: EXCEL_BOOTSTRAP.version },
  );
});
for (const [label, replace] of [
  ["GUID", (value) => value.replace(EXCEL_BOOTSTRAP.id, "00000000-0000-0000-0000-000000000000")],
  ["version", (value) => value.replace(EXCEL_BOOTSTRAP.version, "0.1.1.0")],
  ["HTTPS drift", (value) => value.replaceAll("http://127", "https://127")],
  ["command removed", (value) => value.replace(/<VersionOverrides[\s\S]*<\/VersionOverrides>/, "")],
  ["DOCTYPE", (value) => "<!DOCTYPE OfficeApp>" + value],
])
  test(`Excel bootstrap refuses ${label} before opening files`, async (t) => {
    const options = await fixture(t);
    await writeFile(options.installedManifestPath, replace(await xml()));
    await assert.rejects(
      openExcelBootstrap({ ...options, openPath: () => assert.fail("must not open") }),
      /EXCEL_BOOTSTRAP_MANIFEST_(MISMATCH|INVALID)/,
    );
  });
test("Excel bootstrap refuses altered asset and missing install", async (t) => {
  const options = await fixture(t);
  await writeFile(join(options.projectRoot, EXCEL_BOOTSTRAP.asset), "changed");
  await assert.rejects(
    openExcelBootstrap({ ...options, openPath: () => assert.fail() }),
    /ASSET_MISMATCH/,
  );
  await rm(options.installedManifestPath);
  await assert.rejects(
    openExcelBootstrap({ ...options, openPath: () => assert.fail() }),
    /NOT_INSTALLED/,
  );
});
test("Excel initialization opens fresh private copies only and does not equate OS-open with connected", async (t) => {
  const options = await fixture(t),
    bytes = await asset(),
    opened = [];
  const openPath = async (path) => {
    opened.push(path);
    assert.deepEqual(await readFile(path), bytes);
    assert.equal((await stat(path)).mode & 0o777, 0o600);
    return "";
  };
  const a = await openExcelBootstrap({ ...options, openPath }),
    b = await openExcelBootstrap({ ...options, openPath });
  assert.notEqual(a.path, b.path);
  assert.equal(a.requiresEachExcelProcess, true);
  assert.equal(a.connected, undefined);
  assert.equal(opened.length, 2);
  assert.deepEqual(await readFile(join(options.projectRoot, EXCEL_BOOTSTRAP.asset)), bytes);
  assert.equal(await readFile(options.installedManifestPath, "utf8"), await xml());
  await assert.rejects(
    openExcelBootstrap({ ...options, openPath: async () => "No app" }),
    /OPEN_FAILED/,
  );
});
test("Excel fallback requires explicit opener and verified Mac platform", async () => {
  await assert.rejects(
    openExcelBootstrap({ platform: "win32", openPath: () => assert.fail() }),
    /only verified for Excel on macOS/,
  );
  await assert.rejects(openExcelBootstrap({ platform: "darwin" }), /user-triggered openPath/);
});
