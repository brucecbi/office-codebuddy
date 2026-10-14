import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile, mkdir, mkdtemp, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { WORD_BOOTSTRAP, verifyWordBootstrap, openWordBootstrap } from "../app/word-bootstrap.mjs";

const projectRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const sourceManifest = () => readFile(join(projectRoot, "manifests", "word.xml"), "utf8");
const sourceAsset = () => readFile(join(projectRoot, WORD_BOOTSTRAP.asset));

async function fixture(t) {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-bootstrap-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const root = join(dir, "project");
  const installedManifestPath = join(dir, "installed-word.xml");
  await mkdir(join(root, "manifests"), { recursive: true });
  await mkdir(dirname(join(root, WORD_BOOTSTRAP.asset)), { recursive: true });
  const [manifestXml, assetBytes] = await Promise.all([sourceManifest(), sourceAsset()]);
  await writeFile(join(root, "manifests", "word.xml"), manifestXml);
  await writeFile(join(root, WORD_BOOTSTRAP.asset), assetBytes);
  await writeFile(installedManifestPath, manifestXml);
  return { projectRoot: root, installedManifestPath, tempRoot: dir, platform: "darwin" };
}

test("shipped synthetic bootstrap hash and identity match the Word manifest", async () => {
  const manifestXml = await sourceManifest();
  assert.deepEqual(
    verifyWordBootstrap({
      manifestXml,
      installedManifestXml: manifestXml,
      assetBytes: await sourceAsset(),
    }),
    { id: WORD_BOOTSTRAP.id, version: WORD_BOOTSTRAP.version },
  );
});

test("bootstrap rejects changed asset bytes before opening anything", async (t) => {
  const options = await fixture(t);
  await writeFile(join(options.projectRoot, WORD_BOOTSTRAP.asset), "changed fixture");
  await assert.rejects(
    openWordBootstrap({ ...options, openPath: () => assert.fail("must not open") }),
    /WORD_BOOTSTRAP_ASSET_MISMATCH/,
  );
});

for (const [label, target, replace] of [
  [
    "repo version",
    "repo",
    (xml) => xml.replace("<Version>1.0.0.1</Version>", "<Version>1.0.0.2</Version>"),
  ],
  [
    "repo GUID",
    "repo",
    (xml) => xml.replace(WORD_BOOTSTRAP.id, "00000000-0000-0000-0000-000000000000"),
  ],
  [
    "installed version",
    "installed",
    (xml) => xml.replace("<Version>1.0.0.1</Version>", "<Version>1.0.0.0</Version>"),
  ],
]) {
  test(`bootstrap fails closed on ${label} mismatch`, async (t) => {
    const options = await fixture(t);
    const path =
      target === "repo"
        ? join(options.projectRoot, "manifests", "word.xml")
        : options.installedManifestPath;
    await writeFile(path, replace(await readFile(path, "utf8")));
    await assert.rejects(
      openWordBootstrap({ ...options, openPath: () => assert.fail("must not open") }),
      /WORD_BOOTSTRAP_MANIFEST_MISMATCH/,
    );
  });
}

test("bootstrap requires the Word manifest to have already been installed", async (t) => {
  const options = await fixture(t);
  await rm(options.installedManifestPath);
  await assert.rejects(
    openWordBootstrap({ ...options, openPath: () => assert.fail("must not open") }),
    /WORD_BOOTSTRAP_NOT_INSTALLED/,
  );
});

test("bootstrap opens only a fresh private copy and preserves source files", async (t) => {
  const options = await fixture(t);
  const original = await sourceAsset();
  const manifestBefore = await readFile(options.installedManifestPath);
  const opened = [];
  const openPath = async (path) => {
    opened.push(path);
    assert.deepEqual(await readFile(path), original);
    assert.equal((await stat(path)).mode & 0o777, 0o600);
    return "";
  };
  const first = await openWordBootstrap({ ...options, openPath });
  const second = await openWordBootstrap({ ...options, openPath });
  assert.equal(opened.length, 2);
  assert.notEqual(first.path, second.path);
  assert.equal(first.requiresEachWordProcess, true);
  assert.deepEqual(await readFile(join(options.projectRoot, WORD_BOOTSTRAP.asset)), original);
  assert.deepEqual(await readFile(options.installedManifestPath), manifestBefore);
});

test("bootstrap propagates shell.openPath errors without claiming initialization", async (t) => {
  const options = await fixture(t);
  await assert.rejects(
    openWordBootstrap({
      ...options,
      openPath: async () => "No application is associated with this file",
    }),
    /WORD_BOOTSTRAP_OPEN_FAILED: No application/,
  );
});

test("bootstrap does not run on an unverified platform", async () => {
  await assert.rejects(
    openWordBootstrap({ platform: "win32", openPath: () => assert.fail("must not open") }),
    /only verified for Word on macOS/,
  );
});

test("manifest guard rejects ambiguous or entity-bearing XML", async () => {
  const assetBytes = await sourceAsset();
  const manifestXml = await sourceManifest();
  for (const xml of [
    manifestXml.replace("</OfficeApp>", "<Id>duplicate</Id></OfficeApp>"),
    `<!DOCTYPE OfficeApp [<!ENTITY version "1.0.0.1">]>${manifestXml}`,
  ]) {
    assert.throws(
      () =>
        verifyWordBootstrap({ manifestXml: xml, installedManifestXml: manifestXml, assetBytes }),
      /WORD_BOOTSTRAP_MANIFEST_INVALID/,
    );
  }
});
