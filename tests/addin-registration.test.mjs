import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, dirname } from "node:path";
import {
  inspectManifestRegistration,
  readAddinRegistration,
  officeConnectionDetail,
} from "../app/addin-registration.mjs";
import { renderManifestTransport, transportSettings } from "../daemon/local-transport.mjs";
const plain = transportSettings({});
const secure = { ...plain, origin: "https://127.0.0.1:47834" };
const excel = () => readFile(new URL("../manifests/excel.xml", import.meta.url), "utf8");

test("per-host matching includes HTTPS and the ribbon resources, not file existence", async () => {
  const projectXml = await excel(),
    installedXml = renderManifestTransport(projectXml, secure);
  assert.equal(
    inspectManifestRegistration({ projectXml, installedXml, transport: secure }).state,
    "matched",
  );
  for (const candidate of [
    projectXml,
    installedXml.replace("4b5861d6", "00000000"),
    installedXml.replace("1.0.0.0", "0.1.2.0"),
    installedXml.replace("ReadWriteDocument", "ReadDocument"),
    installedXml.replace(/<VersionOverrides[\s\S]*<\/VersionOverrides>/, ""),
    installedXml.replace("/excel/index.html", "/word/index.html"),
  ])
    assert.equal(
      inspectManifestRegistration({ projectXml, installedXml: candidate, transport: secure }).state,
      "mismatch",
    );
  for (const candidate of [
    "",
    "not XML",
    "<!DOCTYPE OfficeApp>" + installedXml,
    installedXml.replace("</OfficeApp>", "<Id>duplicate</Id></OfficeApp>"),
  ])
    assert.equal(
      inspectManifestRegistration({ projectXml, installedXml: candidate, transport: secure }).state,
      "invalid",
    );
});

test("manifest inspection tolerates comments and inter-element whitespace", async () => {
  const projectXml = await excel();
  assert.equal(
    inspectManifestRegistration({
      projectXml,
      installedXml: projectXml.replace(/>\s+</g, ">\n<!-- comment -->\n<"),
      transport: plain,
    }).state,
    "matched",
  );
});

test("registration reports Word and Excel independently, without writing any settings", async (t) => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-registration-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  for (const [host, name] of [
    ["word", "Word"],
    ["excel", "Excel"],
  ]) {
    const xml = await readFile(new URL(`../manifests/${host}.xml`, import.meta.url), "utf8");
    const project = join(dir, "manifests", `${host}.xml`);
    await mkdir(dirname(project), { recursive: true });
    await writeFile(project, xml);
    if (host === "word") {
      const installed = join(
        dir,
        "Library",
        "Containers",
        `com.microsoft.${name}`,
        "Data",
        "Documents",
        "wef",
        `draftspect-preview-${host}.xml`,
      );
      await mkdir(dirname(installed), { recursive: true });
      await writeFile(installed, renderManifestTransport(xml, secure));
    }
  }
  const options = { projectRoot: dir, home: dir, platform: "darwin", transport: secure };
  assert.deepEqual(
    (await readAddinRegistration(options)).map((row) => [row.host, row.state]),
    [
      ["word", "matched"],
      ["excel", "missing"],
    ],
  );
  assert.ok(
    (await readAddinRegistration({ ...options, platform: "win32" })).every(
      (row) => row.state === "unverified",
    ),
  );
});

test("Office status never promotes a copied manifest, stale heartbeat or invalid count to connected", () => {
  const registration = [{ host: "excel", label: "manifest 匹配（尚不代表已激活）" }];
  const now = 100000,
    health = { phase: "ready", receivedAt: now, connected_panes_by_host: { word: 0, excel: 2 } };
  assert.match(
    officeConnectionDetail(registration, health, now),
    /Excel：.*2 个侧边栏已连接（不代表模型可用）/,
  );
  assert.match(officeConnectionDetail(registration, health, now), /Word：.*尚无侧边栏连接/);
  for (const value of [
    null,
    { ...health, phase: "starting" },
    { ...health, receivedAt: now - 35001 },
    { ...health, receivedAt: now + 1 },
    { ...health, connected_panes_by_host: { excel: -1 } },
    { ...health, connected_panes_by_host: { excel: "1" } },
  ])
    assert.match(officeConnectionDetail(registration, value, now), /Excel：.*侧边栏连接尚未确认/);
});
