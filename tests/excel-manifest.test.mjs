import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { renderManifestTransport } from "../daemon/local-transport.mjs";

test("Excel manifest has a schema-ordered Workbook ShowTaskpane with resolvable resources and version >= 1.0", async () => {
  const xml = await readFile(new URL("../manifests/excel.xml", import.meta.url), "utf8");
  assert.equal(xml.match(/<Id>(.*?)<\/Id>/)[1], "4b5861d6-18ee-4d4a-b9ee-d43d7da80ff4");
  assert.ok(Number(xml.match(/<Version>(.*?)<\/Version>/)[1].split(".")[0]) >= 1);
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
  assert.ok(xml.indexOf("<Hosts>") < xml.indexOf("<Requirements>"));
  const overrides = xml.match(/<VersionOverrides[\s\S]*<\/VersionOverrides>/)[0];
  assert.match(overrides, /<Host xsi:type="Workbook">/);
  assert.match(overrides, /<OfficeTab id="TabHome">/);
  assert.match(overrides, /<Action xsi:type="ShowTaskpane">/);
  const ids = [...overrides.matchAll(/\bid="([^"]+)"/g)].map((match) => match[1]);
  assert.equal(new Set(ids).size, ids.length);
  for (const match of overrides.matchAll(/\bresid="([^"]+)"/g)) {
    assert.ok(match[1].length <= 32);
    assert.ok(ids.includes(match[1]), match[1]);
  }
  const secure = renderManifestTransport(xml, { origin: "https://127.0.0.1:47834" });
  assert.ok(!secure.includes("http://127.0.0.1:47834"));
  assert.equal((secure.match(/https:\/\/127\.0\.0\.1:47834/g) || []).length, 8);
});
