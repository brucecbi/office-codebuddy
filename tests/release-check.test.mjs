import { test } from "node:test";
import assert from "node:assert/strict";
import { scanText, checkRelease } from "../scripts/check-release.mjs";
import { mkdtemp, writeFile, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createHash } from "node:crypto";
test("release scan detects synthetic secrets and personal paths without returning matched values", () => {
  for (const value of [
    "ghp_" + "Z".repeat(35),
    "/Users/" + "real-person/private",
    "-----BEGIN " + "PRIVATE KEY-----\n" + "A".repeat(32),
  ])
    assert.equal(scanText(value).length, 1);
  assert.deepEqual(scanText("/Users/you/notes /Users/x/synthetic"), []);
});
test("release check rejects drift and unapproved files", async (t) => {
  const root = await mkdtemp(join(tmpdir(), "preview-release-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  await mkdir(join(root, "release"));
  await writeFile(join(root, "ok.txt"), "synthetic");
  await writeFile(
    join(root, "release/approved-files.json"),
    JSON.stringify({
      files: [{ path: "ok.txt", sha256: createHash("sha256").update("synthetic").digest("hex") }],
    }),
  );
  assert.equal((await checkRelease(root)).status, "PASS");
  await writeFile(join(root, "ok.txt"), "changed");
  await writeFile(join(root, "extra.txt"), "unreviewed");
  const result = await checkRelease(root);
  assert.equal(result.status, "FAIL");
  assert.deepEqual(result.errors.map((e) => e.issue).sort(), ["hash-mismatch", "not-allowlisted"]);
});
