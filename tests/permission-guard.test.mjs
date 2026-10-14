import { test } from "node:test";
import assert from "node:assert/strict";
import { bashMutatesOfficeFile, customPermissionHandler } from "../daemon/permission-guard.mjs";

test("Write/Edit/MultiEdit against Office files are denied", async () => {
  for (const tool of ["Write", "Edit", "MultiEdit"]) {
    for (const path of [
      "/w/draft.docx",
      "C:\\w\\Book.XLSX",
      "/w/macro.docm",
      "/w/a.xlsm",
      "/w/old.doc",
    ]) {
      const res = await customPermissionHandler(tool, { file_path: path });
      assert.equal(res.behavior, "deny", `${tool} ${path}`);
      assert.match(res.message, /office_\*/);
    }
  }
  assert.equal((await customPermissionHandler("Write", { path: "/w/a.docx" })).behavior, "deny");
});

test("Write/Edit against non-Office files are allowed with input passed through", async () => {
  const input = { file_path: "/w/notes.md", content: "x" };
  assert.deepEqual(await customPermissionHandler("Write", input), {
    behavior: "allow",
    updatedInput: input,
  });
  assert.equal(
    (await customPermissionHandler("Edit", { file_path: "/w/docxtool.js" })).behavior,
    "allow",
  );
});

test("read-only shell use of Office files is allowed", () => {
  for (const cmd of [
    "unzip -p draft.docx word/document.xml",
    "cat report.docx | wc -c",
    "git log -- report.docx",
    "ls *.docx",
    "unzip -l Book.xlsx",
    'printf x > "/tmp/synthetic review/notes.txt"',
  ]) {
    assert.equal(bashMutatesOfficeFile(cmd), false, cmd);
  }
});

test("shell mutations of Office files are denied", () => {
  for (const cmd of [
    "echo x > active.xlsx",
    "printf x >> 'draft.docx'",
    'printf x > "/tmp/synthetic review/draft.docx"',
    "printf x > '/tmp/synthetic review/draft.docx'",
    "cp draft.docx active.docx",
    "rm -rf workspace/*.docx",
    "mv a.docx b.docx",
    "sed -i '' 's/a/b/' report.docx",
    "zip -u draft.docx word/document.xml",
    "cd tmp && zip -r ../draft.docx .",
    "del report.docx",
    "copy /Y a.docx b.docx",
    "Remove-Item .\\Book.xlsx",
    "Set-Content -Path draft.docx -Value x",
    "Get-Content a.txt | Out-File report.docx",
  ]) {
    assert.equal(bashMutatesOfficeFile(cmd), true, cmd);
  }
});

test("Bash mutations are denied through the handler; unrelated tools are allowed", async () => {
  assert.equal((await customPermissionHandler("Bash", { command: "rm a.docx" })).behavior, "deny");
  assert.equal((await customPermissionHandler("Bash", { command: "rm a.txt" })).behavior, "allow");
  assert.equal(
    (await customPermissionHandler("Read", { file_path: "/w/a.docx" })).behavior,
    "allow",
  );
  assert.deepEqual(await customPermissionHandler("Glob", undefined), {
    behavior: "allow",
    updatedInput: {},
  });
});
