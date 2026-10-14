import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { openInMicrosoftExcel } from "../app/macos-excel-open.mjs";

test("Excel bootstrap opens the explicit Excel bundle with an argument array, not the default file association", async () => {
  const path = "/private/tmp/Tester's 中文 folder/synthetic workbook.xlsx";
  let observed;
  const result = await openInMicrosoftExcel(path, {
    platform: "darwin",
    spawnImpl: (...args) => {
      observed = args;
      const child = new EventEmitter();
      queueMicrotask(() => child.emit("close", 0, null));
      return child;
    },
  });
  assert.equal(result, "");
  assert.deepEqual(observed, [
    "/usr/bin/open",
    ["-b", "com.microsoft.Excel", path],
    {
      stdio: "ignore",
      timeout: 10000,
      shell: false,
    },
  ]);
});

test("Excel opener refuses unverified platforms and malformed paths before spawning", async () => {
  const spawnImpl = () => assert.fail("must not spawn");
  assert.ok(await openInMicrosoftExcel("/tmp/synthetic.xlsx", { platform: "win32", spawnImpl }));
  for (const path of [
    undefined,
    {},
    "relative.xlsx",
    "/tmp/new\nline.xlsx",
    "/tmp/tab\tfile.xlsx",
    "/tmp/null\0file.xlsx",
  ])
    assert.ok(await openInMicrosoftExcel(path, { platform: "darwin", spawnImpl }));
});

test("Excel opener reports process failures without exposing arbitrary diagnostic payloads", async () => {
  for (const mode of ["throw", "error", "exit", "signal"]) {
    const result = await openInMicrosoftExcel("/tmp/synthetic.xlsx", {
      platform: "darwin",
      spawnImpl: () => {
        if (mode === "throw") throw new Error("SECRET_TEST_PAYLOAD");
        const child = new EventEmitter();
        queueMicrotask(() => {
          if (mode === "error") child.emit("error", new Error("SECRET_TEST_PAYLOAD"));
          else
            child.emit("close", mode === "exit" ? 1 : null, mode === "signal" ? "SIGTERM" : null);
        });
        return child;
      },
    });
    assert.match(result, /could not open/);
    assert.doesNotMatch(result, /SECRET_TEST_PAYLOAD/);
  }
});
