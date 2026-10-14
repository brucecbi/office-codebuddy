import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { toolGetOutline, toolReplaceSection } from "../taskpane/shared/tools-word.js";
import { toolExcelSortRange } from "../taskpane/shared/tools-excel.js";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";

const source = await readFile(
  new URL("../scripts/native-write-boundary-checks.js", import.meta.url),
  "utf8",
);
function harness(Word, Excel) {
  return vm.runInNewContext(
    source.replace(/^import\s+[\s\S]*?;\s*$/gm, "").replace(/^export /gm, "") +
      "\n({ runNativeWordBoundaryChecks, runNativeExcelBoundaryChecks })",
    {
      Word,
      Excel,
      toolGetOutline,
      toolReplaceSection,
      toolExcelSortRange,
      createOperationQueue,
      queueMicrotask,
    },
  );
}

test("native boundary harness refuses non-marker Word content before any initialization", async () => {
  let writes = 0;
  const Word = {
    ChangeTrackingVersion: { current: "Current" },
    run: (fn) =>
      fn({
        sync: async () => {},
        document: {
          set changeTrackingMode(_) {
            writes++;
          },
          body: { getReviewedText: () => ({ value: "Not a synthetic test document" }) },
        },
      }),
  };
  const result = await harness(Word).runNativeWordBoundaryChecks("MARKER");
  assert.equal(result.length, 1);
  assert.equal(result[0].status, "FAIL");
  assert.match(result[0].error, /REFUSED/);
  assert.equal(writes, 0);
});

test("native boundary harness refuses non-marker Excel data before any write", async () => {
  let writes = 0;
  const Excel = {
    run: (fn) =>
      fn({
        sync: async () => {},
        workbook: {
          worksheets: {
            getActiveWorksheet: () => ({
              name: "Sheet1",
              load() {},
              getUsedRangeOrNullObject: () => ({
                load() {},
                isNullObject: false,
                values: [["Not a synthetic workbook"]],
              }),
              getRange: () => {
                writes++;
                throw new Error("must not write");
              },
            }),
          },
        },
      }),
  };
  const result = await harness(undefined, Excel).runNativeExcelBoundaryChecks("MARKER");
  assert.equal(result.length, 1);
  assert.equal(result[0].status, "FAIL");
  assert.match(result[0].error, /REFUSED/);
  assert.equal(writes, 0);
});

test("native boundary harness exercises the actual section tool and uncertain receipt", async () => {
  const fixtureSource = await readFile(
    new URL("word-write-guards.test.mjs", import.meta.url),
    "utf8",
  );
  const fixture = new Function(
    fixtureSource.slice(fixtureSource.indexOf("const W ="), fixtureSource.indexOf("\ntest(")) +
      "\nreturn fixture;",
  )();
  const host = fixture([{ text: "MARKER" }]);
  host.body.getReviewedText = () => ({ value: host.paras.map((p) => p.current).join("\r") });
  host.body.insertParagraph = (text) => host.paras.at(-1).insertParagraph(text);
  const result = await harness(globalThis.Word).runNativeWordBoundaryChecks("MARKER");
  assert.equal(result.length, 2);
  assert.ok(
    result.every((row) => row.status === "PASS"),
    JSON.stringify(result),
  );
  assert.equal(result[1].detail.receipt, "attempted");
  assert.ok(!host.paras.some((p) => p.current.includes("FORBIDDEN insertion")));
});
