/* global Word, Excel */
// Opt-in disposable-host acceptance only; never initialized by production UI.
import { toolGetOutline, toolReplaceSection } from "/shared/tools-word.js";
import { toolExcelSortRange } from "/shared/tools-excel.js";
import { createOperationQueue } from "/shared/operation-queue.js";

export async function runNativeWordBoundaryChecks(marker, emit = () => {}) {
  const checks = [];
  const check = async (name, run) => {
    let row;
    try {
      row = { name, status: "PASS", detail: await run() };
    } catch (error) {
      row = { name, status: "FAIL", error: error.message };
    }
    checks.push(row);
    emit(row);
    return row.status === "PASS";
  };
  // Only a fresh marker-only file may be initialized. The server/UI also binds
  // the exact synthetic path; no business document is a permitted input.
  if (
    !(await check("marker_only_admission", () =>
      Word.run(async (context) => {
        const text = context.document.body.getReviewedText(Word.ChangeTrackingVersion.current);
        await context.sync();
        if (text.value.trim() !== marker)
          throw new Error("REFUSED: fresh marker-only document required");
        context.document.changeTrackingMode = "Off";
        const heading = context.document.body.insertParagraph("Boundary target", "End");
        heading.styleBuiltIn = "Heading1";
        const body = heading.insertParagraph("Old boundary body.", "After");
        body.styleBuiltIn = "Normal";
        const next = body.insertParagraph("Boundary next", "After");
        next.styleBuiltIn = "Heading1";
        await context.sync();
        return { initialized: true };
      }),
    ))
  )
    return checks;
  const outline = await toolGetOutline();
  const heading = outline.outline.find((item) => item.text === "Boundary target");
  if (!heading) throw new Error("Native built-in heading missing from outline");
  await check("cancel_during_delete_sync_prevents_new_insert_batch", async () => {
    const receipts = new Map();
    const queue = createOperationQueue({
      storage: {
        getItem: (key) => receipts.get(key),
        setItem: (key, value) => receipts.set(key, value),
      },
    });
    let revokeScheduled = false,
      error;
    try {
      await queue.run(
        "native-section-cancel",
        { mutates: true, deadline: Date.now() + 60000 },
        (control) =>
          toolReplaceSection(
            {
              heading_id: heading.id,
              content: ["FORBIDDEN insertion after Stop"],
              style_per_para: ["Normal"],
              track_changes: false,
            },
            {
              guard: control.guard,
              beforeWrite() {
                control.beforeWrite();
                if (!revokeScheduled) {
                  revokeScheduled = true;
                  // Runs while the already queued deletion is syncing to Word.
                  queueMicrotask(() => queue.cancel());
                }
              },
            },
          ),
      );
    } catch (failure) {
      error = failure.message;
    }
    if (!/^RESULT_UNKNOWN:.*CANCELLED_OR_EXPIRED/.test(error || ""))
      throw new Error("Expected uncertain partial-cancel result: " + error);
    const current = await Word.run(async (context) => {
      const text = context.document.body.getReviewedText(Word.ChangeTrackingVersion.current);
      await context.sync();
      return text.value;
    });
    if (current.includes("Old boundary body.") || current.includes("FORBIDDEN insertion"))
      throw new Error("Delete/readback or new-batch cancellation failed");
    if (receipts.get("draftspect-operation:native-section-cancel") !== "attempted")
      throw new Error("Uncertain receipt must not be completed");
    return {
      error,
      current,
      receipt: "attempted",
      cancellation_source: "test queue during native deletion sync",
    };
  });
  return checks;
}

export async function runNativeExcelBoundaryChecks(marker, emit = () => {}) {
  const checks = [];
  const check = async (name, run) => {
    let row;
    try {
      row = { name, status: "PASS", detail: await run() };
    } catch (error) {
      row = { name, status: "FAIL", error: error.message };
    }
    checks.push(row);
    emit(row);
    return row.status === "PASS";
  };
  let sheet;
  if (
    !(await check("marker_only_workbook_admission", () =>
      Excel.run(async (context) => {
        const ws = context.workbook.worksheets.getActiveWorksheet();
        const used = ws.getUsedRangeOrNullObject();
        ws.load("name");
        used.load("values");
        await context.sync();
        const nonempty = used.isNullObject
          ? []
          : used.values.flat().filter((value) => value !== "" && value != null);
        if (nonempty.length !== 1 || nonempty[0] !== marker)
          throw new Error("REFUSED: fresh marker-only workbook required");
        sheet = ws.name;
        ws.getRange("A3:B5").values = [
          ["key", "value"],
          [2, 20],
          [1, 10],
        ];
        await context.sync();
        return { sheet };
      }),
    ))
  )
    return checks;
  const control = { guard() {}, beforeWrite() {} };
  await check("native_sort_preserves_rows_and_header", async () => {
    const result = await toolExcelSortRange(
      { sheet, address: "A3:B5", key: 0, has_headers: true },
      control,
    );
    return {
      result,
      expected: [
        ["key", "value"],
        [1, 10],
        [2, 20],
      ],
    };
  });
  await check("native_sort_preserves_relative_formulas", async () => {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(sheet);
      ws.getRange("D3:D4").values = [[2], [1]];
      ws.getRange("E3:E4").formulasR1C1 = [["=RC[-1]*10"], ["=RC[-1]*10"]];
      await context.sync();
    });
    return await toolExcelSortRange({ sheet, address: "D3:E4", key: 0 }, control);
  });
  return checks;
}
