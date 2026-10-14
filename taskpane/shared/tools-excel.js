/* global Excel */
//
// Excel tool implementations + Excel-specific helpers.
//
// Each `toolExcel*` function is invoked from the dispatcher in taskpane.js
// when a matching `excel_*` tool_call arrives from the daemon. All operations
// go through Excel.run for proper context lifecycle. Address strings use
// A1 notation, optionally sheet-qualified (e.g. "Sheet1!A1:C5").

function _splitSheetAddress(address, fallbackSheetName) {
  // Returns { sheetName, a1 } from "Sheet1!A1:B2" or "A1:B2" (uses fallback).
  if (!address) return { sheetName: fallbackSheetName, a1: null };
  const m = /^(?:'([^']+)'|([^!]+))!(.+)$/.exec(address);
  if (m) return { sheetName: m[1] || m[2], a1: m[3] };
  return { sheetName: fallbackSheetName, a1: address };
}

// Parse the top-left cell of an A1-notation range into 1-based row/col.
// "C5" → { row: 5, col: 3 }; "AA10" → { row: 10, col: 27 }; "C5:F10" → uses C5.
// Used so excel_find_value reports absolute spreadsheet coordinates, not
// 0-based offsets within whatever used-range happens to start at.
function _topLeftCell(a1OrRange) {
  const a1 = String(a1OrRange).split(":")[0];
  const m = /^([A-Za-z]+)(\d+)$/.exec(a1);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10), col };
}

// Spreadsheet column number → A1 letters. 1→"A", 27→"AA", etc.
function _colNumberToLetters(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

async function _activeSheetName(context) {
  const ws = context.workbook.worksheets.getActiveWorksheet();
  ws.load("name");
  await context.sync();
  return ws.name;
}

function armExcelWrite(control) {
  control?.guard();
  control?.beforeWrite();
}

function sameCell(actual, expected) {
  if (Object.is(actual, expected)) return true;
  const got = actual == null ? "" : String(actual);
  const want = expected == null ? "" : String(expected);
  return got === want;
}

function sameGrid(actual, expected) {
  if (!Array.isArray(actual) || actual.length !== expected.length) return false;
  return expected.every(
    (row, i) =>
      Array.isArray(row) &&
      Array.isArray(actual[i]) &&
      actual[i].length === row.length &&
      row.every((cell, j) => sameCell(actual[i][j], cell)),
  );
}

function requireReadback(ok, detail) {
  if (!ok) throw new Error("RESULT_UNKNOWN: " + detail);
}

function addressTail(value) {
  return String(value ?? "")
    .split("!")
    .pop()
    .replace(/\$/g, "")
    .toUpperCase();
}

function columnOrdered(values, key, ascending, hasHeaders) {
  const rows = hasHeaders ? values.slice(1) : values;
  for (let i = 1; i < rows.length; i++) {
    const cmp = String(rows[i - 1]?.[key] ?? "").localeCompare(
      String(rows[i]?.[key] ?? ""),
      undefined,
      {
        numeric: true,
      },
    );
    if (ascending ? cmp > 0 : cmp < 0) return false;
  }
  return true;
}

function sortRowKeys(values, formulas) {
  if (!Array.isArray(values) || !Array.isArray(formulas) || values.length !== formulas.length)
    throw new Error("READBACK_UNAVAILABLE: sort requires values and relative formulas.");
  return values.map((row, i) => {
    if (!Array.isArray(row) || !Array.isArray(formulas[i]) || row.length !== formulas[i].length)
      throw new Error("READBACK_UNAVAILABLE: sort row/formula shapes differ.");
    // R1C1 preserves relative row formulas when rows move. Compare formula
    // identity, not cached results that can legitimately change on recalc.
    // Constants keep their type; flattening or string-coercion loses row links.
    return JSON.stringify(
      row.map((value, j) => {
        const formula = formulas[i][j];
        if (typeof formula === "string" && formula.startsWith("=")) return ["formula", formula];
        const cell = value == null ? "" : value;
        return [typeof cell, cell];
      }),
    );
  });
}

function sameRowMultiset(before, after) {
  const left = before.slice().sort();
  const right = after.slice().sort();
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

export async function toolExcelGetSelectedRange() {
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values, rowCount, columnCount, worksheet/name");
    await context.sync();
    return {
      address: range.address,
      sheet: range.worksheet.name,
      row_count: range.rowCount,
      column_count: range.columnCount,
      values: range.values,
    };
  });
}

export async function toolExcelListSheets() {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name, items/position");
    const active = context.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await context.sync();
    // Load used range per sheet (separate sync because we needed names first).
    const used = sheets.items.map((s) => {
      const r = s.getUsedRangeOrNullObject(true);
      r.load("address, rowCount, columnCount, isNullObject");
      return { sheet: s, used: r };
    });
    await context.sync();
    return {
      active_sheet: active.name,
      sheets: used.map(({ sheet, used: u }) => ({
        name: sheet.name,
        position: sheet.position,
        used_range: u.isNullObject ? null : u.address,
        row_count: u.isNullObject ? 0 : u.rowCount,
        column_count: u.isNullObject ? 0 : u.columnCount,
      })),
    };
  });
}

export async function toolExcelReadRange({ address = null, sheet = null }) {
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    // For whole-sheet reads (no explicit address), use the null-object
    // variant so an empty sheet returns gracefully instead of throwing
    // an InvalidArgument from Excel.js.
    const range = a1 ? ws.getRange(a1) : ws.getUsedRangeOrNullObject(true);
    range.load("address, values, rowCount, columnCount, isNullObject");
    await context.sync();
    if (!a1 && range.isNullObject) {
      return {
        sheet: sheetName,
        address: null,
        row_count: 0,
        column_count: 0,
        values: [],
        empty: true,
      };
    }
    return {
      sheet: sheetName,
      address: range.address,
      row_count: range.rowCount,
      column_count: range.columnCount,
      values: range.values,
    };
  });
}

export async function toolExcelWriteRange({ address, values, sheet = null }, control) {
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error("`values` must be a non-empty 2D array.");
  }
  // Ragged-array guard: every row must have the same length as values[0],
  // otherwise Excel.js fails with a confusing internal error (or silently
  // pads with undefined on some builds). Surface the offending row index.
  const expectedCols = values[0].length;
  for (let i = 1; i < values.length; i++) {
    if (!Array.isArray(values[i]) || values[i].length !== expectedCols) {
      const got = Array.isArray(values[i]) ? `${values[i].length} elements` : "not an array";
      throw new Error(
        `Ragged values: row ${i} is ${got} but row 0 has ${expectedCols}. ` +
          `All rows must be the same length.`,
      );
    }
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    if (!a1) throw new Error("`address` is required for write_range.");
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (range.rowCount !== values.length || range.columnCount !== expectedCols) {
      throw new Error(
        `Shape mismatch: range ${range.address} is ${range.rowCount}×${range.columnCount} ` +
          `but values is ${values.length}×${expectedCols}.`,
      );
    }
    armExcelWrite(control);
    range.values = values;
    await context.sync();
    range.load("values");
    await context.sync();
    if (!sameGrid(range.values, values))
      throw new Error("RESULT_UNKNOWN: written values did not read back.");
    return {
      sheet: sheetName,
      address: range.address,
      written: values.length * expectedCols,
      verification: "values_readback",
      readback: range.values,
    };
  });
}

export async function toolExcelFindValue({
  query,
  sheet = null,
  match_case = false,
  whole_cell = false,
}) {
  if (typeof query !== "string" || !query.length) throw new Error("`query` is required.");
  return await Excel.run(async (context) => {
    const targetSheets = sheet
      ? [context.workbook.worksheets.getItem(sheet)]
      : (() => {
          const all = context.workbook.worksheets;
          all.load("items/name");
          return all;
        })();
    if (Array.isArray(targetSheets)) {
      // single sheet case — load name for the result
      targetSheets[0].load("name");
    }
    await context.sync();
    const sheetsToScan = Array.isArray(targetSheets) ? targetSheets : targetSheets.items;
    const matches = [];
    for (const ws of sheetsToScan) {
      const used = ws.getUsedRangeOrNullObject(true);
      used.load("address, values, rowCount, columnCount, isNullObject");
      await context.sync();
      if (used.isNullObject) continue;
      // The used range may start anywhere on the sheet — e.g. "Sheet1!C5:F10".
      // Translate per-cell offsets within it into absolute spreadsheet rows
      // and columns so the agent's row/column numbers match what the user
      // sees in Excel's row/column headers.
      const origin = _topLeftCell(used.address) || { row: 1, col: 1 };
      const q = match_case ? query : query.toLowerCase();
      for (let r = 0; r < used.rowCount; r++) {
        for (let c = 0; c < used.columnCount; c++) {
          const v = used.values[r][c];
          if (v === null || v === undefined || v === "") continue;
          const s = match_case ? String(v) : String(v).toLowerCase();
          const hit = whole_cell ? s === q : s.includes(q);
          if (hit) {
            const absRow = origin.row + r;
            const absCol = origin.col + c;
            matches.push({
              sheet: ws.name,
              row: absRow,
              column: absCol,
              address: `${_colNumberToLetters(absCol)}${absRow}`,
              value: v,
            });
          }
        }
      }
    }
    return { query, match_count: matches.length, matches };
  });
}

export async function toolExcelInsertRows({ sheet = null, at, count = 1 }, control) {
  if (!Number.isInteger(at) || at < 1) throw new Error("`at` must be a positive integer.");
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    const target = ws.getRange(`${at}:${at + count - 1}`);
    armExcelWrite(control);
    const inserted = target.insert(Excel.InsertShiftDirection.down);
    inserted.load("rowCount");
    await context.sync();
    requireReadback(inserted.rowCount === count, "inserted row count did not read back.");
    return {
      sheet: activeName,
      inserted_at: at,
      count,
      verification: "row_count_readback",
    };
  });
}

export async function toolExcelDeleteRows({ sheet = null, at, count = 1 }, control) {
  if (!Number.isInteger(at) || at < 1) throw new Error("`at` must be a positive integer.");
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    const target = ws.getRange(`${at}:${at + count - 1}`);
    const below = ws.getRange(`${at + count}:${at + count}`);
    below.load("values");
    await context.sync();
    const expected = below.values?.[0] ?? [];
    armExcelWrite(control);
    target.delete(Excel.DeleteShiftDirection.up);
    await context.sync();
    const now = ws.getRange(`${at}:${at}`);
    now.load("values");
    await context.sync();
    requireReadback(
      sameGrid([now.values?.[0] ?? []], [expected]),
      "row below the deletion did not shift into place.",
    );
    return {
      sheet: activeName,
      deleted_at: at,
      count,
      verification: "shifted_row_readback",
    };
  });
}

export async function toolExcelSelectRange({ address, sheet = null }) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range like 'B4' or 'A1:D9') is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    ws.activate(); // make sure the selection is visible on the right sheet
    const range = ws.getRange(address);
    range.select();
    range.load("address");
    await context.sync();
    return { sheet: activeName, address: range.address, selected: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 1 editing tools — formulas / formatting / columns / sheets / clear
// ---------------------------------------------------------------------------

// Tool: excel_write_formula — like write_range but sets formulas (write_range
// is values-only). Each cell gets an A1/R1C1 formula string, e.g. "=SUM(A1:A9)".
export async function toolExcelWriteFormula({ address, formulas, sheet = null }, control) {
  if (!Array.isArray(formulas) || !formulas.length || !Array.isArray(formulas[0])) {
    throw new Error("`formulas` must be a non-empty 2D array.");
  }
  const cols = formulas[0].length;
  for (let i = 1; i < formulas.length; i++) {
    if (!Array.isArray(formulas[i]) || formulas[i].length !== cols) {
      throw new Error(
        `Ragged formulas: row ${i} has ${formulas[i]?.length} but row 0 has ${cols}.`,
      );
    }
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    if (!a1) throw new Error("`address` is required for write_formula.");
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (range.rowCount !== formulas.length || range.columnCount !== cols) {
      throw new Error(
        `Shape mismatch: range ${range.address} is ${range.rowCount}×${range.columnCount} ` +
          `but formulas is ${formulas.length}×${cols}.`,
      );
    }
    armExcelWrite(control);
    range.formulas = formulas;
    await context.sync();
    range.load("formulas");
    await context.sync();
    if (!sameGrid(range.formulas, formulas))
      throw new Error("RESULT_UNKNOWN: written formulas did not read back.");
    return {
      sheet: sheetName,
      address: range.address,
      written: formulas.length * cols,
      verification: "formulas_readback",
      readback: range.formulas,
    };
  });
}

// Tool: excel_set_format — number format / font / fill / borders on a range.
export async function toolExcelSetFormat(
  {
    address,
    sheet = null,
    number_format,
    bold,
    italic,
    font_size,
    font_name,
    font_color,
    fill_color,
    border,
  },
  control,
) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range) is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    armExcelWrite(control);
    if (number_format !== undefined) {
      // numberFormat expects a 2D array matching the range shape.
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => number_format),
      );
    }
    if (bold !== undefined) range.format.font.bold = bold;
    if (italic !== undefined) range.format.font.italic = italic;
    if (font_size !== undefined) range.format.font.size = font_size;
    if (font_name !== undefined) range.format.font.name = font_name;
    if (font_color !== undefined) range.format.font.color = font_color;
    if (fill_color !== undefined) range.format.fill.color = fill_color;
    if (border) {
      for (const edge of [
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideHorizontal",
        "InsideVertical",
      ]) {
        const b = range.format.borders.getItem(edge);
        b.style = "Continuous";
        b.weight = "Thin";
      }
    }
    await context.sync();
    if (number_format !== undefined) range.load("numberFormat");
    if (
      bold !== undefined ||
      italic !== undefined ||
      font_size !== undefined ||
      font_name !== undefined ||
      font_color !== undefined
    )
      range.format.font.load("bold, italic, size, name, color");
    if (fill_color !== undefined) range.format.fill.load("color");
    const topBorder = border ? range.format.borders.getItem("EdgeTop") : null;
    if (topBorder) topBorder.load("style");
    await context.sync();
    if (number_format !== undefined)
      requireReadback(
        range.numberFormat?.flat?.().every((value) => sameCell(value, number_format)),
        "number format did not read back.",
      );
    if (bold !== undefined)
      requireReadback(range.format.font.bold === bold, "bold did not read back.");
    if (italic !== undefined)
      requireReadback(range.format.font.italic === italic, "italic did not read back.");
    if (font_size !== undefined)
      requireReadback(range.format.font.size === font_size, "font size did not read back.");
    if (font_name !== undefined)
      requireReadback(range.format.font.name === font_name, "font name did not read back.");
    if (font_color !== undefined)
      requireReadback(range.format.font.color === font_color, "font color did not read back.");
    if (fill_color !== undefined)
      requireReadback(range.format.fill.color === fill_color, "fill color did not read back.");
    if (topBorder)
      requireReadback(topBorder.style === "Continuous", "border style did not read back.");
    return {
      sheet: sheetName,
      address: range.address,
      formatted: true,
      verification: "format_readback",
    };
  });
}

// Column-letter range helper, e.g. ("C", 2) → "C:D".
function _colSpan(at, count) {
  const start = String(at)
    .toUpperCase()
    .replace(/[^A-Z]/g, "");
  if (!start) throw new Error("`at` must be a column letter, e.g. 'C'.");
  let n = 0;
  for (const ch of start) n = n * 26 + (ch.charCodeAt(0) - 64);
  return `${start}:${_colNumberToLetters(n + count - 1)}`;
}

// Tool: excel_insert_columns — insert blank columns, shifting right.
export async function toolExcelInsertColumns({ sheet = null, at, count = 1 }, control) {
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    armExcelWrite(control);
    const inserted = ws.getRange(_colSpan(at, count)).insert(Excel.InsertShiftDirection.right);
    inserted.load("columnCount");
    await context.sync();
    requireReadback(inserted.columnCount === count, "inserted column count did not read back.");
    return {
      sheet: activeName,
      inserted_at: String(at).toUpperCase(),
      count,
      verification: "column_count_readback",
    };
  });
}

// Tool: excel_delete_columns — delete columns, shifting left.
export async function toolExcelDeleteColumns({ sheet = null, at, count = 1 }, control) {
  if (!Number.isInteger(count) || count < 1) throw new Error("`count` must be a positive integer.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    const span = _colSpan(at, count);
    const start = span.split(":")[0];
    let index = 0;
    for (const ch of start) index = index * 26 + (ch.charCodeAt(0) - 64);
    const neighbor = ws.getRange(
      `${_colNumberToLetters(index + count)}:${_colNumberToLetters(index + count)}`,
    );
    neighbor.load("values");
    await context.sync();
    const expected = neighbor.values;
    armExcelWrite(control);
    ws.getRange(span).delete(Excel.DeleteShiftDirection.left);
    await context.sync();
    const now = ws.getRange(`${start}:${start}`);
    now.load("values");
    await context.sync();
    requireReadback(
      sameGrid(now.values, expected),
      "column to the right did not shift into place.",
    );
    return {
      sheet: activeName,
      deleted_at: String(at).toUpperCase(),
      count,
      verification: "shifted_column_readback",
    };
  });
}

// Tool: excel_add_sheet — add a worksheet (optionally at a 0-based position).
export async function toolExcelAddSheet({ name, position = null }, control) {
  if (!name || typeof name !== "string") throw new Error("`name` is required.");
  return await Excel.run(async (context) => {
    armExcelWrite(control);
    const ws = context.workbook.worksheets.add(name);
    if (Number.isInteger(position)) ws.position = position;
    ws.load("name, position");
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    requireReadback(
      ws.name === name && sheets.items.some((item) => item.name === name),
      "added sheet did not read back.",
    );
    return {
      name: ws.name,
      position: ws.position,
      added: true,
      verification: "sheet_list_readback",
    };
  });
}

// Tool: excel_delete_sheet — delete a worksheet by name.
export async function toolExcelDeleteSheet({ name }, control) {
  if (!name || typeof name !== "string") throw new Error("`name` is required.");
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(name);
    armExcelWrite(control);
    ws.delete();
    await context.sync();
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    requireReadback(
      !sheets.items.some((item) => item.name === name),
      "deleted sheet is still present.",
    );
    return { name, deleted: true, verification: "sheet_list_readback" };
  });
}

// Tool: excel_rename_sheet — rename a worksheet.
export async function toolExcelRenameSheet({ name, new_name }, control) {
  if (!name || !new_name) throw new Error("`name` and `new_name` are required.");
  return await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getItem(name);
    armExcelWrite(control);
    ws.name = new_name;
    await context.sync();
    ws.load("name");
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    requireReadback(
      ws.name === new_name &&
        sheets.items.some((item) => item.name === new_name) &&
        !sheets.items.some((item) => item.name === name),
      "renamed sheet did not read back.",
    );
    return { from: name, to: new_name, renamed: true, verification: "sheet_list_readback" };
  });
}

// Tool: excel_clear_range — clear contents, formats, or both.
export async function toolExcelClearRange({ address, sheet = null, what = "contents" }, control) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  const APPLY = { contents: "Contents", formats: "Formats", all: "All" };
  if (!APPLY[what]) throw new Error(`\`what\` must be one of: ${Object.keys(APPLY).join(", ")}`);
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    armExcelWrite(control);
    range.clear(Excel.ClearApplyTo[what]);
    range.load("address");
    if (what === "contents" || what === "all") range.load("values");
    await context.sync();
    if (
      (what === "contents" || what === "all") &&
      range.values.flat().some((value) => value != null && value !== "")
    )
      throw new Error("RESULT_UNKNOWN: cleared cells still have values.");
    return {
      sheet: sheetName,
      address: range.address,
      cleared: what,
      ...(what === "formats" ? {} : { verification: "values_readback" }),
    };
  });
}

// ---------------------------------------------------------------------------
// Tier 2 editing tools — sort / filter / tables / charts / dimensions
// ---------------------------------------------------------------------------

// Tool: excel_sort_range — sort a range by one column.
export async function toolExcelSortRange(
  { address, sheet = null, key = 0, ascending = true, has_headers = false },
  control,
) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  if (!Number.isInteger(key) || key < 0) throw new Error("`key` must be a 0-based column index.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("values, formulasR1C1, address");
    await context.sync();
    const before = sortRowKeys(range.values, range.formulasR1C1);
    if (!range.values[0] || key >= range.values[0].length)
      throw new Error("INVALID_SORT_KEY: key must identify a column in the target range.");
    armExcelWrite(control);
    range.sort.apply([{ key, ascending: !!ascending }], false, !!has_headers);
    await context.sync();
    range.load("values, formulasR1C1");
    await context.sync();
    let after;
    try {
      after = sortRowKeys(range.values, range.formulasR1C1);
    } catch (error) {
      throw new Error("RESULT_UNKNOWN: " + error.message);
    }
    requireReadback(
      sameRowMultiset(before, after) &&
        (!has_headers || before[0] === after[0]) &&
        columnOrdered(range.values, key, !!ascending, !!has_headers),
      "sorted rows, types, relative formulas or headers did not read back as expected.",
    );
    return {
      sheet: sheetName,
      address: range.address,
      sorted_by_column: key,
      ascending: !!ascending,
      verification: "sort_readback",
      verification_scope: "row_values_and_relative_formulas",
    };
  });
}

// Tool: excel_autofilter — apply an AutoFilter to a range, or clear it.
export async function toolExcelAutoFilter({ address, sheet = null, clear = false }, control) {
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    if (clear) {
      armExcelWrite(control);
      ws.autoFilter.remove();
      await context.sync();
      ws.autoFilter.load("enabled");
      await context.sync();
      requireReadback(ws.autoFilter.enabled === false, "filter is still enabled after clear.");
      return { sheet: activeName, autofilter: "removed", verification: "filter_readback" };
    }
    if (!address || typeof address !== "string") {
      throw new Error("`address` is required (or pass clear: true to remove the filter).");
    }
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const target = context.workbook.worksheets.getItem(sheetName);
    armExcelWrite(control);
    const appliedRange = target.getRange(a1);
    target.autoFilter.apply(appliedRange);
    await context.sync();
    target.autoFilter.load("enabled");
    const filterRange = target.autoFilter.getRange();
    filterRange.load("address");
    await context.sync();
    requireReadback(
      target.autoFilter.enabled === true && addressTail(filterRange.address) === addressTail(a1),
      "filter range did not read back.",
    );
    return {
      sheet: sheetName,
      autofilter: "applied",
      address: a1,
      verification: "filter_readback",
    };
  });
}

// Tool: excel_create_table — turn a range into a named Excel table.
export async function toolExcelCreateTable(
  { address, sheet = null, has_headers = true, name = null },
  control,
) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    armExcelWrite(control);
    const table = ws.tables.add(a1, !!has_headers);
    if (name) table.name = name;
    table.load("name");
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();
    requireReadback(
      !!table.name &&
        (!name || table.name === name) &&
        tables.items.some((item) => item.name === table.name),
      "created table did not read back.",
    );
    return {
      sheet: sheetName,
      table: table.name,
      range: a1,
      verification: "table_readback",
    };
  });
}

// Tool: excel_add_table_rows — append rows to an existing table.
export async function toolExcelAddTableRows({ table, values, index = null }, control) {
  if (!table || typeof table !== "string") throw new Error("`table` (table name) is required.");
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0])) {
    throw new Error("`values` must be a non-empty 2D array.");
  }
  return await Excel.run(async (context) => {
    const t = context.workbook.tables.getItem(table);
    t.rows.load("count");
    await context.sync();
    const before = t.rows.count;
    armExcelWrite(control);
    t.rows.add(Number.isInteger(index) ? index : null, values);
    await context.sync();
    t.rows.load("count");
    await context.sync();
    requireReadback(t.rows.count === before + values.length, "added table rows did not read back.");
    return {
      table,
      added_rows: values.length,
      verification: "table_row_count_readback",
    };
  });
}

const CHART_TYPE = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  scatter: "XYScatter",
  area: "Area",
};

// Tool: excel_create_chart — add a chart from a data range.
export async function toolExcelCreateChart(
  { data_address, sheet = null, chart_type = "column", title = null },
  control,
) {
  if (!data_address || typeof data_address !== "string") {
    throw new Error("`data_address` is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(data_address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const type = CHART_TYPE[String(chart_type).toLowerCase()] || chart_type;
    armExcelWrite(control);
    const chart = ws.charts.add(type, ws.getRange(a1), Excel.ChartSeriesBy.auto);
    if (title) {
      chart.title.text = title;
      chart.title.visible = true;
    }
    chart.load("name");
    if (title) chart.title.load("text, visible");
    const charts = ws.charts;
    charts.load("items/name");
    await context.sync();
    requireReadback(
      !!chart.name &&
        charts.items.some((item) => item.name === chart.name) &&
        (!title || (chart.title.visible === true && chart.title.text === title)),
      "created chart did not read back.",
    );
    return {
      sheet: sheetName,
      chart: chart.name,
      type,
      data: a1,
      verification: "chart_readback",
    };
  });
}

// Tool: excel_set_column_width — set a fixed width, or autofit.
export async function toolExcelSetColumnWidth(
  { address, sheet = null, width = null, autofit = false },
  control,
) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const range = context.workbook.worksheets.getItem(sheetName).getRange(a1);
    if (!(autofit || typeof width === "number"))
      throw new Error("Provide `width` (points) or `autofit: true`.");
    armExcelWrite(control);
    if (autofit) range.format.autofitColumns();
    else range.format.columnWidth = width;
    range.format.load("columnWidth");
    await context.sync();
    requireReadback(
      typeof range.format.columnWidth === "number" &&
        (autofit || range.format.columnWidth === width),
      "column width did not read back.",
    );
    return {
      sheet: sheetName,
      address: a1,
      autofit: !!autofit,
      width: range.format.columnWidth,
      verification: "dimension_readback",
    };
  });
}

// Tool: excel_set_row_height — set a fixed height, or autofit.
export async function toolExcelSetRowHeight(
  { address, sheet = null, height = null, autofit = false },
  control,
) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const range = context.workbook.worksheets.getItem(sheetName).getRange(a1);
    if (!(autofit || typeof height === "number"))
      throw new Error("Provide `height` (points) or `autofit: true`.");
    armExcelWrite(control);
    if (autofit) range.format.autofitRows();
    else range.format.rowHeight = height;
    range.format.load("rowHeight");
    await context.sync();
    requireReadback(
      typeof range.format.rowHeight === "number" && (autofit || range.format.rowHeight === height),
      "row height did not read back.",
    );
    return {
      sheet: sheetName,
      address: a1,
      autofit: !!autofit,
      height: range.format.rowHeight,
      verification: "dimension_readback",
    };
  });
}

// ---------------------------------------------------------------------------
// Write preconditions. Captured before the confirmation card and again after
// approval: the target sheet is pinned by name and worksheet id (the active
// sheet may change while the card is open) and the target's values/formulas
// must be unchanged. Formatting is not compared. Ranges above one snapshot
// are compared in row chunks. Ranges too large to recheck are refused.
// ---------------------------------------------------------------------------
export const MAX_SNAPSHOT_CELLS = 20000;
export const MAX_VERIFIED_CELLS = 100000;
export const MAX_PREVIEW_CELLS = 200;

async function loadFormulas(context, used) {
  const cells = used.rowCount * used.columnCount;
  if (cells > MAX_VERIFIED_CELLS) {
    throw new Error(
      `TARGET_TOO_LARGE: ${cells} cells exceed the ${MAX_VERIFIED_CELLS} that can be rechecked after approval. Narrow the target.`,
    );
  }
  if (!used.columnCount || cells <= MAX_SNAPSHOT_CELLS) {
    used.load("formulas");
    await context.sync();
    return { formulas: used.formulas, chunks: 1 };
  }
  const rowsPerChunk = Math.max(1, Math.floor(MAX_SNAPSHOT_CELLS / used.columnCount));
  const parts = [];
  for (let row = 0; row < used.rowCount; row += rowsPerChunk) {
    const take = Math.min(rowsPerChunk, used.rowCount - row);
    const part = used.getCell(row, 0).getResizedRange(take - 1, used.columnCount - 1);
    part.load("formulas");
    parts.push(part);
  }
  await context.sync();
  return { formulas: parts.flatMap((part) => part.formulas), chunks: parts.length };
}

async function _rangeState(context, range, worksheet) {
  worksheet?.load("id,name");
  const used =
    typeof range.getUsedRangeOrNullObject === "function"
      ? range.getUsedRangeOrNullObject(true)
      : range;
  range.load?.("address");
  used.load("isNullObject, address, rowCount, columnCount");
  await context.sync();
  const identity = worksheet ? { sheet_id: worksheet.id, sheet: worksheet.name } : {};
  if (used.isNullObject)
    return {
      address: range.address ?? null,
      used: null,
      ...identity,
      current_contents: { empty: true, values: [], formulas: [], shown_cells: 0, truncated: false },
    };
  const loaded = await loadFormulas(context, used);
  const columns = Math.min(used.columnCount, MAX_PREVIEW_CELLS);
  const rows = Math.min(used.rowCount, Math.max(1, Math.floor(MAX_PREVIEW_CELLS / columns)));
  const previewRange = used.getCell(0, 0).getResizedRange(rows - 1, columns - 1);
  previewRange.load("address,values");
  await context.sync();
  return {
    address: range.address,
    used: used.address,
    formulas: loaded.formulas,
    chunks: loaded.chunks,
    current_contents: {
      address: previewRange.address,
      values: previewRange.values.map((row) => row.slice()),
      formulas: loaded.formulas.slice(0, rows).map((row) => row.slice(0, columns)),
      shown_cells: rows * columns,
      used_cell_count: used.rowCount * used.columnCount,
      truncated: rows * columns < used.rowCount * used.columnCount,
    },
    ...identity,
  };
}

function excelRecheck(state) {
  if (state?.formulas) return state.chunks > 1 ? "formulas_chunked" : "formulas";
  if (state) return "identity";
  return "none";
}

export async function captureExcelPrecondition(name, args = {}) {
  return await Excel.run(async (context) => {
    const pinned = { ...args };
    const target = async (address) => {
      const active = args.sheet || (await _activeSheetName(context));
      const { sheetName, a1 } = _splitSheetAddress(address, active);
      pinned.sheet = sheetName;
      return { ws: context.workbook.worksheets.getItem(sheetName), a1 };
    };
    const span = (count) => (Number.isInteger(count) && count > 0 ? count : 1);
    let state = null;
    switch (name) {
      case "excel_write_range":
      case "excel_write_formula":
      case "excel_clear_range":
      case "excel_sort_range":
      case "excel_create_table":
      case "excel_set_format":
      case "excel_set_column_width":
      case "excel_set_row_height":
      case "excel_create_chart": {
        const { ws, a1 } = await target(
          name === "excel_create_chart" ? args.data_address : args.address,
        );
        if (!a1) throw new Error("`address` is required.");
        state = await _rangeState(context, ws.getRange(a1), ws);
        break;
      }
      case "excel_autofilter": {
        const { ws, a1 } = await target(args.address);
        state =
          !args.clear && a1
            ? await _rangeState(context, ws.getRange(a1), ws)
            : await _rangeState(context, ws.getUsedRangeOrNullObject(true), ws);
        break;
      }
      case "excel_delete_rows": {
        const { ws } = await target(null);
        if (!Number.isInteger(args.at) || args.at < 1)
          throw new Error("`at` must be a positive integer.");
        state = await _rangeState(
          context,
          ws.getRange(`${args.at}:${args.at + span(args.count) - 1}`),
          ws,
        );
        break;
      }
      case "excel_delete_columns": {
        const { ws } = await target(null);
        state = await _rangeState(context, ws.getRange(_colSpan(args.at, span(args.count))), ws);
        break;
      }
      case "excel_insert_rows":
      case "excel_insert_columns": {
        const { ws } = await target(null);
        state = await _rangeState(context, ws.getUsedRangeOrNullObject(true), ws);
        break;
      }
      case "excel_delete_sheet": {
        const ws = context.workbook.worksheets.getItem(args.name);
        state = await _rangeState(context, ws.getRange(), ws);
        break;
      }
      case "excel_add_sheet":
      case "excel_rename_sheet": {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name, items/id");
        await context.sync();
        state = {
          sheets: sheets.items.map((ws) => ws.name),
          sheet_ids: sheets.items.map((ws) => ws.id),
        };
        break;
      }
      case "excel_add_table_rows":
        state = await _rangeState(context, context.workbook.tables.getItem(args.table).getRange());
        break;
    }
    // Display values are bounded preview data, not additional live preconditions.
    // Keep the existing full/chunked formula and sheet-identity recheck unchanged.
    const currentContents = state?.current_contents;
    if (state) delete state.current_contents;
    return {
      args: pinned,
      state,
      preview: {
        ...(pinned.sheet ? { sheet: pinned.sheet } : {}),
        ...(state?.address ? { target_address: state.address } : {}),
        ...(state && "used" in state ? { target_used_cells: state.used } : {}),
        ...(currentContents ? { current_contents: currentContents } : {}),
        recheck: excelRecheck(state),
      },
    };
  });
}
