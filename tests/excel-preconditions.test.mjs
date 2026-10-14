import { test } from "node:test";
import assert from "node:assert/strict";
import {
  captureExcelPrecondition,
  toolExcelAddSheet,
  toolExcelAddTableRows,
  toolExcelAutoFilter,
  toolExcelClearRange,
  toolExcelCreateChart,
  toolExcelCreateTable,
  toolExcelDeleteColumns,
  toolExcelDeleteRows,
  toolExcelDeleteSheet,
  toolExcelInsertColumns,
  toolExcelInsertRows,
  toolExcelRenameSheet,
  toolExcelSetColumnWidth,
  toolExcelSetFormat,
  toolExcelSortRange,
  toolExcelWriteRange,
} from "../taskpane/shared/tools-excel.js";

// Minimal synthetic workbook: sparse cells per sheet, A1 ranges, used ranges.
const col = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const letters = (n) => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
function bounds(a1) {
  const [a, b = a] = a1.split(":");
  const part = (ref, end) => {
    const m = /^([A-Z]*)(\d*)$/.exec(ref);
    return {
      c: m[1] ? col(m[1]) : end ? 26 : 1,
      r: m[2] ? Number(m[2]) : end ? 2000 : 1,
    };
  };
  const s = part(a, false),
    e = part(b, true);
  return { r1: s.r, c1: s.c, r2: e.r, c2: e.c };
}
function workbook(sheets, active = Object.keys(sheets)[0], calculated = {}) {
  const state = { active, filters: [] };
  const identities = new WeakMap();
  let nextId = 1;
  const identity = (data) => {
    if (!identities.has(data)) identities.set(data, `sheet-${nextId++}`);
    return identities.get(data);
  };
  const grid = (name, box) =>
    Array.from({ length: box.r2 - box.r1 + 1 }, (_, i) =>
      Array.from(
        { length: box.c2 - box.c1 + 1 },
        (_, j) => sheets[name][`${letters(box.c1 + j)}${box.r1 + i}`] ?? "",
      ),
    );
  const rangeFor = (name, box) => {
    const addr = `${name}!${letters(box.c1)}${box.r1}:${letters(box.c2)}${box.r2}`;
    const self = {
      address: addr,
      isNullObject: false,
      rowCount: box.r2 - box.r1 + 1,
      columnCount: box.c2 - box.c1 + 1,
      load() {},
      get formulas() {
        return grid(name, box);
      },
      get formulasR1C1() {
        return grid(name, box);
      },
      get values() {
        return grid(name, box).map((row, i) =>
          row.map((value, j) => calculated[name]?.[`${letters(box.c1 + j)}${box.r1 + i}`] ?? value),
        );
      },
      set values(rows) {
        rows.forEach((row, i) =>
          row.forEach((value, j) => {
            const ref = `${letters(box.c1 + j)}${box.r1 + i}`;
            if (value == null || value === "") delete sheets[name][ref];
            else sheets[name][ref] = value;
          }),
        );
      },
      set formulas(rows) {
        this.values = rows;
      },
      clear() {
        for (const ref of Object.keys(sheets[name])) {
          const hit = bounds(ref);
          if (hit.r1 >= box.r1 && hit.r1 <= box.r2 && hit.c1 >= box.c1 && hit.c1 <= box.c2)
            delete sheets[name][ref];
        }
      },
      numberFormat: null,
      format: (() => {
        const borders = {};
        return {
          font: { bold: null, italic: null, size: null, name: null, color: null, load() {} },
          fill: { color: null, load() {} },
          borders: {
            getItem(edge) {
              if (!borders[edge]) borders[edge] = { style: null, weight: null, load() {} };
              return borders[edge];
            },
          },
          columnWidth: null,
          rowHeight: null,
          load() {},
          autofitColumns() {
            this.columnWidth = 64;
          },
          autofitRows() {
            this.rowHeight = 15;
          },
        };
      })(),
      sort: {
        apply(fields, _matchCase, hasHeaders) {
          const key = fields[0].key;
          const ascending = fields[0].ascending !== false;
          const rows = [];
          for (let r = box.r1; r <= box.r2; r++) {
            const row = [];
            for (let c = box.c1; c <= box.c2; c++)
              row.push(sheets[name][`${letters(c)}${r}`] ?? "");
            rows.push(row);
          }
          const head = hasHeaders ? rows.slice(0, 1) : [];
          const body = (hasHeaders ? rows.slice(1) : rows).slice();
          body.sort((a, b) => {
            const cmp = String(a[key] ?? "").localeCompare(String(b[key] ?? ""), undefined, {
              numeric: true,
            });
            return ascending ? cmp : -cmp;
          });
          head.concat(body).forEach((row, i) =>
            row.forEach((value, j) => {
              const ref = `${letters(box.c1 + j)}${box.r1 + i}`;
              if (value == null || value === "") delete sheets[name][ref];
              else sheets[name][ref] = value;
            }),
          );
        },
      },
      insert(direction) {
        const rowCount = box.r2 - box.r1 + 1;
        const columnCount = box.c2 - box.c1 + 1;
        const entries = Object.entries(sheets[name]).filter(([ref]) => /^[A-Z]+[0-9]+$/.test(ref));
        for (const ref of Object.keys(sheets[name])) {
          if (/^[A-Z]+[0-9]+$/.test(ref)) delete sheets[name][ref];
        }
        for (const [ref, value] of entries) {
          const hit = bounds(ref);
          const down = direction === "Down" || direction === Excel.InsertShiftDirection?.down;
          const row = down && hit.r1 >= box.r1 ? hit.r1 + rowCount : hit.r1;
          const column = !down && hit.c1 >= box.c1 ? hit.c1 + columnCount : hit.c1;
          sheets[name][`${letters(column)}${row}`] = value;
        }
        return rangeFor(name, box);
      },
      delete(direction) {
        const rowCount = box.r2 - box.r1 + 1;
        const columnCount = box.c2 - box.c1 + 1;
        const up = direction === "Up" || direction === Excel.DeleteShiftDirection?.up;
        const entries = Object.entries(sheets[name]).filter(([ref]) => /^[A-Z]+[0-9]+$/.test(ref));
        for (const ref of Object.keys(sheets[name])) {
          if (/^[A-Z]+[0-9]+$/.test(ref)) delete sheets[name][ref];
        }
        for (const [ref, value] of entries) {
          const hit = bounds(ref);
          if (up && hit.r1 >= box.r1 && hit.r1 <= box.r2) continue;
          if (!up && hit.c1 >= box.c1 && hit.c1 <= box.c2) continue;
          const row = up && hit.r1 > box.r2 ? hit.r1 - rowCount : hit.r1;
          const column = !up && hit.c1 > box.c2 ? hit.c1 - columnCount : hit.c1;
          sheets[name][`${letters(column)}${row}`] = value;
        }
      },
      getCell(row, column) {
        return rangeFor(name, {
          r1: box.r1 + row,
          c1: box.c1 + column,
          r2: box.r1 + row,
          c2: box.c1 + column,
        });
      },
      getResizedRange(deltaRows, deltaColumns) {
        return rangeFor(name, {
          r1: box.r1,
          c1: box.c1,
          r2: box.r2 + deltaRows,
          c2: box.c2 + deltaColumns,
        });
      },
      getUsedRangeOrNullObject() {
        const hits = Object.keys(sheets[name])
          .filter((ref) => /^[A-Z]+[0-9]+$/.test(ref))
          .map((ref) => bounds(ref))
          .filter((h) => h.r1 >= box.r1 && h.r1 <= box.r2 && h.c1 >= box.c1 && h.c1 <= box.c2);
        if (!hits.length) return { isNullObject: true, load() {} };
        return rangeFor(name, {
          r1: Math.min(...hits.map((h) => h.r1)),
          r2: Math.max(...hits.map((h) => h.r1)),
          c1: Math.min(...hits.map((h) => h.c1)),
          c2: Math.max(...hits.map((h) => h.c1)),
        });
      },
    };
    return self;
  };
  const rangeOn = (name, a1) => rangeFor(name, bounds(a1));
  const charts = [];
  const tables = [];
  const sheet = (sheetName) => {
    if (!sheets[sheetName]) throw new Error(`ItemNotFound: ${sheetName}`);
    const record = { key: sheetName };
    return {
      get name() {
        return record.key;
      },
      set name(next) {
        if (!next || next === record.key) return;
        sheets[next] = sheets[record.key];
        delete sheets[record.key];
        if (state.active === record.key) state.active = next;
        for (const chart of charts) if (chart.sheet === record.key) chart.sheet = next;
        for (const table of tables) if (table.sheet === record.key) table.sheet = next;
        record.key = next;
      },
      id: identity(sheets[sheetName]),
      position: Object.keys(sheets).indexOf(sheetName),
      load() {},
      delete() {
        delete sheets[record.key];
      },
      getRange: (a1 = "A1:Z2000") => rangeOn(record.key, a1),
      getUsedRangeOrNullObject: () => rangeOn(record.key, "A1:XFD2000").getUsedRangeOrNullObject(),
      autoFilter: {
        enabled: false,
        address: null,
        load() {},
        getRange() {
          return { address: this.address, load() {} };
        },
        apply: (range) => {
          const filter = sheet(record.key).autoFilter;
          void filter;
          state.filters.push(range.address);
        },
        remove() {},
      },
      charts: {
        load() {},
        get items() {
          return charts.filter((chart) => chart.sheet === record.key);
        },
        add(type) {
          const chart = {
            name: `Chart${charts.length + 1}`,
            type,
            sheet: record.key,
            title: { text: "", visible: false, load() {} },
            load() {},
          };
          charts.push(chart);
          return chart;
        },
      },
      tables: {
        add(a1, hasHeaders) {
          const table = {
            name: `Table${tables.length + 1}`,
            sheet: record.key,
            rows: {
              count: hasHeaders ? 1 : 0,
              load() {},
              add(_index, values) {
                this.count += values.length;
              },
            },
            load() {},
            getRange: () => rangeOn(record.key, a1),
          };
          tables.push(table);
          return table;
        },
      },
    };
  };
  // Filter state has to survive on the worksheet instance the tool keeps.
  const filters = new Map();
  const sheetWithFilter = (sheetName) => {
    const ws = sheet(sheetName);
    if (!filters.has(ws.id)) {
      const filter = {
        enabled: false,
        address: null,
        load() {},
        getRange() {
          return { address: this.address, load() {} };
        },
        apply(range) {
          this.enabled = true;
          this.address = range.address;
          state.filters.push(range.address);
        },
        remove() {
          this.enabled = false;
          this.address = null;
        },
      };
      filters.set(ws.id, filter);
    }
    ws.autoFilter = filters.get(ws.id);
    return ws;
  };
  globalThis.Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats", all: "All" },
    InsertShiftDirection: { down: "Down", right: "Right" },
    DeleteShiftDirection: { up: "Up", left: "Left" },
    ChartSeriesBy: { auto: "Auto" },
    run: (fn) =>
      fn({
        sync: async () => {},
        workbook: {
          worksheets: {
            getActiveWorksheet: () => sheetWithFilter(state.active),
            getItem: sheetWithFilter,
            add(sheetName) {
              sheets[sheetName] = {};
              const created = sheetWithFilter(sheetName);
              created.position = Object.keys(sheets).length - 1;
              return created;
            },
            load() {},
            get items() {
              return Object.keys(sheets).map((name) => ({ name, id: identity(sheets[name]) }));
            },
          },
          tables: {
            load() {},
            get items() {
              return tables.map((table) => ({ name: table.name }));
            },
            getItem: (tableName) =>
              tables.find((table) => table.name === tableName) || {
                name: tableName,
                getRange: () => rangeOn("Sheet1", "A1:B3"),
                rows: {
                  count: 0,
                  load() {},
                  add(_index, values) {
                    this.count += values.length;
                  },
                },
                load() {},
              },
          },
        },
      }),
  };
  return state;
}

test("approval preview includes bounded old values and formulas, not just an address", async () => {
  const calculated = { Sheet1: { B1: 6 } };
  workbook({ Sheet1: { A1: "old value", B1: "=2*3" } }, "Sheet1", calculated);
  const pre = await captureExcelPrecondition("excel_write_range", {
    address: "A1:B1",
    values: [["new", 9]],
  });
  assert.deepEqual(pre.preview.current_contents.values, [["old value", 6]]);
  assert.deepEqual(pre.preview.current_contents.formulas, [["old value", "=2*3"]]);
  assert.equal(pre.preview.current_contents.shown_cells, 2);
  assert.equal(pre.preview.current_contents.truncated, false);
  assert.equal(pre.state.current_contents, undefined);
  calculated.Sheet1.B1 = 7;
  assert.deepEqual(
    (await captureExcelPrecondition("excel_write_range", pre.args)).state,
    pre.state,
  );

  workbook({ Sheet1: {} });
  const empty = await captureExcelPrecondition("excel_write_range", {
    address: "A1",
    values: [[1]],
  });
  assert.equal(empty.preview.current_contents.empty, true);
  assert.deepEqual(empty.preview.current_contents.values, []);
});

test("approval preview caps tall and wide ranges at 200 cells without truncating preconditions", async () => {
  for (const cells of [
    Object.fromEntries(Array.from({ length: 201 }, (_, i) => [`A${i + 1}`, `old-${i}`])),
    Object.fromEntries(Array.from({ length: 201 }, (_, i) => [`${letters(i + 1)}1`, `old-${i}`])),
  ]) {
    workbook({ Sheet1: cells });
    const address = Object.keys(cells).at(-1);
    const pre = await captureExcelPrecondition("excel_clear_range", { address: `A1:${address}` });
    assert.equal(pre.preview.current_contents.shown_cells, 200);
    assert.equal(pre.preview.current_contents.truncated, true);
    assert.equal(pre.preview.current_contents.values.flat().length, 200);
    assert.equal(pre.preview.current_contents.formulas.flat().length, 200);
    assert.equal(pre.state.formulas.flat().length, 201);
  }
});

test("the active sheet is pinned before confirmation", async () => {
  const wb = workbook({ Sheet1: { A1: "x" }, Sheet2: {} });
  const pre = await captureExcelPrecondition("excel_write_range", {
    address: "A1",
    values: [["y"]],
  });
  assert.equal(pre.args.sheet, "Sheet1");
  assert.equal(pre.preview.sheet, "Sheet1");
  wb.active = "Sheet2";
  const later = await captureExcelPrecondition("excel_write_range", pre.args);
  assert.equal(later.args.sheet, "Sheet1");
  assert.deepEqual(later.state, pre.state);
});

test("edits inside the target change the state; edits elsewhere do not", async () => {
  const cells = { A1: "1", B2: "=A1*2", D9: "other" };
  workbook({ Sheet1: cells });
  const pre = await captureExcelPrecondition("excel_clear_range", { address: "A1:B2" });
  assert.deepEqual(pre.state.formulas, [
    ["1", ""],
    ["", "=A1*2"],
  ]);
  cells.D9 = "changed elsewhere";
  assert.deepEqual(
    (await captureExcelPrecondition("excel_clear_range", pre.args)).state,
    pre.state,
  );
  cells.A2 = "typed by user";
  assert.notDeepEqual(
    (await captureExcelPrecondition("excel_clear_range", pre.args)).state,
    pre.state,
  );
});

test("row deletion and sheet operations snapshot what they destroy or rename", async () => {
  const cells = { A3: "keep?", C4: "x" };
  workbook({ Sheet1: cells, Other: {} });
  const rows = await captureExcelPrecondition("excel_delete_rows", { at: 3, count: 2 });
  assert.equal(rows.args.sheet, "Sheet1");
  assert.equal(rows.state.used, "Sheet1!A3:C4");
  const names = await captureExcelPrecondition("excel_rename_sheet", {
    name: "Other",
    new_name: "N",
  });
  assert.deepEqual(names.state.sheets, ["Sheet1", "Other"]);
  assert.equal(names.state.sheet_ids.length, 2);
  assert.notEqual(names.state.sheet_ids[0], names.state.sheet_ids[1]);
  await assert.rejects(
    captureExcelPrecondition("excel_delete_rows", { at: 0 }),
    /positive integer/,
  );
});

test("large targets are rechecked by chunked formulas, not by shape alone", async () => {
  const cells = { A1: "top", Z1000: "bottom" };
  workbook({ Sheet1: cells });
  const pre = await captureExcelPrecondition("excel_sort_range", { address: "A:Z" });
  assert.equal(pre.preview.recheck, "formulas_chunked");
  assert.equal(pre.state.formulas[0][0], "top");
  assert.equal(pre.state.formulas.at(-1).at(-1), "bottom");
  cells.A1 = "changed";
  assert.notDeepEqual(
    (await captureExcelPrecondition("excel_sort_range", pre.args)).state,
    pre.state,
  );
});

test("targets too large to recheck are refused before a proposal", async () => {
  workbook({ Sheet1: { A1: "a", AY2000: "z" } });
  await assert.rejects(
    captureExcelPrecondition("excel_clear_range", { address: "A:AY" }),
    /TARGET_TOO_LARGE/,
  );
});

test("replacing a worksheet keeps the name but changes its id", async () => {
  const sheets = { Sheet1: { A1: "same content" } };
  workbook(sheets);
  const before = await captureExcelPrecondition("excel_write_range", {
    address: "A1",
    values: [["new"]],
  });
  sheets.Sheet1 = { A1: "same content" };
  const after = await captureExcelPrecondition("excel_write_range", before.args);
  assert.notEqual(before.state.sheet_id, after.state.sheet_id);
  assert.notDeepEqual(before.state, after.state);
});

test("row insertion and filter clearing snapshot sheet identity and used cells", async () => {
  const cells = { A1: "keep" };
  workbook({ Sheet1: cells });
  const rows = await captureExcelPrecondition("excel_insert_rows", { at: 2, count: 1 });
  assert.equal(rows.args.sheet, "Sheet1");
  assert.ok(rows.state.sheet_id);
  assert.equal(rows.state.formulas[0][0], "keep");
  cells.A1 = "edited";
  assert.notDeepEqual(
    (await captureExcelPrecondition("excel_insert_rows", rows.args)).state,
    rows.state,
  );
  const cleared = await captureExcelPrecondition("excel_autofilter", {
    clear: true,
    sheet: "Sheet1",
  });
  assert.equal(cleared.state.sheet_id, rows.state.sheet_id);
  assert.equal(cleared.state.formulas[0][0], "edited");
});

test("value writes read back and a cancel during preparation does not write", async () => {
  const cells = { A1: "old" };
  workbook({ Sheet1: cells });
  const wrote = await toolExcelWriteRange(
    { address: "A1", values: [["new"]], sheet: "Sheet1" },
    { guard() {}, beforeWrite() {} },
  );
  assert.equal(wrote.verification, "values_readback");
  assert.deepEqual(wrote.readback, [["new"]]);
  assert.equal(cells.A1, "new");
  await toolExcelClearRange(
    { address: "A1", sheet: "Sheet1", what: "contents" },
    { guard() {}, beforeWrite() {} },
  );
  assert.equal(cells.A1, undefined);

  cells.A1 = "old";
  const original = Excel.run;
  let syncs = 0;
  let cancelled = false;
  Excel.run = (fn) =>
    original((ctx) =>
      fn({
        ...ctx,
        sync: async () => {
          if (++syncs === 1) cancelled = true;
        },
      }),
    );
  await assert.rejects(
    toolExcelWriteRange(
      { address: "A1", values: [["new"]], sheet: "Sheet1" },
      {
        guard() {
          if (cancelled) throw new Error("CANCELLED_OR_EXPIRED");
        },
        beforeWrite() {
          throw new Error("write was marked after cancel");
        },
      },
    ),
    /CANCELLED_OR_EXPIRED/,
  );
  Excel.run = original;
  assert.equal(cells.A1, "old");
});

test("autofilter honors the sheet named in the address", async () => {
  const wb = workbook({ Sheet1: {}, Data: { A1: "h" } });
  const applied = await toolExcelAutoFilter({ address: "Data!A1:C5" });
  assert.equal(applied.verification, "filter_readback");
  assert.deepEqual(wb.filters, ["Data!A1:C5"]);
  const cleared = await toolExcelAutoFilter({ clear: true, sheet: "Data" });
  assert.equal(cleared.verification, "filter_readback");
});

const arm = { guard() {}, beforeWrite() {} };

function sortReadback(before, after, beforeFormulas = before, afterFormulas = after) {
  let sorted = false;
  const range = {
    address: "Sheet1!A1:B3",
    load() {},
    get values() {
      return (sorted ? after : before).map((row) => row.slice());
    },
    get formulasR1C1() {
      return (sorted ? afterFormulas : beforeFormulas).map((row) => row.slice());
    },
    sort: {
      apply() {
        sorted = true;
      },
    },
  };
  globalThis.Excel = {
    run: (fn) =>
      fn({
        sync: async () => {},
        workbook: { worksheets: { getItem: () => ({ getRange: () => range }) } },
      }),
  };
  return () => sorted;
}

test("sort readback rejects lost row associations, changed types and headers", async () => {
  for (const [before, after, has_headers] of [
    [
      [
        [2, 20],
        [1, 10],
      ],
      [
        [1, 20],
        [2, 10],
      ],
      false,
    ],
    [
      [
        [2, 20],
        [1, 10],
      ],
      [
        [1, "10"],
        [2, 20],
      ],
      false,
    ],
    [
      [
        ["key", "value"],
        [2, 20],
        [1, 10],
      ],
      [
        [1, 10],
        [2, 20],
        ["key", "value"],
      ],
      true,
    ],
  ]) {
    sortReadback(before, after);
    await assert.rejects(
      toolExcelSortRange(
        {
          address: "A1:B3",
          sheet: "Sheet1",
          key: 0,
          has_headers,
        },
        arm,
      ),
      /RESULT_UNKNOWN/,
    );
  }
});

test("sort preserves duplicate complete rows and accepts relative-formula recalculation", async () => {
  sortReadback(
    [
      [2, 20],
      [1, 10],
      [1, 10],
    ],
    [
      [1, 10],
      [1, 10],
      [2, 20],
    ],
  );
  assert.equal(
    (await toolExcelSortRange({ address: "A1:B3", sheet: "Sheet1" }, arm)).verification,
    "sort_readback",
  );
  sortReadback(
    [
      [2, 20],
      [1, 10],
    ],
    [
      [1, 11],
      [2, 21],
    ],
    [
      [2, "=RC[-1]*10"],
      [1, "=RC[-1]*10"],
    ],
    [
      [1, "=RC[-1]*10"],
      [2, "=RC[-1]*10"],
    ],
  );
  assert.equal(
    (await toolExcelSortRange({ address: "A1:B2", sheet: "Sheet1" }, arm)).verification,
    "sort_readback",
  );
});

test("sort readback rejects changed relative formulas and invalid keys before writing", async () => {
  sortReadback(
    [
      [2, 20],
      [1, 10],
    ],
    [
      [1, 10],
      [2, 20],
    ],
    [
      [2, "=RC[-1]*10"],
      [1, "=RC[-1]*10"],
    ],
    [
      [1, "=RC[-1]*99"],
      [2, "=RC[-1]*10"],
    ],
  );
  await assert.rejects(
    toolExcelSortRange({ address: "A1:B2", sheet: "Sheet1" }, arm),
    /RESULT_UNKNOWN/,
  );
  const sorted = sortReadback(
    [
      [2, 20],
      [1, 10],
    ],
    [
      [1, 10],
      [2, 20],
    ],
  );
  await assert.rejects(
    toolExcelSortRange({ address: "A1:B2", sheet: "Sheet1", key: 2 }, arm),
    /column|key/i,
  );
  assert.equal(sorted(), false);
});

test("sort refuses unavailable formula inspection and marks failed post-write readback unknown", async () => {
  let sorted = sortReadback([[2, 20]], [[2, 20]], [], [[2, 20]]);
  await assert.rejects(
    toolExcelSortRange({ address: "A1:B1", sheet: "Sheet1" }, arm),
    /READBACK_UNAVAILABLE/,
  );
  assert.equal(sorted(), false);
  sorted = sortReadback([[2, 20]], [[2, 20]], [[2, 20]], []);
  await assert.rejects(
    toolExcelSortRange({ address: "A1:B1", sheet: "Sheet1" }, arm),
    /RESULT_UNKNOWN:.*READBACK_UNAVAILABLE/,
  );
  assert.equal(sorted(), true);
});

test("structural Excel writes read back rows, columns, sort, format and sheets", async () => {
  const cells = { A1: "b", A2: "a", C1: "right" };
  workbook({ Sheet1: cells, Extra: { A1: "gone" } });
  const inserted = await toolExcelInsertRows({ sheet: "Sheet1", at: 1, count: 1 }, arm);
  assert.equal(inserted.verification, "row_count_readback");
  assert.equal(cells.A2, "b");
  assert.equal(cells.A3, "a");
  const deleted = await toolExcelDeleteRows({ sheet: "Sheet1", at: 1, count: 1 }, arm);
  assert.equal(deleted.verification, "shifted_row_readback");
  assert.equal(cells.A1, "b");
  const columns = await toolExcelInsertColumns({ sheet: "Sheet1", at: "B", count: 1 }, arm);
  assert.equal(columns.verification, "column_count_readback");
  assert.equal(cells.D1, "right");
  const removed = await toolExcelDeleteColumns({ sheet: "Sheet1", at: "B", count: 1 }, arm);
  assert.equal(removed.verification, "shifted_column_readback");
  assert.equal(cells.C1, "right");
  const sorted = await toolExcelSortRange(
    { address: "A1:A2", sheet: "Sheet1", key: 0, ascending: true },
    arm,
  );
  assert.equal(sorted.verification, "sort_readback");
  assert.equal(cells.A1, "a");
  assert.equal(cells.A2, "b");
  const formatted = await toolExcelSetFormat(
    { address: "A1", sheet: "Sheet1", bold: true, number_format: "0.00" },
    arm,
  );
  assert.equal(formatted.verification, "format_readback");
  const width = await toolExcelSetColumnWidth({ address: "A:A", sheet: "Sheet1", width: 22 }, arm);
  assert.equal(width.verification, "dimension_readback");
  assert.equal(width.width, 22);
  const added = await toolExcelAddSheet({ name: "New" }, arm);
  assert.equal(added.verification, "sheet_list_readback");
  const renamed = await toolExcelRenameSheet({ name: "New", new_name: "Newer" }, arm);
  assert.equal(renamed.verification, "sheet_list_readback");
  const gone = await toolExcelDeleteSheet({ name: "Extra" }, arm);
  assert.equal(gone.verification, "sheet_list_readback");
  const table = await toolExcelCreateTable(
    { address: "A1:A2", sheet: "Sheet1", has_headers: true, name: "Items" },
    arm,
  );
  assert.equal(table.table, "Items");
  const rows = await toolExcelAddTableRows({ table: "Items", values: [["c"]] }, arm);
  assert.equal(rows.verification, "table_row_count_readback");
  const chart = await toolExcelCreateChart(
    { data_address: "A1:A2", sheet: "Sheet1", title: "Count" },
    arm,
  );
  assert.equal(chart.verification, "chart_readback");
  assert.equal(chart.chart, "Chart1");
});
