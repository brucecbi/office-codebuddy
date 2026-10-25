// Guard for daemon/office-read-only.mjs: the daemon-side list of Office tools
// that never modify the document. The turn-timeout copy and any future
// daemon-side write detection rely on it. The taskpane keeps its own copy
// (readOnlyOfficeTools in taskpane/shared/taskpane.js) — the two MUST stay
// equal; this test is the tripwire.

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { OFFICE_READ_ONLY_TOOLS } from "../daemon/office-read-only.mjs";

const EXPECTED = [
  "office_document_map",
  "office_read_story",
  "office_read_formatting",
  "office_read_revisions",
  "office_audit_redline",
  "office_get_selection",
  "office_read_paragraphs",
  "office_get_document_text",
  "office_get_outline",
  "office_find",
  "office_list_comments",
  "excel_get_selected_range",
  "excel_list_sheets",
  "excel_read_range",
  "excel_find_value",
  "excel_select_range",
];

test("read-only office tool set matches the taskpane set exactly", () => {
  assert.deepEqual([...OFFICE_READ_ONLY_TOOLS].sort(), EXPECTED.sort());
  const source = readFileSync(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const match = source.match(/const readOnlyOfficeTools = new Set\(\[([\s\S]*?)\]\);/);
  assert.ok(match, "taskpane read-only set was not found");
  const pane = [...match[1].matchAll(/"([^"]+)"/g)].map((item) => item[1]);
  assert.deepEqual(pane.sort(), [...OFFICE_READ_ONLY_TOOLS].sort());
});

test("read-only names all use the office_/excel_ tool namespace", () => {
  for (const name of OFFICE_READ_ONLY_TOOLS) {
    assert.match(name, /^(office|excel)_[a-z_]+$/);
  }
});
