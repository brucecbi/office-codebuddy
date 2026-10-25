// Office tools that never modify the document. Used by the daemon to decide
// whether an interrupted turn could have edits in flight (turn-timeout copy).
// The taskpane keeps its own identical list (readOnlyOfficeTools in
// taskpane/shared/taskpane.js) for confirmation-card gating — keep the two
// in sync; tests/office-read-only.test.mjs pins the expected contents.

export const OFFICE_READ_ONLY_TOOLS = new Set([
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
]);
