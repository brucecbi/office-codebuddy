import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toolReadParagraphs,
  toolInsertParagraphs,
  toolReplaceSection,
  toolClearHighlights,
  toolClearComments,
  prepareSafeWordWrite,
} from "../taskpane/shared/tools-word.js";
import { assertTextTargetSafe } from "../taskpane/shared/word-write-safety.js";

function fixture(rows) {
  let writes = 0;
  const xmlOf = (row) =>
    (row.xml ?? `<w:p><w:r><w:t>${row.text}</w:t></w:r></w:p>`).replace(
      "<w:p>",
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    );
  const items = rows.map((row, i) => {
    const paragraph = {
      uniqueLocalId: `P${i}`,
      style: row.heading ? "Heading 1" : "Normal",
      text: row.raw ?? row.text,
      listItemOrNullObject: { isNullObject: true },
      load() {},
      insertText() {
        writes++;
      },
      insertParagraph() {
        writes++;
        return paragraph;
      },
      delete() {
        writes++;
      },
      getRange() {
        return {
          getReviewedText: () => ({ value: row.text }),
          getOoxml: () => ({ value: xmlOf(row) }),
          font: {
            set highlightColor(v) {
              writes++;
            },
          },
        };
      },
    };
    return paragraph;
  });
  const body = {
    paragraphs: { items, load() {} },
    getComments: () => ({ items: [], load() {} }),
    getOoxml: () => ({
      value:
        '<w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        rows.map(xmlOf).join("") +
        "</w:body>",
    }),
  };
  const doc = { body, changeTrackingMode: "Off", load() {}, getSelection: () => ({ select() {} }) };
  globalThis.Word = {
    run: async (fn) => fn({ document: doc, sync: async () => {} }),
    ChangeTrackingVersion: { current: "Current" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    InsertLocation: { replace: "Replace", after: "After" },
    HeaderFooterType: { primary: "Primary" },
  };
  return { writes: () => writes, rows, body, doc, items };
}

test("RED every heading-based path refuses duplicate Current headings before a write", async () => {
  const rows = [
    { text: "Background", heading: true },
    { text: "First body" },
    { text: "Background", heading: true },
    { text: "Second body" },
  ];
  for (const action of [
    () => toolReadParagraphs({ heading_section: "background" }),
    () => toolInsertParagraphs({ after: { heading: "Background" }, content: ["Change"] }),
    () => toolReplaceSection({ heading: "Background", content: ["Change"] }),
    () => toolClearHighlights({ heading_section: "Background" }),
    () => toolClearComments({ heading_section: "Background" }),
  ]) {
    const f = fixture(rows);
    await assert.rejects(
      action(),
      (e) => /AMBIGUOUS_HEADING/.test(e.message) && /P0/.test(e.message) && /P2/.test(e.message),
    );
    assert.equal(f.writes(), 0);
  }
});

test("RED stable ID chooses a duplicate heading and raw tracked text is not the key", async () => {
  fixture([
    { text: "Current Heading", raw: "Old HeadingCurrent Heading", heading: true },
    { text: "First" },
    { text: "Current Heading", heading: true },
    { text: "Second" },
  ]);
  const selected = await toolReadParagraphs({ heading_id: "P2" });
  assert.deepEqual(
    selected.paragraphs.map((p) => p.text),
    ["Current Heading", "Second"],
  );
  await assert.rejects(
    toolReadParagraphs({ heading_section: "Old HeadingCurrent Heading" }),
    /Heading not found/,
  );
});

test("RED missing/duplicate IDs and invalid ranges cannot return misleading partial coverage", async () => {
  const f = fixture([{ text: "A" }, { text: "B" }]);
  for (const args of [
    { ids: ["P0", "missing"] },
    { ids: ["P0", "P0"] },
    { ids: [] },
    { range: [-1, 2] },
    { range: [1, 0] },
    { range: [0, 3] },
    { range: [0, 1.5] },
    { ids: ["P0"], range: [0, 1] },
    { heading_id: "P0" },
  ])
    await assert.rejects(toolReadParagraphs(args), /PARAGRAPH_NOT_FOUND|INVALID_SELECTOR/);
  assert.equal(f.writes(), 0);
});

test("RED broad paragraph writes reject field/review/mixed-run structure before confirmation", async () => {
  const rich = [
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>1</w:t></w:r></w:p>',
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Text</w:t></w:r><w:commentRangeEnd w:id="0"/></w:p>',
    "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t> Plain</w:t></w:r></w:p>",
    '<w:p><w:ins w:id="1"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>',
  ];
  for (const xml of rich) {
    const f = fixture([{ text: "Text", xml }]);
    await assert.rejects(
      prepareSafeWordWrite("office_replace_paragraphs", { ids: ["P0"], content: ["Replacement"] }),
      /PROTECTED_FIELD|UNSUPPORTED_RICH_TARGET/,
    );
    assert.equal(f.writes(), 0);
  }
});

test("RED structure introduced after a plain target preview invalidates the approved write", async () => {
  const rows = [{ text: "Plain" }];
  const f = fixture(rows);
  const prepared = await prepareSafeWordWrite("office_replace_paragraphs", {
    ids: ["P0"],
    content: ["Changed"],
  });
  rows[0].xml =
    '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>Plain</w:t></w:r><w:commentRangeEnd w:id="1"/></w:p>';
  let receipts = 0;
  await assert.rejects(
    prepared.apply({
      guard() {},
      beforeWrite() {
        receipts++;
      },
    }),
    /STALE_PROPOSAL|UNSUPPORTED_RICH_TARGET/,
  );
  assert.equal(receipts, 0);
  assert.equal(f.writes(), 0);
});

test("RED plain result paragraph inside a cross-paragraph field cannot evade field protection", async () => {
  const rows = [
    {
      text: "",
      xml: '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TOC</w:instrText></w:r></w:p>',
    },
    { text: "Chapter One 1" },
    { text: "", xml: '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' },
  ];
  const f = fixture(rows);
  let receipts = 0;
  await assert.rejects(async () => {
    const prepared = await prepareSafeWordWrite("office_replace_paragraphs", {
      ids: ["P1"],
      content: ["Corrupt TOC"],
    });
    await prepared.apply({
      guard() {},
      beforeWrite() {
        receipts++;
      },
    });
  }, /PROTECTED_FIELD/);
  assert.equal(receipts, 0);
  assert.equal(f.writes(), 0);
});

test("RED heading comment cleanup preserves same-text foreign anchors and cross-boundary comments", async () => {
  const f = fixture([
    { text: "Background", heading: true },
    { text: "same body" },
    { text: "Background", heading: true },
    { text: "same body" },
  ]);
  const deleted = [];
  const comment = (id, indices) => ({
    id,
    delete() {
      deleted.push(id);
    },
    getRange: () => ({ paragraphs: { items: indices.map((i) => f.items[i]), load() {} } }),
  });
  f.body.getComments = () => ({
    items: [comment("first", [1]), comment("second", [3]), comment("cross", [1, 3])],
    load() {},
  });
  await assert.rejects(toolClearComments({ heading_section: "Background" }), /AMBIGUOUS_HEADING/);
  assert.deepEqual(deleted, []);
  const result = await toolClearComments({ heading_id: "P2" });
  assert.equal(result.cleared_comments, 1);
  assert.deepEqual(deleted, ["second"]);
});

test("RED header changes and fields are inspected independently of an unchanged body", async () => {
  const f = fixture([{ text: "Body stays unchanged" }]);
  let clears = 0,
    receipts = 0;
  const header = {
    text: "Plain header",
    rich: "",
    getReviewedText() {
      return { value: this.text };
    },
    getOoxml() {
      return {
        value: `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>${this.rich}<w:r><w:t>${this.text}</w:t></w:r></w:p></w:hdr>`,
      };
    },
    clear() {
      clears++;
    },
    insertText() {
      assert.fail("must not write");
    },
  };
  f.doc.sections = { items: [{ getHeader: () => header }], load() {} };
  header.paragraphs = { items: [{ uniqueLocalId: "HEADER-P1" }], load() {} };
  const prepared = await prepareSafeWordWrite("office_header_footer", {
    which: "header",
    text: "Replacement",
  });
  header.rich =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r>';
  await assert.rejects(
    prepared.apply({
      guard() {},
      beforeWrite() {
        receipts++;
      },
    }),
    /PROTECTED_FIELD|STALE_PROPOSAL/,
  );
  assert.equal(clears, 0);
  assert.equal(receipts, 0);
  assert.equal(f.rows[0].text, "Body stays unchanged");
});

test("RED field aliases cannot bypass XML safety and malformed/foreign namespace targets fail closed", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  assert.throws(
    () =>
      assertTextTargetSafe(`<x:p xmlns:x="${W}"><x:fldSimple x:instr="PAGE"/></x:p>`, {
        broad: true,
      }),
    /PROTECTED_FIELD/,
  );
  for (const raw of [
    `<w:p xmlns:w="${W}"><w:r></w:p>`,
    `<w:p xmlns:w="urn:not-word"><w:fldSimple/></w:p>`,
    "not XML",
  ])
    assert.throws(
      () => assertTextTargetSafe(raw, { broad: true }),
      /TARGET_INSPECTION_UNAVAILABLE/,
    );
  assert.doesNotThrow(() =>
    assertTextTargetSafe(`<w:p xmlns:w="${W}"><w:r><w:t>&lt;w:fldChar/&gt;</w:t></w:r></w:p>`, {
      broad: true,
    }),
  );
});
