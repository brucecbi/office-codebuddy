// Adversaries for the cell-only accommodation of native Mac getOoxml exports.
// Minimal FlatOPC shape matches recorded synthetic native evidence; host calls
// are synthetic and are not a claim about native rendering or save/reopen.
import { test } from "node:test";
import assert from "node:assert/strict";
import { prepareSafeWordWrite } from "../taskpane/shared/tools-word.js";
import { assertTextTargetSafe } from "../taskpane/shared/word-write-safety.js";
import { stableCellOoxml } from "../taskpane/shared/word-ooxml-state.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const escape = (text) => text.replaceAll("&", "&amp;").replaceAll("<", "&lt;");
function fixture({ shape = "single" } = {}) {
  let writes = 0;
  const state = { text: "Plain cell", id: "NATIVE-CELL-P1", generation: 1, rich: "", shape };
  const paragraph = () =>
    `<w:p w14:paraId="${state.generation === 1 ? "11111111" : "22222222"}" w14:textId="77777777" w:rsidR="${state.generation === 1 ? "00111111" : "00222222"}" w:rsidRDefault="${state.generation === 1 ? "00111111" : "00222222"}" w:rsidP="${state.generation === 1 ? "00111111" : "00222222"}"><w:r>${state.rich}<w:t>${escape(state.text)}</w:t></w:r></w:p>`;
  const xml = () => {
    const cell = `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>${paragraph()}${state.shape === "nested" ? "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Nested protected</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" : ""}</w:tc>`;
    return `<pkg:package xmlns:pkg="${PKG}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr w:rsidR="${state.generation === 1 ? "00111111" : "00222222"}" w:rsidTr="${state.generation === 1 ? "00111111" : "00222222"}">${cell}${state.shape === "multiple" ? cell : ""}</w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  };
  const cell = {
    paragraphs: {
      load() {},
      get items() {
        return [{ uniqueLocalId: state.id }];
      },
    },
    getOoxml: () => ({ value: xml() }),
    getReviewedText: () => ({ value: state.text + "\t" }),
    clear() {
      writes++;
      state.text = "";
    },
    insertText(text) {
      writes++;
      state.text = state.dropLiteralTab ? text.replace(/\t$/, "") : text;
    },
  };
  const body = {
    tables: { items: [{ getCell: () => ({ body: cell }) }], load() {} },
    getOoxml: () => ({ value: xml() }),
  };
  const doc = { body, changeTrackingMode: "Off", load() {}, getSelection: () => ({ select() {} }) };
  globalThis.Word = {
    run: async (fn) => fn({ document: doc, sync: async () => {} }),
    ChangeTrackingVersion: { current: "Current" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    InsertLocation: { start: "Start" },
  };
  return { state, xml, writes: () => writes };
}
const args = { table_index: 0, row: 0, column: 0, text: "Updated cell" };

test("RED cell-wrapper exception refuses nested/multiple cells and never relaxes generic broad targets", async () => {
  for (const shape of ["nested", "multiple"]) {
    const f = fixture({ shape });
    await assert.rejects(
      prepareSafeWordWrite("office_set_table_cell", args),
      /UNSUPPORTED|INSPECTION|AMBIGUOUS|CELL/,
    );
    assert.equal(f.writes(), 0);
  }
  const f = fixture();
  assert.throws(() => assertTextTargetSafe(f.xml(), { broad: true }), /UNSUPPORTED_RICH_TARGET/);
  const withRel = (target) =>
    f
      .xml()
      .replace(
        "</pkg:package>",
        `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="synthetic" Target="${target}"/></Relationships></pkg:xmlData></pkg:part></pkg:package>`,
      );
  assert.notEqual(
    stableCellOoxml(withRel("original.xml"), { paragraphIdentity: true }),
    stableCellOoxml(withRel("changed.xml"), { paragraphIdentity: true }),
    "cell normalization cannot erase a changed relationship target",
  );
});

test("RED cell-only export metadata changes may normalize with unchanged native IDs, text and structure", async () => {
  const f = fixture();
  let receipts = 0;
  const prepared = await prepareSafeWordWrite("office_set_table_cell", args);
  f.state.generation = 2;
  const result = await prepared.apply({
    guard() {},
    beforeWrite() {
      receipts++;
    },
  });
  assert.equal(receipts, 1);
  assert.equal(f.writes(), 2);
  assert.equal(f.state.text, args.text);
  assert.equal(result.verification, "target_current_text_readback");
});

test("RED replaced identity, new formatting and lost literal tab cannot be verified as the approved cell", async () => {
  for (const mutation of ["identity", "format"]) {
    const f = fixture();
    let receipts = 0;
    const prepared = await prepareSafeWordWrite("office_set_table_cell", args);
    if (mutation === "identity") f.state.id = "NEW-NATIVE-CELL-P1";
    else f.state.rich = "<w:rPr><w:b/></w:rPr>";
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
  }
  const f = fixture();
  f.state.dropLiteralTab = true;
  const prepared = await prepareSafeWordWrite("office_set_table_cell", {
    ...args,
    text: args.text + "\t",
  });
  await assert.rejects(prepared.apply({ guard() {}, beforeWrite() {} }), /RESULT_UNKNOWN/);
  assert.equal(f.writes(), 2, "host changed once; no automatic write replay");
  const correct = fixture();
  const literal = await prepareSafeWordWrite("office_set_table_cell", {
    ...args,
    text: args.text + "\t",
  });
  const result = await literal.apply({ guard() {}, beforeWrite() {} });
  assert.equal(result.verification, "target_current_text_readback");
  assert.equal(
    correct.state.text,
    args.text + "\t",
    "caller literal tab is retained independently of host delimiter",
  );
});
