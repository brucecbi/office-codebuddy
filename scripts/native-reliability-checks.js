/* global Word, Office */
import {
  toolReadParagraphs,
  toolInsertParagraphs,
  toolReplaceSection,
  toolClearHighlights,
  toolClearComments,
  toolSetTableCell,
  toolHeaderFooter,
  prepareSafeWordWrite,
  prepareTextReplacement,
  toolApplyTextReplacement,
} from "/shared/tools-word.js";

// Read-only evidence for native export wrappers and paragraph marks. No
// normalization or fixture modification is performed by this diagnostic.
export async function inspectNativeReliabilityTargets(marker) {
  return Word.run(async (c) => {
    c.document.body.load("text");
    await c.sync();
    if (
      typeof marker !== "string" ||
      marker.length <= 10 ||
      c.document.body.text.split(/[\r\n]/)[0].trim() !== marker
    )
      throw new Error("REFUSED: diagnostics require the designated synthetic document");
    const paragraphs = c.document.body.paragraphs;
    paragraphs.load("items/uniqueLocalId,items/text,items/style,items/styleBuiltIn");
    c.document.body.tables.load("items");
    await c.sync();
    const current = paragraphs.items.map((p) => p.getRange().getReviewedText("Current"));
    const table = c.document.body.tables.items[0];
    const cells = [];
    if (table) {
      for (let row = 0; row < 4; row++) {
        const body = table.getCell(row, 0).body;
        body.paragraphs.load("items/uniqueLocalId");
        const value = body.getOoxml();
        const text = body.getReviewedText("Current");
        await c.sync();
        const paragraphXml = body.paragraphs.items.map((p) => p.getRange().getOoxml());
        const repeated = body.getOoxml();
        await c.sync();
        cells.push({
          row,
          current: text.value,
          ooxml: value.value,
          repeated_ooxml: repeated.value,
          paragraphs: body.paragraphs.items.map((p, i) => ({
            id: p.uniqueLocalId,
            ooxml: paragraphXml[i].value,
          })),
        });
      }
    }
    await c.sync();
    return {
      paragraphs: paragraphs.items.map((p, i) => ({
        id: p.uniqueLocalId,
        raw: p.text,
        current: current[i].value,
        style: p.style,
        styleBuiltIn: p.styleBuiltIn,
      })),
      cells,
    };
  });
}

const withoutParagraphMark = (text) => text.replace(/\r$/, "");
const fieldFixturePackage = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document></pkg:xmlData>
  </pkg:part>
</pkg:package>`;

// Explicitly mutating, model-free acceptance. Only a new marker-only disposable
// document may be initialized. Every test rechecks the marker/document binding.
export async function runNativeReliabilityChecks(marker, onResult = () => {}) {
  const rows = [],
    url = Office.context.document.url || null;
  const assert = (ok, message) => {
    if (!ok) throw new Error(message);
  };
  const controls = {
    guard() {
      assert((Office.context.document.url || null) === url, "REFUSED: active document changed");
    },
    beforeWrite() {
      controls.guard();
    },
  };
  const bound = async (c, onlyMarker = false) => {
    controls.guard();
    c.document.body.load("text");
    await c.sync();
    controls.guard();
    assert(typeof marker === "string" && marker.length > 10, "REFUSED: invalid synthetic marker");
    assert(
      onlyMarker
        ? c.document.body.text.trim() === marker
        : c.document.body.text.split(/[\r\n]/)[0].trim() === marker,
      "REFUSED: not the designated disposable synthetic document",
    );
  };
  const run = async (name, fn) => {
    try {
      await Word.run((c) => bound(c));
      const detail = await fn();
      const row = { name, status: "PASS", detail };
      rows.push(row);
      onResult(row);
      return detail;
    } catch (e) {
      const row = { name, status: "FAIL", error: e.message, code: e.code, debug: e.debugInfo };
      rows.push(row);
      onResult(row);
      return null;
    }
  };
  const rejectsBeforeWrite = async (fn, pattern) => {
    const before = await toolReadParagraphs({ preview: false });
    let rejected = false;
    try {
      await fn();
    } catch (e) {
      assert(pattern.test(e.message), `Unexpected rejection: ${e.message}`);
      rejected = true;
    }
    assert(rejected, "Expected rejection before mutation");
    const after = await toolReadParagraphs({ preview: false });
    assert(
      JSON.stringify(before.paragraphs) === JSON.stringify(after.paragraphs),
      "Rejected operation changed Current body text",
    );
    return { rejected_before_mutation: true };
  };
  let previousMode;
  try {
    await Word.run(async (c) => {
      await bound(c, true);
      c.document.load("changeTrackingMode");
      await c.sync();
      previousMode = c.document.changeTrackingMode;
      c.document.changeTrackingMode = "Off";
      c.document.body.insertText(
        marker +
          "\nRepeated heading\nFirst section body\nRepeated heading\nSecond section body\nwidget widgets\nOld heading\nRevised heading body\nFormatting target.",
        "Replace",
      );
      c.document.body.paragraphs.load("items");
      await c.sync();
      for (const i of [1, 3, 6]) c.document.body.paragraphs.items[i].styleBuiltIn = "Heading1";
      c.document.body.insertTable(4, 1, "End", [
        ["Plain cell"],
        ["Mixed cell"],
        ["Field cell"],
        ["Revision cell"],
      ]);
      c.document.sections.load("items");
      await c.sync();
      c.document.sections.items[0]
        .getHeader("Primary")
        .insertText("Plain synthetic header", "Replace");
      await c.sync();
    });
    const setup = { name: "reliability_fixture_guard_and_setup", status: "PASS" };
    rows.push(setup);
    onResult(setup);
  } catch (e) {
    const row = { name: "reliability_fixture_guard_and_setup", status: "FAIL", error: e.message };
    rows.push(row);
    onResult(row);
    return rows;
  }
  try {
    await run("duplicate_current_headings_refuse_all_five_paths", async () => {
      for (const fn of [
        () => toolReadParagraphs({ heading_section: "Repeated heading" }),
        () =>
          toolInsertParagraphs({
            after: { heading: "Repeated heading" },
            content: ["MUST NOT INSERT"],
          }),
        () => toolReplaceSection({ heading: "Repeated heading", content: ["MUST NOT REPLACE"] }),
        () => toolClearHighlights({ heading_section: "Repeated heading" }),
        () => toolClearComments({ heading_section: "Repeated heading" }),
      ])
        await rejectsBeforeWrite(fn, /AMBIGUOUS_HEADING/);
      return { rejected_paths: 5 };
    });
    await run("word_whole_word_exact_plan_and_tracked_readback", async () => {
      const p = (await toolReadParagraphs({ preview: false })).paragraphs.find(
        (p) => withoutParagraphMark(p.text) === "widget widgets",
      );
      assert(p, "Missing whole-word fixture");
      const args = {
        paragraph_ids: [p.id],
        find: "widget",
        replace: "device",
        whole_word: true,
        match_case: true,
      };
      const plan = await prepareTextReplacement(args);
      assert(
        withoutParagraphMark(plan[0].after) === "device widgets",
        "Whole-word plan changed excluded substring",
      );
      return toolApplyTextReplacement(args, plan, controls);
    });
    await run("plain_cell_target_precondition_and_readback", async () => {
      const prepared = await prepareSafeWordWrite("office_set_table_cell", {
        table_index: 0,
        row: 0,
        column: 0,
        text: "Verified plain cell",
        track_changes: true,
      });
      return prepared.apply(controls);
    });
    await run("plain_header_target_precondition_and_readback", async () => {
      const prepared = await prepareSafeWordWrite("office_header_footer", {
        which: "header",
        text: "Verified synthetic header",
      });
      return prepared.apply(controls);
    });
    await run("cell_surgical_replacement_uses_native_wrapper_comparison", async () => {
      const p = (await toolReadParagraphs({ preview: false })).paragraphs.find(
        (p) => p.text.replace(/\t$/, "") === "Verified plain cell",
      );
      assert(p, "Missing successful plain-cell fixture");
      const args = {
        paragraph_ids: [p.id],
        find: "plain",
        replace: "safe",
        whole_word: true,
        match_case: true,
      };
      const plan = await prepareTextReplacement(args);
      assert(
        plan[0].after === "Verified safe cell\t",
        "Cell plan did not preserve surrounding text and boundary",
      );
      return toolApplyTextReplacement(args, plan, controls);
    });
    await run("header_changed_without_body_change_rejects_stale", async () => {
      const prepared = await prepareSafeWordWrite("office_header_footer", {
        which: "header",
        text: "MUST NOT APPLY",
      });
      await Word.run(async (c) => {
        await bound(c);
        c.document.sections.load("items");
        await c.sync();
        c.document.sections.items[0]
          .getHeader("Primary")
          .insertText("Manual synthetic header change", "Replace");
        await c.sync();
      });
      let writes = 0,
        rejected = false;
      try {
        await prepared.apply({
          guard: controls.guard,
          beforeWrite() {
            writes++;
          },
        });
      } catch (e) {
        assert(/STALE_PROPOSAL/.test(e.message), e.message);
        rejected = true;
      }
      assert(rejected && writes === 0, "Stale header was not rejected before write");
      return { stale_target_rejected: true };
    });
    await run("Current_revised_heading_is_used_for_read_and_insertion", async () => {
      await Word.run(async (c) => {
        await bound(c);
        const hits = c.document.body.search("Old heading", { matchCase: true });
        hits.load("items");
        await c.sync();
        assert(hits.items.length === 1, "Missing revised-heading fixture");
        c.document.changeTrackingMode = "TrackAll";
        hits.items[0].insertText("New heading", "Replace");
        await c.sync();
        c.document.changeTrackingMode = "Off";
        await c.sync();
      });
      const read = await toolReadParagraphs({ heading_section: "New heading" });
      assert(
        withoutParagraphMark(read.paragraphs[0].text) === "New heading",
        "Current revised heading did not resolve",
      );
      return toolInsertParagraphs({
        after: { heading: "New heading" },
        content: ["Inserted by Current heading"],
        track_changes: true,
      });
    });
    await run("native_rich_cell_field_revision_and_mixed_run_refusal", async () => {
      await Word.run(async (c) => {
        await bound(c);
        c.document.body.tables.load("items");
        await c.sync();
        const t = c.document.body.tables.items[0];
        const mixed = t.getCell(1, 0).body.search("Mixed", { matchCase: true });
        mixed.load("items");
        await c.sync();
        mixed.items[0].font.bold = true;
        t.getCell(2, 0).body.insertOoxml(fieldFixturePackage, "Replace");
        await c.sync();
        c.document.changeTrackingMode = "TrackAll";
        t.getCell(3, 0).body.insertText("Tracked revision cell", "Replace");
        await c.sync();
        c.document.changeTrackingMode = "Off";
        await c.sync();
      });
      for (const row of [1, 2, 3])
        await rejectsBeforeWrite(
          () => toolSetTableCell({ table_index: 0, row, column: 0, text: "MUST NOT REPLACE" }),
          /UNSUPPORTED_RICH_TARGET|PROTECTED_FIELD/,
        );
      return { protected_cells: 3 };
    });
  } finally {
    await Word.run(async (c) => {
      await bound(c);
      c.document.changeTrackingMode = previousMode;
      await c.sync();
    });
  }
  return rows;
}
