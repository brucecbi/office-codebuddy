/* global Word, Office */
import {
  toolDocumentMap,
  toolReadStory,
  toolReadFormatting,
  toolReadComments,
  toolReadRevisions,
  prepareReviewWrite,
} from "/shared/word-review.js";

// Read-only diagnostic for the explicitly marked synthetic QA document.
// Preserve raw host output so differences can be investigated, not normalized away.
export async function inspectNativeSnapshots(marker) {
  await Word.run(async (c) => {
    c.document.body.load("text");
    await c.sync();
    if (c.document.body.text.split(/[\r\n]/)[0].trim() !== marker)
      throw new Error("REFUSED: snapshot diagnostics require the synthetic marker");
  });
  const snapshots = [];
  for (let i = 0; i < 2; i++)
    snapshots.push(
      await toolReadFormatting({
        target: { query: "Formatting target." },
        ooxml_limit: 100000,
      }),
    );
  const bodyOoxml = [];
  for (let i = 0; i < 2; i++)
    bodyOoxml.push(
      await Word.run(async (c) => {
        const result = c.document.body.getOoxml();
        await c.sync();
        return result.value;
      }),
    );
  return { snapshots, body_ooxml: bodyOoxml };
}

// A focused, explicitly mutating diagnostic in the marked disposable document.
// Preserve before/after raw exports; this is evidence collection, not a PASS.
export async function inspectNativeRevisionSemantics(marker) {
  const label = "DRAFTSPECT REVISION DIAGNOSTIC";
  await Word.run(async (c) => {
    c.document.body.load("text");
    c.document.load("changeTrackingMode");
    await c.sync();
    if (
      c.document.body.text.split(/[\r\n]/)[0].trim() !== marker ||
      c.document.body.text.includes(label + " 200")
    )
      throw new Error("REFUSED: requires marked synthetic document without a completed diagnostic");
    const previous = c.document.changeTrackingMode;
    try {
      c.document.changeTrackingMode = "Off";
      if (!c.document.body.text.includes(label))
        c.document.body.insertParagraph(label + " 100", "End");
      await c.sync();
    } finally {
      c.document.changeTrackingMode = previous;
      await c.sync();
    }
  });
  const target_snapshots = [];
  for (let i = 0; i < 2; i++)
    target_snapshots.push(
      await toolReadFormatting({ target: { query: label + " 100" }, ooxml_limit: 100000 }),
    );
  // Direct host calls construct/observe this fixture; they are not a claimed
  // production guarded-write PASS. The twelve checks exercise production code.
  await Word.run(async (c) => {
    const hits = c.document.body.search(label + " 100", { matchCase: true });
    hits.load("items");
    c.document.load("changeTrackingMode");
    await c.sync();
    if (hits.items.length !== 1) throw new Error("Expected one diagnostic fixture range");
    const previous = c.document.changeTrackingMode;
    try {
      c.document.changeTrackingMode = "TrackAll";
      hits.items[0].insertText(label + " 200", "Replace");
      await c.sync();
    } finally {
      c.document.changeTrackingMode = previous;
      await c.sync();
    }
  });
  const raw = () =>
    Word.run(async (c) => {
      const value = c.document.body.getOoxml();
      await c.sync();
      return value.value;
    });
  const before = await toolReadRevisions(),
    before_body_ooxml = await raw();
  const selected = before.revisions.filter(
    (r) => r.type === "Added" && r.current.includes(label + " 200"),
  );
  if (selected.length !== 1) throw new Error("Expected one synthetic diagnostic Added revision");
  let outcome;
  try {
    outcome = await Word.run(async (c) => {
      const revisions = c.document.body.getTrackedChanges();
      revisions.load("items/type,items/text");
      await c.sync();
      const hits = revisions.items.filter((r) => r.type === "Added" && r.text === label + " 200");
      if (hits.length !== 1) throw new Error("Expected one exact diagnostic Added revision");
      hits[0].accept();
      await c.sync();
      return { direct_host_diagnostic: true, production_guard_test: false };
    });
  } catch (error) {
    outcome = { error: error.message };
  }
  const after = await toolReadRevisions(),
    after_body_ooxml = await raw();
  return {
    target_snapshots,
    before,
    before_body_ooxml,
    selected_indices: [selected[0].index],
    outcome,
    after,
    after_body_ooxml,
  };
}

export async function runNativeChecks(marker, emit) {
  const results = [];
  const assert = (ok, message) => {
    if (!ok) throw new Error(message);
  };
  const check = async (name, fn) => {
    try {
      const detail = await fn();
      const row = { name, status: "PASS", detail };
      results.push(row);
      emit(row);
      return detail;
    } catch (e) {
      const row = {
        name,
        status: "FAIL",
        error: e.message,
        code: e.code,
        debug: e.debugInfo,
        ...(e.qa_detail ? { detail: e.qa_detail } : {}),
      };
      results.push(row);
      emit(row);
      return null;
    }
  };
  const controls = { guard() {}, beforeWrite() {} };
  const initialized = await check("synthetic_document_guard_and_setup", () =>
    Word.run(async (c) => {
      c.document.body.load("text");
      await c.sync();
      assert(
        c.document.body.text.trim() === marker,
        "REFUSED: this is not the new marker-only synthetic document",
      );
      c.document.changeTrackingMode = "Off";
      c.document.body.insertText(
        marker +
          "\nThe price is USD 100.\nDuplicate sentence.\nDuplicate sentence.\nFormatting target.",
        "Replace",
      );
      c.document.sections.load("items");
      await c.sync();
      c.document.sections.items[0]
        .getHeader("Primary")
        .insertText("Synthetic header 100", "Replace");
      c.document.sections.items[0].getFooter("Primary").insertText("Synthetic footer", "Replace");
      await c.sync();
      return true;
    }),
  );
  if (!initialized) return results;
  await check("current_original_body_read", async () => {
    const a = await toolReadStory(),
      b = await toolReadStory({ version: "Original" });
    assert(
      a.paragraphs.some((p) => p.text.includes("USD 100")),
      "Current missing price",
    );
    assert(
      b.paragraphs.some((p) => p.text.includes("USD 100")),
      "Original missing price",
    );
    return { paragraphs: a.paragraphs.length, addressing: a.addressing };
  });
  await check("document_story_map", async () => {
    const map = await toolDocumentMap();
    assert(
      map.stories.some((s) => s.target.story === "header" && s.current_characters > 0),
      "Header not inventoried",
    );
    return map;
  });
  await check("format_ooxml_and_scoped_write", async () => {
    const target = { story: "body", query: "Formatting target." };
    const before = await toolReadFormatting({ target });
    assert(before.ooxml.includes("<"), "No OOXML returned");
    const plan = await prepareReviewWrite("office_edit_range", {
      target,
      action: "format",
      font: { bold: true, italic: true },
    });
    await plan.apply(controls);
    const after = await toolReadFormatting({ target });
    assert(
      after.font.bold === true && after.font.italic === true,
      "Requested format not read back",
    );
    assert(after.current === before.current, "Formatting changed text");
    return { font: after.font, current: after.current };
  });
  await check("ambiguous_duplicate_refusal", async () => {
    try {
      await prepareReviewWrite("office_edit_range", {
        target: { query: "Duplicate sentence." },
        action: "format",
        font: { bold: true },
      });
    } catch (e) {
      assert(e.message.includes("AMBIGUOUS_RANGE"), e.message);
      return true;
    }
    throw new Error("Ambiguous edit was not refused");
  });
  await check("stale_format_proposal_refusal", async () => {
    const args = { target: { query: "Formatting target." }, action: "format", font: { size: 15 } };
    const plan = await prepareReviewWrite("office_edit_range", args);
    await Word.run(async (c) => {
      const hits = c.document.body.search("Formatting target.");
      hits.load("items");
      await c.sync();
      hits.items[0].font.color = "#336699";
      await c.sync();
    });
    try {
      await plan.apply(controls);
    } catch (e) {
      assert(e.message.includes("STALE_PROPOSAL"), e.message);
      return true;
    }
    throw new Error("Stale proposal was not refused");
  });
  await check("tracked_header_replace_readback", async () => {
    const plan = await prepareReviewWrite("office_edit_range", {
      target: { story: "header", section: 0, query: "100" },
      action: "replace_text",
      text: "120",
    });
    await plan.apply(controls);
    const story = await toolReadStory({ target: { story: "header", section: 0 } });
    assert(
      story.paragraphs.some((p) => p.text.includes("120")),
      "Header replacement not Current",
    );
    return story.paragraphs;
  });
  await check("comment_thread_crud_and_anchor", async () => {
    const add = await prepareReviewWrite("office_manage_comment", {
      target: { query: "USD 100" },
      action: "add",
      text: "Synthetic review comment",
    });
    const inserted = await add.apply(controls),
      id = inserted.comment.id;
    for (const [action, text] of [
      ["reply", "Synthetic reply"],
      ["edit", "Updated comment"],
      ["resolve"],
      ["reopen"],
    ]) {
      await (
        await prepareReviewWrite("office_manage_comment", { action, comment_id: id, text })
      ).apply(controls);
    }
    const row = (await toolReadComments()).comments.find((c) => c.id === id);
    assert(row.anchor.current === "USD 100", "Comment anchor not exact");
    assert(row.replies.length === 1 && !row.resolved, "Replies/resolution mismatch");
    return row;
  });
  await check("tracked_body_replace_current_original", async () => {
    await (
      await prepareReviewWrite("office_edit_range", {
        target: { query: "100" },
        action: "replace_text",
        text: "150",
      })
    ).apply(controls);
    const a = await toolReadStory(),
      b = await toolReadStory({ version: "Original" });
    assert(
      a.paragraphs.some((p) => p.text.includes("USD 150")),
      "Current did not update",
    );
    assert(
      b.paragraphs.some((p) => p.text.includes("USD 100")),
      "Original was lost",
    );
    return { current: a.paragraphs, original: b.paragraphs };
  });
  await check("existing_revision_inspection_and_accept", async () => {
    const read = await toolReadRevisions();
    const index = read.revisions.findIndex(
      (r) => r.type === "Added" && (r.current.includes("150") || r.text.includes("150")),
    );
    assert(index >= 0, "Added price revision not exposed");
    let result;
    try {
      result = await (
        await prepareReviewWrite("office_manage_revisions", {
          action: "accept",
          snapshot_id: read.snapshot_id,
          indices: [index],
        })
      ).apply(controls);
    } catch (error) {
      const metadata = (rows) =>
        rows.map(({ ooxml, ...row }) => ({ ...row, ooxml_characters: ooxml.length }));
      error.qa_detail = {
        selected_indices: [index],
        before: metadata(read.revisions),
        after: metadata((await toolReadRevisions()).revisions),
      };
      throw error;
    }
    return {
      before: read.revisions.map(({ index, type, text }) => ({ index, type, text })),
      result,
    };
  });
  await check("footnote_and_table_scope", async () => {
    await Word.run(async (c) => {
      c.document.body.getRange("End").insertFootnote("Synthetic footnote 100");
      c.document.body.insertTable(2, 2, "End", [
        ["A1", "B1"],
        ["A2", "B2"],
      ]);
      await c.sync();
    });
    const note = await toolReadStory({ target: { story: "footnote", note: 0 } });
    const cell = await toolReadFormatting({ target: { table: 0, row: 1, column: 1 } });
    assert(
      note.paragraphs.some((p) => p.text.includes("Synthetic footnote")),
      "Footnote missing",
    );
    assert(cell.current.includes("B2"), "Wrong table cell");
    return { note: note.paragraphs, cell: cell.current };
  });
  await check("tracking_mode_restored", () =>
    Word.run(async (c) => {
      c.document.load("changeTrackingMode");
      await c.sync();
      assert(c.document.changeTrackingMode === "Off", "Tracking mode was not restored");
      return c.document.changeTrackingMode;
    }),
  );
  return results;
}
