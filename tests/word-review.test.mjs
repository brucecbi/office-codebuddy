import { test } from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";
import {
  toolDocumentMap,
  toolReadStory,
  toolReadFormatting,
  toolReadComments,
  toolReadRevisions,
  toolAuditRedline,
  prepareReviewWrite,
} from "../taskpane/shared/word-review.js";
import { toolClearComments, toolAddComment } from "../taskpane/shared/tools-word.js";

// Deterministic Office-shaped proxies: flow/verification tests, not native proof.
function fixture({ notes = true, shapes = false } = {}) {
  let writes = 0;
  const list = (get) => ({
    load() {},
    get items() {
      return get();
    },
  });
  const font = () => ({
    bold: false,
    italic: false,
    underline: "None",
    size: 11,
    color: "#000000",
    name: "Arial",
    highlightColor: null,
    strikeThrough: false,
    subscript: false,
    superscript: false,
    load() {},
  });
  function body(text, id) {
    const b = { text, original: text, id, comments: [], revisions: [], fonts: {}, load() {} };
    const p = {
      uniqueLocalId: id,
      style: "Normal",
      alignment: "Left",
      leftIndent: 0,
      rightIndent: 0,
      firstLineIndent: 0,
      spaceBefore: 0,
      spaceAfter: 0,
      lineSpacing: 12,
      get text() {
        return b.text;
      },
      listItemOrNullObject: { isNullObject: true },
      getRange: () => range(b),
      search: (q) => range(b).search(q),
    };
    b.paragraphs = list(() => [p]);
    b.tables = list(() => []);
    b.shapes = list(() => []);
    b.getRange = () => range(b);
    b.getReviewedText = (v) => ({
      get value() {
        return v === "Original" ? b.original : b.text;
      },
    });
    b.getOoxml = () => ({
      value: `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>${JSON.stringify(
        {
          text: b.text,
          fonts: b.fonts,
          style: p.style,
          revisions: b.revisions.map((r) => r.type + r.text),
        },
      )
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")}</w:t></w:r></w:p>`,
    });
    b.getComments = () => list(() => b.comments);
    b.getTrackedChanges = () => list(() => b.revisions);
    return b;
  }
  function range(b, start = 0, length = null) {
    const key = `${start}:${length}`;
    const f = (b.fonts[key] ??= font());
    return {
      font: f,
      paragraphs: b.paragraphs,
      load() {},
      select() {},
      get text() {
        return length === null ? b.text : b.text.slice(start, start + length);
      },
      getReviewedText(v) {
        const text = v === "Original" ? b.original : b.text;
        return { value: length === null ? text : text.slice(start, start + length) };
      },
      getOoxml: () => b.getOoxml(),
      search(q) {
        const hits = [];
        let at = b.text.indexOf(q);
        while (at !== -1) {
          hits.push(range(b, at, q.length));
          at = b.text.indexOf(q, at + q.length);
        }
        return list(() => hits);
      },
      insertText(t) {
        assert.equal(doc.changeTrackingMode, "TrackAll");
        writes++;
        const count = length ?? b.text.length;
        b.text = b.text.slice(0, start) + t + b.text.slice(start + count);
        return range(b, start, t.length);
      },
      insertComment(t) {
        writes++;
        const c = comment(b, t);
        b.comments.push(c);
        return c;
      },
    };
  }
  function comment(b, text, id = randomUUID()) {
    const replies = [];
    const c = {
      id,
      content: text,
      authorName: "Fixture author",
      creationDate: new Date("2026-01-01T00:00:00Z"),
      resolved: false,
      load() {},
      contentRange: { text }, // Deliberately has NO paragraphs member.
      getRange: () => range(b),
      replies: list(() => replies),
      reply(t) {
        writes++;
        replies.push({
          id: randomUUID(),
          content: t,
          authorName: "Fixture",
          creationDate: new Date("2026-01-02T00:00:00Z"),
        });
      },
      delete() {
        writes++;
        b.comments = b.comments.filter((x) => x !== c);
      },
    };
    return c;
  }
  const main = body("The price is USD 100.", "P1"),
    header = body("Header 100", "H1"),
    note = body("Footnote 100", "N1");
  main.original = "The price is USD 90.";
  main.footnotes = list(() => [{ body: note }]);
  main.endnotes = list(() => []);
  const doc = {
    body: main,
    load() {},
    changeTrackingMode: "Off",
    getSelection: () => range(main),
    sections: list(() => [{ getHeader: () => header, getFooter: () => body("", "F1") }]),
  };
  const context = { document: doc, sync: async () => {} };
  globalThis.Office = {
    context: {
      requirements: {
        isSetSupported: (name, v) =>
          name === "WordApiDesktop" ? shapes : v === "1.5" ? notes : true,
      },
    },
  };
  globalThis.Word = {
    run: async (fn) => fn(context),
    ChangeTrackingVersion: { current: "Current" },
  };
  return {
    main,
    header,
    note,
    doc,
    body,
    range,
    list,
    comment,
    get writes() {
      return writes;
    },
    addRevision(type, text) {
      const r = {
        author: "Fixture",
        date: new Date("2026-01-01"),
        type,
        text,
        getRange: () => range(main),
        accept() {
          writes++;
          main.revisions = main.revisions.filter((x) => x !== r);
        },
        reject() {
          this.accept();
        },
      };
      main.revisions.push(r);
      return r;
    },
  };
}
const controls = { guard() {}, beforeWrite() {} };
const nativeOoxml = JSON.parse(
  readFileSync(new URL("./fixtures/word-native-ooxml-snapshots.json", import.meta.url)),
);

test("document map enumerates all header variants and notes; unavailable objects stay explicit", async () => {
  const f = fixture();
  const map = await toolDocumentMap();
  assert.equal(map.stories.filter((s) => s.target.story === "header").length, 3);
  assert.equal(map.stories.filter((s) => s.target.story === "footer").length, 3);
  assert.ok(map.stories.some((s) => s.target.story === "footnote"));
  assert.ok(map.coverage.gaps.some((s) => s.component === "shapes"));
  assert.equal(map.coverage.complete_file_review, false);
  assert.equal(f.writes, 0);
});
test("Current and Original story reads remain distinct, paginated and untruncated", async () => {
  const f = fixture();
  const current = await toolReadStory({ limit: 1 }),
    original = await toolReadStory({ version: "Original" });
  assert.equal(current.paragraphs[0].text, f.main.text);
  assert.equal(original.paragraphs[0].text, f.main.original);
  assert.equal(current.next_offset, null);
  assert.equal(
    (await toolReadStory({ target: { story: "footnote", note: 0 } })).paragraphs[0].text,
    f.note.text,
  );
  fixture({ notes: false });
  await assert.rejects(
    toolReadStory({ target: { story: "footnote", note: 0 } }),
    /UNSUPPORTED_API/,
  );
});
test("format inspection returns exact OOXML chunks and scoped effective properties", async () => {
  fixture();
  const r = await toolReadFormatting({
    target: { paragraph_id: "P1", query: "USD" },
    ooxml_limit: 10,
  });
  assert.equal(r.current, "USD");
  assert.equal(r.font.name, "Arial");
  assert.equal(r.ooxml.length, 10);
  assert.equal(r.next_ooxml_offset, 10);
  assert.ok(!("body_ooxml" in r));
});
test("OOXML pagination requires its unchanged snapshot and story readers reject ignored selectors", async () => {
  const f = fixture();
  const first = await toolReadFormatting({ ooxml_limit: 10 });
  await assert.rejects(toolReadFormatting({ ooxml_offset: 10 }), /snapshot_id/);
  const second = await toolReadFormatting({
    snapshot_id: first.snapshot_id,
    ooxml_offset: 10,
    ooxml_limit: 10,
  });
  assert.equal(second.ooxml_offset, 10);
  f.main.text += " manual";
  await assert.rejects(
    toolReadFormatting({ snapshot_id: first.snapshot_id, ooxml_offset: 20 }),
    /STALE_PROPOSAL/,
  );
  await assert.rejects(toolReadStory({ target: { table: 0 } }), /entire story/);
});
test("scoped font edit preserves text and tracking mode and verifies requested property", async () => {
  const f = fixture();
  const plan = await prepareReviewWrite("office_edit_range", {
    action: "format",
    target: { paragraph_id: "P1", query: "USD" },
    font: { bold: true },
  });
  assert.equal(plan.preview.before.text, "USD");
  assert.equal((await plan.apply(controls)).verified, true);
  assert.equal((await toolReadFormatting({ target: { query: "USD" } })).font.bold, true);
  assert.equal(f.main.text, "The price is USD 100.");
  assert.equal(f.doc.changeTrackingMode, "Off");
});
test("do-not-enable leaves an already tracked document tracked", async () => {
  for (const initial of ["TrackAll", "TrackMine"]) {
    const f = fixture();
    const assigned = [];
    let mode = initial;
    Object.defineProperty(f.doc, "changeTrackingMode", {
      configurable: true,
      get() {
        return mode;
      },
      set(value) {
        assigned.push(value);
        mode = value;
      },
    });
    const plan = await prepareReviewWrite("office_edit_range", {
      action: "format",
      target: { paragraph_id: "P1", query: "USD" },
      font: { bold: true },
      track_changes: false,
    });
    assert.equal(plan.preview.tracked, true);
    assert.equal((await plan.apply(controls)).verified, true);
    assert.equal(mode, initial);
    assert.equal(assigned.includes("Off"), false);
  }
});
test("range text replacement does not switch off existing tracking", async () => {
  const f = fixture();
  f.doc.changeTrackingMode = "TrackAll";
  const plan = await prepareReviewWrite("office_edit_range", {
    action: "replace_text",
    target: { paragraph_id: "P1", query: "USD" },
    text: "EUR",
    track_changes: false,
  });
  assert.equal(plan.preview.tracked, true);
  await plan.apply(controls);
  assert.equal(f.main.text, "The price is EUR 100.");
  assert.equal(f.doc.changeTrackingMode, "TrackAll");
});
test("header and note edits use the selected story, verify surrounding text and restore tracking", async () => {
  const f = fixture();
  const plan = await prepareReviewWrite("office_edit_range", {
    action: "replace_text",
    target: { story: "header", section: 0, query: "100" },
    text: "120",
  });
  await plan.apply(controls);
  assert.equal(f.header.text, "Header 120");
  assert.equal(f.main.text, "The price is USD 100.");
  assert.equal(f.doc.changeTrackingMode, "Off");
});
test("manual text or formatting changes and cancellation invalidate a range proposal", async () => {
  const f = fixture();
  const p = await prepareReviewWrite("office_edit_range", {
    action: "format",
    target: {},
    font: { italic: true },
  });
  f.main.text = "Manual change";
  await assert.rejects(p.apply(controls), /STALE_PROPOSAL/);
  assert.equal(f.writes, 0);
  const p2 = await prepareReviewWrite("office_edit_range", {
    action: "format",
    target: {},
    font: { italic: true },
  });
  await assert.rejects(
    p2.apply({
      guard() {
        throw new Error("cancelled");
      },
      beforeWrite() {
        assert.fail();
      },
    }),
    /cancelled/,
  );
  assert.equal(f.doc.changeTrackingMode, "Off");
});
test("ambiguous range and unknown format keys fail before a write", async () => {
  const f = fixture();
  f.main.text = "USD and USD";
  await assert.rejects(
    prepareReviewWrite("office_edit_range", {
      action: "format",
      target: { query: "USD" },
      font: { bold: true },
    }),
    /AMBIGUOUS_RANGE/,
  );
  await assert.rejects(
    prepareReviewWrite("office_edit_range", { action: "format", font: { imaginary: true } }),
    /Unsupported format/,
  );
  await assert.rejects(
    prepareReviewWrite("office_edit_range", { action: "format", font: { bold: "false" } }),
    /Invalid format/,
  );
  assert.equal(f.writes, 0);
});
test("comments expose actual document anchors and full replies; editing is verified", async () => {
  const f = fixture();
  const c = f.comment(f.main, "Please review", "C1");
  f.main.comments.push(c);
  c.reply("Reply one");
  const read = await toolReadComments();
  assert.equal(read.comments[0].anchor.current, f.main.text);
  assert.deepEqual(read.comments[0].anchor.paragraph_ids, ["P1"]);
  assert.equal(read.comments[0].replies[0].text, "Reply one");
  const plan = await prepareReviewWrite("office_manage_comment", {
    action: "edit",
    comment_id: "C1",
    text: "Updated comment",
  });
  assert.equal((await plan.apply(controls)).comment.text, "Updated comment");
});
test("a new reply after preview invalidates comment deletion", async () => {
  const f = fixture();
  const c = f.comment(f.main, "Keep", "C1");
  f.main.comments.push(c);
  const plan = await prepareReviewWrite("office_manage_comment", {
    action: "delete",
    comment_id: "C1",
  });
  c.reply("New user reply");
  await assert.rejects(plan.apply(controls), /STALE_PROPOSAL/);
  assert.equal(f.main.comments.length, 1);
});
test("revision snapshots show both text versions and accept only explicit indices", async () => {
  const f = fixture();
  f.addRevision("Added", "100");
  f.addRevision("Deleted", "");
  const read = await toolReadRevisions();
  assert.equal(read.revisions[1].original, f.main.original);
  const plan = await prepareReviewWrite("office_manage_revisions", {
    action: "accept",
    snapshot_id: read.snapshot_id,
    indices: [0],
  });
  const result = await plan.apply(controls);
  assert.equal(result.remaining, 1);
  assert.equal(f.main.revisions[0].type, "Deleted");
  await assert.rejects(
    prepareReviewWrite("office_manage_revisions", {
      action: "reject",
      snapshot_id: read.snapshot_id,
      indices: [0],
    }),
    /expired/,
  );
});
test("changed revision list refuses stale acceptance and never substitutes accept-all", async () => {
  const f = fixture();
  f.addRevision("Added", "100");
  const read = await toolReadRevisions();
  const plan = await prepareReviewWrite("office_manage_revisions", {
    action: "reject",
    snapshot_id: read.snapshot_id,
    indices: [0],
  });
  f.addRevision("Deleted", "90");
  await assert.rejects(plan.apply(controls), /STALE_PROPOSAL/);
  assert.equal(f.writes, 0);
});
test("scoped comment cleanup uses anchor IDs, not duplicate paragraph text", async () => {
  const f = fixture();
  const twin = f.body(f.main.text, "P2");
  const p1 = f.main.paragraphs.items[0],
    p2 = twin.paragraphs.items[0];
  const c = f.comment(f.main, "Comment on second duplicate", "C2");
  c.getRange = () => f.range(twin);
  f.main.comments.push(c);
  f.main.paragraphs = f.list(() => [p1, p2]);
  assert.equal((await toolClearComments({ paragraph_ids: ["P1"] })).cleared_comments, 0);
  assert.equal(f.main.comments.length, 1);
  assert.equal((await toolClearComments({ paragraph_ids: ["P2"] })).cleared_comments, 1);
  assert.equal(f.main.comments.length, 0);
});
test("legacy add-comment no longer silently picks the first repeated query", async () => {
  const f = fixture();
  f.main.text = "USD USD";
  await assert.rejects(
    toolAddComment({ paragraph_id: "P1", query: "USD", text: "Review" }),
    /AMBIGUOUS_RANGE/,
  );
  assert.equal(f.writes, 0);
});

test("native export metadata fluctuation permits unchanged guarded format/comment/revision operations", async () => {
  const f = fixture();
  let exports = 0;
  f.main.getOoxml = () => ({ value: nativeOoxml.body[exports++ % 2] });
  const format = await prepareReviewWrite("office_edit_range", {
    action: "format",
    font: { bold: true },
  });
  assert.equal((await format.apply(controls)).verified, true);
  const comment = await prepareReviewWrite("office_manage_comment", {
    action: "add",
    text: "Review this",
  });
  assert.equal((await comment.apply(controls)).verified, true);
  f.addRevision("Added", "100");
  const read = await toolReadRevisions();
  const revision = await prepareReviewWrite("office_manage_revisions", {
    action: "accept",
    snapshot_id: read.snapshot_id,
    indices: [0],
  });
  assert.equal((await revision.apply(controls)).verified, true);
});
test("OOXML pagination returns the first exact raw export while rechecking later equivalent exports", async () => {
  const f = fixture();
  let exports = 0;
  f.main.getOoxml = () => ({ value: nativeOoxml.target[Math.floor(exports++ / 2) % 2] });
  const first = await toolReadFormatting({ ooxml_limit: 100 });
  const rest = await toolReadFormatting({
    snapshot_id: first.snapshot_id,
    ooxml_offset: 100,
    ooxml_limit: 100000,
  });
  assert.equal(first.ooxml + rest.ooxml, nativeOoxml.target[0]);
  assert.equal(rest.ooxml_length, nativeOoxml.target[0].length);
});
test("deleting and recreating identical text with a different native paragraph identity invalidates the proposal", async () => {
  const f = fixture();
  let exports = 0;
  f.main.getOoxml = () => ({ value: nativeOoxml.body[exports++ % 2] });
  const p = await prepareReviewWrite("office_edit_range", {
    action: "format",
    font: { bold: true },
  });
  f.main.paragraphs.items[0].uniqueLocalId = "P-recreated";
  await assert.rejects(p.apply(controls), /STALE_PROPOSAL/);
  assert.equal(f.writes, 0);
});
test("revision authors/dates and comment anchor identities cannot change across confirmation", async () => {
  for (const field of ["author", "date"]) {
    const f = fixture();
    const r = f.addRevision("Added", "100");
    const read = await toolReadRevisions();
    const p = await prepareReviewWrite("office_manage_revisions", {
      action: "accept",
      snapshot_id: read.snapshot_id,
      indices: [0],
    });
    r[field] = field === "author" ? "Different author" : new Date("2026-09-12");
    await assert.rejects(p.apply(controls), /STALE_PROPOSAL/);
    assert.equal(f.writes, 0);
  }
  const f = fixture();
  const c = f.comment(f.main, "Review", "C1");
  f.main.comments.push(c);
  const p = await prepareReviewWrite("office_manage_comment", {
    action: "edit",
    comment_id: "C1",
    text: "Changed",
  });
  c.getRange = () => f.range(f.body(f.main.text, "Different-anchor"));
  await assert.rejects(p.apply(controls), /STALE_PROPOSAL/);
  assert.equal(f.writes, 0);
});

test("stale diagnostics identify only fixed schema fields and never include document text", async () => {
  const f = fixture();
  const p = await prepareReviewWrite("office_edit_range", {
    action: "format",
    font: { bold: true },
  });
  f.main.text = "CONFIDENTIAL DOCUMENT CONTENT";
  await assert.rejects(p.apply(controls), (error) => {
    assert.match(error.message, /STALE_PROPOSAL.*snapshot.current/);
    assert.ok(!error.message.includes("CONFIDENTIAL"));
    assert.ok(!error.message.includes("USD"));
    return true;
  });
  assert.equal(f.writes, 0);
});

test("unexpected remaining revision changes stay unverified and diagnostics never disclose authors or text", async () => {
  const f = fixture();
  const selected = f.addRevision("Added", "PRIVATE TEXT");
  const remaining = f.addRevision("Formatted", "PRIVATE OTHER TEXT");
  const accept = selected.accept;
  selected.accept = () => {
    accept.call(selected);
    remaining.author = "PRIVATE AUTHOR";
  };
  const read = await toolReadRevisions();
  const plan = await prepareReviewWrite("office_manage_revisions", {
    action: "accept",
    snapshot_id: read.snapshot_id,
    indices: [0],
  });
  await assert.rejects(plan.apply(controls), (error) => {
    assert.match(
      error.message,
      /RESULT_UNKNOWN.*snapshot\[0\].author.*expected_count=1; actual_count=1/,
    );
    assert.ok(!error.message.includes("PRIVATE"));
    return true;
  });
});

test("native paired replacement exposes but never accepts the hidden deletion, with full OOXML proof", async () => {
  const data = JSON.parse(
    readFileSync(new URL("./fixtures/word-native-accepted-insertion.json", import.meta.url)),
  );
  const f = fixture();
  let after = false;
  f.main.getOoxml = () => ({ value: after ? data.after : data.before });
  const makeRevision = (row) => ({
    author: row.author,
    date: new Date(row.date),
    type: row.type,
    text: row.text,
    getRange() {
      return {
        ...f.range(f.main),
        paragraphs: f.list(() => row.paragraph_ids.map((uniqueLocalId) => ({ uniqueLocalId }))),
        getReviewedText: (version) => ({
          value: version === "Original" ? row.original : row.current,
        }),
        getOoxml: () => ({
          value: row.type === "Added" ? data.selected.ooxml : after ? data.after : data.before,
        }),
      };
    },
    accept() {
      assert.equal(row.type, "Added", "only the explicit insertion is accepted");
      after = true;
      f.main.revisions = data.after_rows.map(makeRevision);
    },
  });
  f.main.revisions = data.before_rows.map(makeRevision);
  const read = await toolReadRevisions();
  assert.equal(read.enumeration_complete, false);
  const plan = await prepareReviewWrite("office_manage_revisions", {
    action: "accept",
    snapshot_id: read.snapshot_id,
    indices: [2],
  });
  const result = await plan.apply(controls);
  assert.equal(result.verified, true);
  assert.equal(result.remaining, 3);
  assert.equal(result.api_revision_list_changed, true);
  assert.match(result.verification, /full_ooxml_and_Current/);
  assert.equal(f.main.revisions.at(-1).type, "Deleted");
});

test("read-only redline audit keeps authors, empty deletions and coverage limits without writing", async () => {
  const f = fixture();
  const deleted = f.addRevision("Deleted", "");
  const added = f.addRevision("Added", "100");
  deleted.author = "甲";
  added.author = "乙";
  const comment = f.comment(f.main, "请确认条款", "C1");
  comment.authorName = "甲";
  comment.resolved = true;
  comment.reply("已看");
  f.main.comments.push(comment);
  const beforeRevisions = f.main.revisions.length;
  const writes = f.writes;
  const audit = await toolAuditRedline({
    submitter: "Filing Clerk",
    file_sha256: "ab".repeat(32),
    anchor_queries: ["USD 100"],
  });
  assert.equal(f.writes, writes);
  assert.equal(f.main.revisions.length, beforeRevisions);
  assert.equal(comment.resolved, true);
  assert.equal(audit.office_writes, 0);
  assert.equal(audit.model_invoked, false);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.verified_state, false);
  assert.equal(audit.usable_for_writes, false);
  assert.equal(audit.claims_no_revisions, false);
  assert.equal(audit.snapshot.source.kind, "live_word");
  assert.equal(audit.snapshot.source.submitter, "Filing Clerk");
  assert.equal(audit.snapshot.source.file_sha256, null);
  assert.equal(audit.snapshot.source.hash_scope, "collected_evidence_only");
  assert.equal(audit.snapshot.consistency.full_file_consistent, false);
  assert.equal(audit.snapshot.consistency.evidence_probe_matched, true);
  assert.equal(audit.revisions.length, 2);
  const empty = audit.revisions.find((row) => row.native_type === "Deleted");
  assert.equal(empty.native_text, "");
  assert.equal(empty.mapping_status, "partial_native_pair");
  assert.deepEqual(audit.revisions.map((row) => row.word_author).sort(), ["乙", "甲"]);
  assert.ok(audit.revisions.every((row) => row.word_author !== "Filing Clerk"));
  assert.equal(audit.summary.comment_threads, 1);
  assert.equal(audit.summary.replies, 1);
  assert.equal(audit.comments[0].resolved, true);
  assert.equal(audit.comments[0].replies[0].text, "已看");
  assert.equal(audit.mappings[0].enumeration_complete, false);
  assert.equal(audit.mappings[0].count_equality_is_completeness, false);
  assert.equal(
    audit.coverage.entries.some((row) => row.status === "confirmed_absent"),
    false,
  );
  const footer = audit.coverage.entries.find(
    (row) => row.scope_label === "footer / section 0 / Primary",
  );
  assert.equal(footer.status, "not_read");
  assert.equal(footer.absence_confirmed, false);
  assert.equal(
    audit.anchors.find((row) => row.view === "Current" && row.query === "USD 100").status,
    "found",
  );
  const currentPage = audit.coverage.entries.find(
    (row) =>
      row.capability === "story_text" && row.version === "Current" && row.scope_label === "body",
  );
  assert.ok(currentPage.evidence_ids.length > 0);
  assert.ok(
    audit.story_pages.some(
      (page) =>
        page.evidence_id === currentPage.evidence_ids[0] &&
        page.snapshot_id === audit.snapshot_id &&
        page.version === "Current" &&
        page.paragraphs[0].text === f.main.text,
    ),
  );
  await assert.rejects(
    prepareReviewWrite("office_manage_revisions", {
      action: "accept",
      snapshot_id: audit.snapshot_id,
      indices: [0],
    }),
    /expired/,
  );
  assert.equal(f.writes, writes);
});

test("redline audit reports an unfinished story page instead of a coverage pass", async () => {
  const f = fixture();
  const second = f.body("第二条 同意。", "P2");
  const third = f.body("第三条 同意。", "P3");
  const paragraphs = [
    f.main.paragraphs.items[0],
    second.paragraphs.items[0],
    third.paragraphs.items[0],
  ];
  f.main.paragraphs = f.list(() => paragraphs);
  const audit = await toolAuditRedline({
    story_limit: 1,
    max_story_pages: 1,
    anchor_queries: ["同意。"],
  });
  const current = audit.coverage.entries.find(
    (row) =>
      row.capability === "story_text" && row.version === "Current" && row.scope_label === "body",
  );
  assert.equal(current.paging_complete, false);
  assert.equal(current.next_cursor, 1);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.summary.coverage_pass, false);
  assert.equal(audit.snapshot.consistency.evidence_probe_matched, true);
  const anchor = audit.anchors.find((row) => row.view === "Current" && row.query === "同意。");
  assert.equal(anchor.status, "not_found");
  assert.match(anchor.reason, /Unread pages were not searched/);
  assert.equal(f.writes, 0);
});

test("a changed document expires the prior audit token without writing or replaying it", async () => {
  const f = fixture();
  const first = await toolAuditRedline({ source_id: "synthetic-doc" });
  f.main.text = "changed after the audit";
  const second = await toolAuditRedline({ prior_snapshot_id: first.snapshot_id });
  assert.equal(second.snapshot.consistency.prior_snapshot.status, "not_current");
  assert.equal(second.snapshot.consistency.prior_snapshot.used_as_verified_state, false);
  assert.equal(second.verified_state, false);
  assert.notEqual(second.snapshot_id, first.snapshot_id);
  const missing = await toolAuditRedline({ prior_snapshot_id: "missing-token" });
  assert.equal(missing.snapshot.consistency.prior_snapshot.status, "expired");
  assert.equal(missing.snapshot.consistency.prior_snapshot.used_as_verified_state, false);
  assert.ok(
    missing.findings.some(
      (finding) => finding.category === "stale_or_inconsistent" && /expired/.test(finding.detail),
    ),
  );
  assert.equal(f.writes, 0);
});

test("an unsupported story stays an explicit audit gap and does not become zero revisions", async () => {
  const f = fixture({ notes: false });
  const audit = await toolAuditRedline({ target: { story: "footnote", note: 0 } });
  const revisions = audit.coverage.entries.find((row) => row.capability === "revisions");
  assert.equal(revisions.status, "unsupported");
  assert.match(revisions.reason, /UNSUPPORTED_API/);
  assert.match(revisions.reason, /does not establish that revisions are absent/);
  assert.equal(audit.summary.native_revision_rows, null);
  assert.equal(audit.revisions.length, 0);
  assert.equal(audit.claims_no_revisions, false);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(f.writes, 0);
});

test("optional range OOXML paging stays incomplete and still performs no write", async () => {
  const f = fixture();
  const audit = await toolAuditRedline({
    include_ooxml: true,
    ooxml_limit: 10,
    max_ooxml_pages: 1,
  });
  const xml = audit.coverage.entries.find((row) => row.capability === "range_ooxml");
  assert.equal(xml.status, "read");
  assert.equal(xml.paging_complete, false);
  assert.ok(xml.evidence_ids.length > 0);
  const chunk = audit.ooxml_evidence.find((item) => item.evidence_id === xml.evidence_ids[0]);
  assert.equal(chunk.snapshot_id, audit.snapshot_id);
  assert.equal(typeof chunk.ooxml, "string");
  assert.ok(chunk.ooxml.length > 0);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(f.writes, 0);
});

test("cancelling an audit starts no further Word.run and a later queued task still runs", async () => {
  const f = fixture();
  const queue = createOperationQueue();
  const original = globalThis.Word.run;
  let calls = 0;
  globalThis.Word.run = async (fn) => {
    calls += 1;
    if (calls === 1) queue.cancel();
    return original(fn);
  };
  await assert.rejects(
    queue.run(
      "audit-cancel",
      { mutates: false, deadline: () => Date.now() + 60_000 },
      ({ guard }) => toolAuditRedline({}, { guard }),
    ),
    /CANCELLED_OR_EXPIRED/,
  );
  assert.equal(calls, 1);
  assert.equal(f.writes, 0);
  const later = await queue.run(
    "after-cancel",
    { mutates: false, deadline: () => Date.now() + 60_000 },
    async () => (await toolReadStory({ limit: 1 })).paragraphs[0].text,
  );
  assert.equal(later, f.main.text);
  assert.ok(calls > 1);
});
