import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toolReadParagraphs,
  toolInsertParagraphs,
  toolReplaceSection,
  toolReplaceParagraphs,
  toolClearHighlights,
  toolClearComments,
  toolHeaderFooter,
  toolSetTableCell,
  prepareSafeWordWrite,
  prepareTextReplacement,
  toolApplyTextReplacement,
  resolveWordHeadingArgs,
  toolInsertHyperlink,
  toolInsertBookmark,
  toolHighlight,
  toolGetOutline,
  toolSetFont,
  toolApplyStyle,
  toolSetList,
  toolSetParagraphFormatting,
  captureWordFormattingPrecondition,
} from "../taskpane/shared/tools-word.js";
import {
  assertTextTargetSafe,
  assertNoCrossParagraphFields,
  assertLiteralSearchText,
} from "../taskpane/shared/word-write-safety.js";
import { paragraphFieldOpenAfter } from "../taskpane/shared/word-ooxml-state.js";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const escape = (s) => s.replaceAll("&", "&amp;").replaceAll("<", "&lt;");
const xml = (text, rich = "") =>
  `<w:p xmlns:w="${W}">${rich}<w:r><w:t>${escape(text)}</w:t></w:r></w:p>`;

// Office-shaped synthetic host: deterministic lifecycle proof, not rendering.
function fixture(
  rows = [{ text: "widget widgets" }],
  { wrongReadback = false, serializedTracking = false } = {},
) {
  let writes = 0;
  const list = (get) => ({
    load() {},
    get items() {
      return get();
    },
  });
  const paras = rows.map((row, i) => ({
    text: row.raw ?? row.text,
    current: row.text,
    rich: row.rich ?? "",
    style: row.style ?? "Normal",
    styleBuiltIn: row.styleBuiltIn,
    uniqueLocalId: `ID${i}`,
    listItemOrNullObject: { isNullObject: true },
    load() {},
  }));
  const write = (p, text) => {
    writes++;
    p.current = wrongReadback ? "unexpected" : text;
    p.text = p.current;
  };
  function range(p, start = 0, end = p.current.length) {
    return {
      _p: p,
      _start: start,
      _end: end,
      font: {},
      load() {},
      select() {},
      get text() {
        return p.current.slice(start, end);
      },
      getReviewedText: () => ({ value: p.current.slice(start, end) }),
      getOoxml: () => ({ value: xml(p.current, p.rich) }),
      expandTo: (r) => range(p, start, r._end),
      insertText: (replacement) =>
        write(p, p.current.slice(0, start) + replacement + p.current.slice(end)),
    };
  }
  function attach(p) {
    p.getRange = (location) => range(p, 0, location === "Start" ? 0 : p.current.length);
    p.insertText = (text) => write(p, text);
    p.delete = () => {
      writes++;
      paras.splice(paras.indexOf(p), 1);
    };
    p.insertParagraph = (text) => {
      writes++;
      const next = { ...p, uniqueLocalId: `NEW${writes}`, current: text, text, rich: "" };
      attach(next);
      paras.splice(paras.indexOf(p) + 1, 0, next);
      return next;
    };
    p.search = (find, opts) => {
      const hits = [];
      const source = opts.matchCase ? p.current : p.current.toLowerCase();
      const needle = opts.matchCase ? find : find.toLowerCase();
      let at = source.indexOf(needle);
      while (at >= 0) {
        if (
          !opts.matchWholeWord ||
          (!/[a-z0-9]/i.test(source[at - 1] ?? "") &&
            !/[a-z0-9]/i.test(source[at + needle.length] ?? ""))
        )
          hits.push(range(p, at, at + needle.length));
        at = source.indexOf(needle, at + needle.length);
      }
      return list(() => hits);
    };
  }
  paras.forEach(attach);
  const paragraphList = list(() => paras);
  const body = {
    paragraphs: paragraphList,
    load() {},
    getOoxml: () => ({
      value: `<w:body xmlns:w="${W}">${paras.map((p) => xml(p.current, p.rich)).join("")}</w:body>`,
    }),
    getComments: () => list(() => []),
  };
  function container(text) {
    return {
      current: text,
      rich: "",
      load() {},
      paragraphs: list(() => [{ uniqueLocalId: "container-" + text }]),
      getOoxml() {
        const raw = xml(this.current, this.rich);
        return {
          value: serializedTracking
            ? `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W}"><w:body>${raw}</w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><w:settings xmlns:w="${W}">${document.changeTrackingMode === "TrackAll" ? "<w:trackRevisions/>" : ""}</w:settings></pkg:xmlData></pkg:part></pkg:package>`
            : raw,
        };
      },
      getReviewedText() {
        return { value: this.current };
      },
      clear() {
        writes++;
        this.current = "";
      },
      insertText(text) {
        writes++;
        this.current = wrongReadback ? "unexpected" : text;
      },
    };
  }
  const header = container("Initial header"),
    cell = container("Cell content");
  body.tables = list(() => [{ getCell: () => ({ body: cell }) }]);
  const document = {
    body,
    changeTrackingMode: "Off",
    load() {},
    sections: list(() => [{ getHeader: () => header, getFooter: () => header }]),
    getSelection: () => ({ select() {} }),
  };
  globalThis.Word = {
    run: (fn) => fn({ document, sync: async () => {} }),
    ChangeTrackingVersion: { current: "Current" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    InsertLocation: { replace: "Replace", after: "After", start: "Start" },
    HeaderFooterType: { primary: "Primary" },
  };
  return {
    paras,
    header,
    cell,
    body,
    get writes() {
      return writes;
    },
  };
}

test("all generic formatting channels refuse OOXML-only edits after approval, before any write", async () => {
  const cases = [
    ["office_set_font", toolSetFont, { ids: ["ID0"], bold: true }],
    ["office_apply_style", toolApplyStyle, { ids: ["ID0"], style: "Heading 1" }],
    [
      "office_set_paragraph_formatting",
      toolSetParagraphFormatting,
      { ids: ["ID0"], alignment: "center" },
    ],
    ["office_highlight", toolHighlight, { targets: [{ paragraph_id: "ID0", query: "widget" }] }],
    ["office_clear_highlights", toolClearHighlights, { paragraph_ids: ["ID0"] }],
    ["office_set_list", toolSetList, { ids: ["ID0"], ordered: false }],
  ];
  for (const [name, tool, args] of cases) {
    const f = fixture();
    const expected = await captureWordFormattingPrecondition(name, args);
    f.paras[0].rich = "<w:r><w:rPr><w:i/></w:rPr></w:r>";
    await assert.rejects(
      tool(args, {
        formatExpected: expected,
        guard() {},
        beforeWrite() {
          assert.fail(`${name} attempted a stale write`);
        },
      }),
      /STALE_PROPOSAL/,
    );
    assert.equal(f.writes, 0);
    assert.equal(f.paras[0].current, "widget widgets");
  }
});

test("format snapshots detect style-only changes but do not reject rich text or unrelated formatting", async () => {
  const f = fixture([
    { text: "target", rich: "<w:r><w:rPr><w:b/></w:rPr></w:r>" },
    { text: "other" },
  ]);
  const args = { ids: ["ID0"], bold: true };
  const expected = await captureWordFormattingPrecondition("office_set_font", args);
  f.paras[1].rich = '<w:pPr><w:jc w:val="center"/></w:pPr>';
  let writes = 0;
  await toolSetFont(args, {
    formatExpected: expected,
    guard() {},
    beforeWrite() {
      writes++;
    },
  });
  assert.equal(writes, 1);
  f.paras[0].style = "Heading 1";
  await assert.rejects(
    toolSetFont(args, {
      formatExpected: expected,
      beforeWrite() {
        assert.fail();
      },
    }),
    /STALE_PROPOSAL/,
  );
  assert.equal(await captureWordFormattingPrecondition("office_get_selection", {}), null);
});

test("formatting recheck preserves native paragraph references without reloading the collection", async () => {
  const f = fixture();
  const args = { ids: ["ID0"], size: 13, track_changes: true };
  const expected = await captureWordFormattingPrecondition("office_set_font", args);
  let generation = 0;
  const paragraphs = f.body.paragraphs;
  f.body.paragraphs = {
    load() {
      generation++;
    },
    get items() {
      const captured = generation;
      return paragraphs.items.map(
        (paragraph) =>
          new Proxy(paragraph, {
            get(target, name) {
              if (name === "getRange")
                return (...values) => {
                  assert.equal(captured, generation, "native paragraph reference was invalidated");
                  return target.getRange(...values);
                };
              return Reflect.get(target, name);
            },
          }),
      );
    },
  };
  let writes = 0;
  await toolSetFont(args, {
    formatExpected: expected,
    guard() {},
    beforeWrite() {
      writes++;
    },
  });
  assert.equal(generation, 1);
  assert.equal(writes, 1);
});

test("inner formatting gate catches changes during tool preparation after the pane recheck", async () => {
  const f = fixture();
  const args = { ids: ["ID0"], bold: true };
  const expected = await captureWordFormattingPrecondition("office_set_font", args);
  const run = Word.run;
  Word.run = (fn) =>
    run((context) => {
      const sync = context.sync;
      let calls = 0;
      context.sync = async () => {
        await sync();
        if (++calls === 1) f.paras[0].rich = '<w:pPr><w:spacing w:after="120"/></w:pPr>';
      };
      return fn(context);
    });
  await assert.rejects(
    toolSetFont(args, {
      formatExpected: expected,
      beforeWrite() {
        assert.fail();
      },
    }),
    /STALE_PROPOSAL/,
  );
  assert.equal(f.writes, 0);
});

test("targeted reads reject missing, repeated, mixed and invalid ranges without partial coverage", async () => {
  fixture();
  for (const args of [
    { ids: ["ID0", "missing"] },
    { ids: ["ID0", "ID0"] },
    { ids: [] },
    { ids: ["ID0"], range: [0, 1] },
    { range: [-1, 1] },
    { range: [1, 0] },
    { range: [0, 9] },
    { range: null },
  ])
    await assert.rejects(toolReadParagraphs(args), /INVALID_SELECTOR|PARAGRAPH_NOT_FOUND/);
});

test("duplicate headings reject every affected read/write route and allow explicit heading ID", async () => {
  const f = fixture([
    { text: "Background", style: "Heading 1" },
    { text: "first" },
    { text: "Background", style: "Heading 1" },
    { text: "second" },
  ]);
  const calls = [
    () => toolReadParagraphs({ heading_section: "Background" }),
    () => toolInsertParagraphs({ after: { heading: "Background" }, content: ["x"] }),
    () => toolReplaceSection({ heading: "Background", content: ["x"] }),
    () => toolClearHighlights({ heading_section: "Background" }),
    () => toolClearComments({ heading_section: "Background" }),
  ];
  for (const call of calls) await assert.rejects(call(), /AMBIGUOUS_HEADING.*ID0.*ID2/);
  assert.equal(f.writes, 0);
  const result = await toolReadParagraphs({ heading_id: "ID2" });
  assert.deepEqual(
    result.paragraphs.map((p) => p.text),
    ["Background", "second"],
  );
});

test("Current revised heading resolves before confirmation and is reused for section replacement", async () => {
  const f = fixture([
    { text: "New heading", raw: "Old headingNew heading", style: "Heading 1" },
    { text: "body" },
  ]);
  const resolved = await resolveWordHeadingArgs("office_insert_paragraphs", {
    after: { heading: "New heading" },
    content: ["addition"],
  });
  assert.deepEqual(resolved.after, { id: "ID0" });
  await toolInsertParagraphs(resolved);
  await toolClearComments({ heading_section: "New heading" });
  await toolClearHighlights({ heading_section: "New heading" });
  await toolReplaceSection({ heading: "New heading", content: ["new body"] });
  assert.deepEqual(
    f.paras.map((p) => p.current),
    ["New heading", "new body"],
  );
});

test("whole-word plan uses Word matches, keeps excluded substrings and rechecks before write", async () => {
  const f = fixture();
  const args = { paragraph_ids: ["ID0"], find: "widget", replace: "device", whole_word: true };
  const plan = await prepareTextReplacement(args);
  assert.equal(plan[0].after, "device widgets");
  await toolApplyTextReplacement(args, plan, { guard() {}, beforeWrite() {} });
  assert.equal(f.paras[0].current, "device widgets");
  assert.equal(f.writes, 1);
});

test("localized built-in heading labels use the locale-independent built-in level", async () => {
  fixture([
    { text: "本地化标题", style: "标题 1", styleBuiltIn: "Heading1" },
    { text: "章节正文" },
    { text: "下一标题", style: "标题 1", styleBuiltIn: "Heading1" },
  ]);
  const result = await toolReadParagraphs({ heading_section: "本地化标题" });
  assert.deepEqual(
    result.paragraphs.map((p) => p.text),
    ["本地化标题", "章节正文"],
  );
});

test("fields, review markup and direct-run formats are refused before broad mutation", async () => {
  for (const rich of [
    '<w:fldSimple w:instr="PAGE"/>',
    '<w:commentRangeStart w:id="0"/>',
    '<w:ins w:id="1"/>',
    "<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>",
  ]) {
    const f = fixture([{ text: "Protected", rich }]);
    await assert.rejects(
      toolReplaceParagraphs({ ids: ["ID0"], content: ["new"] }),
      /PROTECTED_FIELD|UNSUPPORTED_RICH_TARGET/,
    );
    assert.equal(f.writes, 0);
  }
});

test("multi-paragraph TOC field cannot be bypassed by targeting a marker-free middle paragraph", async () => {
  const f = fixture([
    { text: "TOC", rich: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' },
    { text: "Entry 1" },
    { text: "end", rich: '<w:r><w:fldChar w:fldCharType="end"/></w:r>' },
  ]);
  await assert.rejects(
    prepareTextReplacement({ paragraph_ids: ["ID1"], find: "Entry", replace: "Other" }),
    /PROTECTED_FIELD/,
  );
  assert.equal(f.writes, 0);
});

test("header target changes with unchanged body invalidate approved proposal", async () => {
  const f = fixture();
  const prepared = await prepareSafeWordWrite("office_header_footer", {
    which: "header",
    text: "new",
  });
  assert.equal(prepared.preview.targets[0].current_text, "Initial header");
  f.header.current = "Manual header change";
  await assert.rejects(
    prepared.apply({
      guard() {},
      beforeWrite() {
        assert.fail("no write");
      },
    }),
    /STALE_PROPOSAL/,
  );
  assert.equal(f.writes, 0);
});

test("broad write refuses a marker-free paragraph inside a cross-paragraph comment", async () => {
  const f = fixture([
    { text: "start", rich: '<w:commentRangeStart w:id="1"/>' },
    { text: "middle" },
    { text: "end", rich: '<w:commentRangeEnd w:id="1"/>' },
  ]);
  await assert.rejects(
    toolReplaceParagraphs({ ids: ["ID1"], content: ["changed"] }),
    /UNSUPPORTED_RICH_TARGET/,
  );
  assert.equal(f.writes, 0);
});

test("plain cell and header changes verify the target, and bad target readback is unknown", async () => {
  fixture();
  const result = await toolSetTableCell({ table_index: 0, row: 0, column: 0, text: "new" });
  assert.equal(result.verification, "target_current_text_readback");
  const f = fixture(undefined, { wrongReadback: true });
  await assert.rejects(toolHeaderFooter({ which: "header", text: "expected" }), /RESULT_UNKNOWN/);
  assert.equal(f.writes, 2);
});

test("prepared tracked broad write does not change serialized mode before checking its approved target", async () => {
  const f = fixture(undefined, { serializedTracking: true });
  const prepared = await prepareSafeWordWrite("office_set_table_cell", {
    table_index: 0,
    row: 0,
    column: 0,
    text: "tracked new cell",
    track_changes: true,
  });
  const result = await prepared.apply({ guard() {}, beforeWrite() {} });
  assert.equal(result.verification, "target_current_text_readback");
  assert.equal(f.cell.current, "tracked new cell");
  assert.equal(f.writes, 2);
});

test("namespace-aware field scan ignores escaped text/style parts and fails closed on malformed XML", () => {
  assert.doesNotThrow(() => assertTextTargetSafe(xml("<w:fldChar/>"), { broad: true }));
  assert.throws(
    () => assertTextTargetSafe('<w:p xmlns:w="wrong"><w:fldChar/></w:p>'),
    /UNAVAILABLE/,
  );
  assert.throws(() => assertTextTargetSafe(`<w:p xmlns:w="${W}"><w:r></w:p>`), /UNAVAILABLE/);
  assert.throws(
    () => assertNoCrossParagraphFields(xml("x", '<w:fldChar w:fldCharType="end"/>')),
    /PROTECTED_FIELD/,
  );
});

test("hyperlink and bookmark queries must match exactly once", async () => {
  const f = fixture([{ text: "the fee and the term" }]);
  const [p] = f.paras;
  const bookmarks = [];
  const origSearch = p.search;
  p.search = (find, opts) => {
    const hits = origSearch(find, opts);
    for (const r of hits.items) r.insertBookmark = (name) => bookmarks.push([r.text, name]);
    return hits;
  };
  await assert.rejects(
    toolInsertHyperlink({ paragraph_id: "ID0", query: "the", url: "https://example.com" }),
    /AMBIGUOUS_RANGE: query has 2 matches/,
  );
  await assert.rejects(
    toolInsertBookmark({ paragraph_id: "ID0", query: "the", name: "b1" }),
    /AMBIGUOUS_RANGE/,
  );
  assert.deepEqual(bookmarks, []);
  const linked = await toolInsertHyperlink({
    paragraph_id: "ID0",
    query: "fee",
    url: "https://example.com",
  });
  assert.equal(linked.anchored_on, "fee");
  await toolInsertBookmark({ paragraph_id: "ID0", query: "term", name: "b2" });
  assert.deepEqual(bookmarks, [["term", "b2"]]);
});

test("hyperlinks reject non-web schemes before touching Word", async () => {
  const f = fixture([{ text: "click here" }]);
  for (const url of ["javascript:alert(1)", "file:///etc/passwd", "vbscript:x", "data:text/html,x"])
    await assert.rejects(
      toolInsertHyperlink({ paragraph_id: "ID0", query: "here", url }),
      /UNSUPPORTED_URL/,
    );
  for (const url of ["http://a.test", "HTTPS://a.test/x", "mailto:a@b.test", "#section-2"])
    await toolInsertHyperlink({ paragraph_id: "ID0", query: "here", url });
  assert.equal(f.writes, 0);
});

test("field scan reports insertion points inside multi-paragraph fields only", () => {
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const body = (...ps) => `<w:body xmlns:w="${W}">${ps.join("")}</w:body>`;
  assert.deepEqual(
    paragraphFieldOpenAfter(
      body(xml("before"), xml("toc 1", begin), xml("toc 2"), xml("toc 3", end), xml("after")),
    ),
    [false, true, true, false, false],
  );
  // A complete inline field and a text-box field do not open the paragraph.
  const textBox = `<w:r><w:drawing><w:txbxContent><w:p>${begin}</w:p></w:txbxContent></w:drawing></w:r>`;
  assert.deepEqual(paragraphFieldOpenAfter(body(xml("page", begin + end), xml("box", textBox))), [
    false,
    false,
  ]);
  assert.equal(paragraphFieldOpenAfter(body(xml("dangling", begin))), null);
  assert.equal(paragraphFieldOpenAfter(body(xml("orphan end", end))), null);
});

test("paragraph insertion refuses anchors inside a TOC and allows anchors after it", async () => {
  const f = fixture([
    { text: "TOC entry 1", rich: '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' },
    { text: "TOC entry 2" },
    { text: "TOC entry 3", rich: '<w:r><w:fldChar w:fldCharType="end"/></w:r>' },
    { text: "Body" },
  ]);
  for (const id of ["ID0", "ID1"])
    await assert.rejects(
      toolInsertParagraphs({ after: { id }, content: ["new"] }),
      /PROTECTED_FIELD: the insertion point is inside/,
    );
  assert.equal(f.writes, 0);
  await toolInsertParagraphs({ after: { id: "ID2" }, content: ["new"] });
  assert.deepEqual(
    f.paras.map((p) => p.current),
    ["TOC entry 1", "TOC entry 2", "TOC entry 3", "new", "Body"],
  );
});

test("write searches reject Word special codes and over-long text before touching Word", async () => {
  assert.doesNotThrow(() => assertLiteralSearchText("“Agreement” (as amended)"));
  assert.throws(() => assertLiteralSearchText("x".repeat(256)), /SEARCH_TEXT_TOO_LONG/);
  assert.doesNotThrow(() => assertLiteralSearchText("x".repeat(255)));
  const f = fixture([{ text: "cost^tbasis and cost\tbasis" }]);
  const calls = [
    () => prepareTextReplacement({ paragraph_ids: ["ID0"], find: "cost^tbasis", replace: "x" }),
    () => toolHighlight({ targets: [{ paragraph_id: "ID0", query: "^#" }] }),
    () => toolInsertHyperlink({ paragraph_id: "ID0", query: "^p", url: "https://a.test" }),
    () => toolInsertBookmark({ paragraph_id: "ID0", query: "cost^t", name: "b" }),
  ];
  for (const call of calls) await assert.rejects(call(), /UNSUPPORTED_SEARCH_TEXT/);
  await assert.rejects(
    prepareTextReplacement({ paragraph_ids: ["ID0"], find: "c".repeat(300), replace: "x" }),
    /SEARCH_TEXT_TOO_LONG/,
  );
  assert.equal(f.writes, 0);
});

test("outline uses the built-in heading identity, including a localized style name", async () => {
  fixture([{ text: "Synthetic heading", style: "标题 1", styleBuiltIn: "Heading1" }]);
  const outline = await toolGetOutline();
  assert.equal(outline.heading_count, 1);
  assert.equal(outline.outline[0].level, 1);
  assert.equal(outline.outline[0].text, "Synthetic heading");
  const section = await toolReadParagraphs({ heading_section: "Synthetic heading" });
  assert.equal(section.paragraphs.length, 1);

  fixture([{ text: "Custom", style: "Acme-Heading 2", styleBuiltIn: "Normal" }]);
  assert.equal((await toolGetOutline()).outline[0].level, 2);

  fixture([{ text: "Only localized", style: "标题 1", styleBuiltIn: "Normal" }]);
  assert.equal((await toolGetOutline()).heading_count, 0);
});

test("highlight reports a partial result and does not call a total miss a success", async () => {
  const f = fixture([{ text: "Synthetic highlight" }]);
  let paints = 0;
  const getRange = f.paras[0].getRange;
  f.paras[0].getRange = (...args) => {
    const range = getRange(...args);
    return {
      ...range,
      font: {
        set highlightColor(value) {
          paints++;
          range.font.highlightColor = value;
        },
      },
    };
  };
  const partial = await toolHighlight({
    targets: [{ paragraph_id: "ID0" }, { paragraph_id: "missing" }],
  });
  assert.equal(partial.status, "partial");
  assert.equal(partial.applied.length, 1);
  assert.equal(partial.not_applied[0].error, "paragraph_not_found");
  assert.equal(paints, 1);
  await assert.rejects(
    toolHighlight({ targets: [{ paragraph_id: "missing" }] }),
    /HIGHLIGHT_FAILED/,
  );
});

test("font, style, list and highlight do not mutate after a cancel in the load sync", async () => {
  async function rejectDuringFirstSync(run) {
    const original = Word.run;
    let release;
    const gate = new Promise((resolve) => {
      release = resolve;
    });
    let started;
    const entered = new Promise((resolve) => {
      started = resolve;
    });
    let count = 0;
    let cancelled = false;
    Word.run = (fn) =>
      original((ctx) =>
        fn({
          ...ctx,
          sync: async () => {
            if (++count === 1) {
              started();
              await gate;
            }
          },
        }),
      );
    const control = {
      guard() {
        if (cancelled) throw new Error("CANCELLED_OR_EXPIRED");
      },
      beforeWrite() {
        throw new Error("write marked after cancel");
      },
    };
    const pending = run(control);
    await entered;
    cancelled = true;
    release();
    await assert.rejects(pending, /CANCELLED_OR_EXPIRED/);
    Word.run = original;
  }

  const host = fixture([{ text: "Synthetic cancellation target" }]);
  let paints = 0;
  host.paras[0].getRange = () => ({
    font: {
      set bold(value) {
        paints++;
        void value;
      },
    },
    load() {},
  });
  let styled = 0;
  Object.defineProperty(host.paras[0], "style", {
    configurable: true,
    set() {
      styled++;
    },
  });
  let lists = 0;
  host.paras[0].startNewList = () => {
    lists++;
    return { id: "L", load() {} };
  };

  await rejectDuringFirstSync((control) =>
    toolSetFont({ ids: ["ID0"], bold: true, track_changes: true }, control),
  );
  assert.equal(paints, 0);
  await rejectDuringFirstSync((control) =>
    toolApplyStyle({ ids: ["ID0"], style: "Heading 1", track_changes: true }, control),
  );
  assert.equal(styled, 0);
  await rejectDuringFirstSync((control) =>
    toolSetList({ ids: ["ID0"], ordered: false, track_changes: true }, control),
  );
  assert.equal(lists, 0);

  let highlights = 0;
  host.paras[0].getRange = () => ({
    font: {
      set highlightColor(value) {
        highlights++;
        void value;
      },
    },
    load() {},
  });
  await rejectDuringFirstSync((control) =>
    toolHighlight({ targets: [{ paragraph_id: "ID0" }] }, control),
  );
  assert.equal(highlights, 0);
});

test("section replacement rechecks authorization before the insertion batch", async () => {
  for (const track_changes of [false, true]) {
    for (const reason of ["cancel", "expiry", "document-change"]) {
      const host = fixture([
        { text: "Target", style: "Heading 1", styleBuiltIn: "Heading1" },
        { text: "Old body" },
        { text: "Next", style: "Heading 1", styleBuiltIn: "Heading1" },
      ]);
      const original = Word.run;
      const document = await Word.run((context) => context.document);
      const receipts = new Map();
      const queue = createOperationQueue({
        storage: {
          getItem: (key) => receipts.get(key),
          setItem: (key, value) => receipts.set(key, value),
        },
      });
      let deadline = Date.now() + 60000;
      let valid = true;
      let revoked = false;
      Word.run = (fn) =>
        original((context) =>
          fn({
            ...context,
            sync: async () => {
              if (!revoked && host.writes > 0) {
                revoked = true;
                if (reason === "cancel") queue.cancel();
                else if (reason === "expiry") deadline = 0;
                else valid = false;
              }
            },
          }),
        );
      try {
        await assert.rejects(
          queue.run(
            "section",
            { mutates: true, deadline: () => deadline, valid: () => valid },
            (control) =>
              toolReplaceSection(
                {
                  heading_id: "ID0",
                  content: ["Replacement"],
                  style_per_para: ["Normal"],
                  track_changes,
                },
                control,
              ),
          ),
          /RESULT_UNKNOWN:.*CANCELLED_OR_EXPIRED/,
        );
        assert.equal(host.writes, 1, `${reason}: no new insertion after the committed delete`);
        assert.deepEqual(
          host.paras.map((p) => p.current),
          ["Target", "Next"],
        );
        assert.equal(receipts.get("draftspect-operation:section"), "attempted");
        assert.equal(
          document.changeTrackingMode,
          "Off",
          "restore the original tracking preference",
        );
      } finally {
        Word.run = original;
      }
    }
  }
});
