/* global Word */
import { assertPreconditions, planReplacement } from "./word-edit-plan.js";
import { toolReadComments } from "./word-review.js";
import { stableOoxml, stableCellOoxml, paragraphFieldOpenAfter } from "./word-ooxml-state.js";
import {
  assertTextTargetSafe,
  assertNoCrossParagraphFields,
  assertDistinctIds,
  assertOneSelector,
  assertLiteralSearchText,
} from "./word-write-safety.js";

const BODY_COVERAGE = {
  scope: "body",
  text_version: "current",
  authoritative: true,
  excluded: [
    "headers",
    "footers",
    "footnotes",
    "endnotes",
    "comments",
    "text_boxes",
    "embedded_objects",
    "original_revision_text",
    "visual_layout",
  ],
};
//
// Word tool implementations + Word-specific helpers.
//
// Each `tool*` function is invoked from the dispatcher in taskpane.js when a
// matching `office_*` tool_call arrives from the daemon. They all run inside
// Word.run(...) so Office.js manages the request context lifecycle. The
// helpers above (paragraph addressing, withTrackChanges, snapshot, heading
// detection) are private to this module — taskpane.js doesn't need them.

// ---------------------------------------------------------------------------
// Paragraph addressing
//
// We use Word's session-local `paragraph.uniqueLocalId` — stable
// across insertions/deletions, which is what the agent needs when it does
// "read these paragraphs, then edit p7" across multiple tool calls. Available
// in WordApi 1.6+. It is NOT OOXML w:paraId: it changes across Word sessions
// and coauthors, so archived IDs must never be reused without a fresh read.
//
// On older Word builds where uniqueLocalId is unsupported, we fall back to
// position-based IDs (`p${index}`). These are NOT stable across structural
// edits — when the fallback is in play, snapshotParagraphs() flags the mode
// and tools should re-read before each operation rather than caching IDs.
// ---------------------------------------------------------------------------
const fallbackId = (index) => `p${index}`;
const parseFallbackId = (id) => {
  const m = /^p(\d+)$/.exec(id);
  return m ? parseInt(m[1], 10) : null;
};

async function getParagraphsWithIds(context) {
  const paragraphs = context.document.body.paragraphs;
  // CRITICAL: load every per-item property we'll need in a SINGLE
  // paragraphs.load() before the first sync. A second load/sync on the same
  // paragraphs collection silently un-loads the previously-loaded scalars,
  // so every later `p.text` read throws "property 'text' is not available."
  // changeTrackingMode is a document-level scalar, not a paragraph property —
  // safe to piggyback on the same sync, no scalar-refresh issue.
  try {
    context.document.load("changeTrackingMode");
    paragraphs.load(
      "items/text, items/style, items/styleBuiltIn, items/uniqueLocalId, items/listItemOrNullObject/listString",
    );
    await context.sync();
    if (paragraphs.items.length > 0 && !paragraphs.items[0].uniqueLocalId) {
      throw new Error("uniqueLocalId not populated");
    }
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";
    return { paragraphs, idMode: "uniqueLocalId", trackChangesOn };
  } catch {
    context.document.load("changeTrackingMode");
    paragraphs.load("items/text, items/style, items/styleBuiltIn");
    await context.sync();
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";
    return { paragraphs, idMode: "index", trackChangesOn };
  }
}

// Pull a literally-typed bracketed paragraph number (e.g. "[0034]") off the
// front of a paragraph's text. Fallback for docs that number paragraphs by
// typing, not by Word list-numbering.
function extractBracketedNumber(text) {
  if (!text) return null;
  const m = /^\s*(\[[^\]]+\])/.exec(text);
  return m ? m[1] : null;
}

// The user-facing number for a paragraph. Preferred: Word's rendered list
// number (e.g. "[0034]") via Paragraph.listItemOrNullObject.listString —
// NOT present in paragraph.text, so a regex on text alone misses every
// auto-numbered doc. Fallback: a literally-typed [NNNN] prefix in the text.
// Used so the agent can cite paragraph numbers the user actually sees on
// the page, instead of opaque uniqueLocalId hex strings.
function listNumberOf(p, text) {
  try {
    const li = p.listItemOrNullObject;
    if (li && !li.isNullObject) {
      const s = li.listString;
      if (s && String(s).trim()) return String(s).trim();
    }
  } catch {
    /* listItemOrNullObject not loaded on this build / fallback path */
  }
  return extractBracketedNumber(text);
}

// Run a batch of getReviewedText(...) calls on ranges, working around
// office-js#3691 (getReviewedText scrolls the document to the change). Save
// the user's selection once, do the reads, restore the selection. Returns
// an array of Current strings aligned with ranges, or an explicit error.
// Recording Off does not prove there are no unresolved revisions. Never
// silently substitute raw .text. office-js#5188 makes getTrackedChanges() unusable as an
// alternative (TrackedChange.text returns empty for Deleted), so this is
// the only supported path.
async function getReviewedTextBatch(context, ranges) {
  if (!ranges || ranges.length === 0) return [];
  const saved = context.document.getSelection();
  try {
    const results = ranges.map((r) => r.getReviewedText(Word.ChangeTrackingVersion.current));
    await context.sync();
    try {
      saved.select();
      await context.sync();
    } catch {
      /* best effort — restore failure shouldn't mask the read */
    }
    const text = results.map((r) => r.value);
    if (text.some((value) => typeof value !== "string"))
      throw new Error("Current text was not populated");
    return text;
  } catch {
    try {
      saved.select();
      await context.sync();
    } catch {
      /* best effort */
    }
    throw new Error(
      "REVIEWED_TEXT_UNAVAILABLE: Current text could not be read. Raw revision text was not substituted.",
    );
  }
}

function getId(paragraph, index, idMode) {
  return idMode === "uniqueLocalId" ? paragraph.uniqueLocalId : fallbackId(index);
}

function findIndexById(paragraphs, id, idMode) {
  if (idMode === "uniqueLocalId") {
    return paragraphs.items.findIndex((p) => p.uniqueLocalId === id);
  }
  const i = parseFallbackId(id);
  if (i === null || i < 0 || i >= paragraphs.items.length) return -1;
  return i;
}

// Run `body` with the document's changeTrackingMode set to TrackAll if
// `track_changes` is truthy, and unconditionally restore it in `finally` so
// that exceptions thrown mid-edit don't leave Word stuck in Track All mode.
// The restore-sync is wrapped in a try/catch so a failed restore doesn't
// shadow the original error.
async function withTrackChanges(context, track_changes, body, deferUntilWrite = false, control) {
  let prevTracking = null;
  if (track_changes) {
    context.document.load("changeTrackingMode");
    await context.sync();
    // The load sync is a gap. A cancel that arrived there must not turn
    // tracking on, and must not enter the edit.
    control?.guard();
    prevTracking = context.document.changeTrackingMode;
    if (deferUntilWrite) context.__draftspectPendingTrackingMode = Word.ChangeTrackingMode.trackAll;
    else context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  }
  try {
    return await body();
  } catch (error) {
    if (context.__draftspectWriteAttempted && !String(error.message).includes("RESULT_UNKNOWN"))
      throw new Error("RESULT_UNKNOWN: " + error.message);
    throw error;
  } finally {
    if (
      track_changes &&
      prevTracking !== null &&
      (!deferUntilWrite || context.__draftspectPendingTrackingMode === null)
    ) {
      try {
        context.document.changeTrackingMode = prevTracking;
        await context.sync();
      } catch (error) {
        throw new Error(
          `RESULT_UNKNOWN: tracking-mode restore failed; inspect Word. ${error.message}`,
        );
      }
    }
  }
}

const withProtectedTracking = (context, trackChanges, body, control) =>
  withTrackChanges(context, trackChanges, body, true, control);

// Check cancellation, mark the write, then apply a deferred tracking mode in
// the same queued batch as the mutation. Call this after the last preparatory
// sync and before any property that changes the document.
function armWrite(context, control) {
  control?.guard();
  control?.beforeWrite();
  context.__draftspectWriteAttempted = true;
  if (context.__draftspectPendingTrackingMode) {
    const nextMode = context.__draftspectPendingTrackingMode;
    context.__draftspectPendingTrackingMode = null;
    context.document.changeTrackingMode = nextMode;
  }
}

// ---------------------------------------------------------------------------
// Heading / section helpers
// ---------------------------------------------------------------------------
function headingLevelOf(paragraph) {
  return headingLevel(paragraph?.styleBuiltIn) ?? headingLevel(paragraph?.style);
}

function headingLevel(style) {
  // Returns the heading level if the style name *contains* "Heading <N>".
  // Many organizations use prefixed style names like "Acme-Heading 2" or
  // "Firm-Heading-1" — we want to recognize those as headings too, not just
  // Word's built-in "Heading 1" / "Heading 2".
  if (!style) return null;
  const m = /heading[\s_-]*(\d)/i.exec(style);
  return m ? parseInt(m[1], 10) : null;
}

// Build the "view" the agent sees: list of {id, style, text, [number]}.
// `number` is the user-facing paragraph number (Word's rendered list string,
// or a literally-typed [NNNN]) when present — omitted otherwise, so the
// agent's `if (p.number)` checks stay clean.
function snapshotParagraphs(paragraphs, idMode) {
  return paragraphs.items.map((p, i) => {
    const text = p.text;
    const number = listNumberOf(p, text);
    return {
      index: i,
      id: getId(p, i, idMode),
      style: p.style,
      heading_level: headingLevelOf(p),
      text,
      ...(number ? { number } : {}),
    };
  });
}

// Find the index of a heading paragraph by case-insensitive exact text match.
function findHeadingIndex(snapshot, headingText) {
  if (typeof headingText !== "string" || !headingText.trim())
    throw new Error("INVALID_SELECTOR: heading must be nonempty.");
  const target = headingText.trim().toLowerCase();
  const matches = snapshot.filter(
    (p) => p.heading_level !== null && p.text.trim().toLowerCase() === target,
  );
  if (matches.length > 1)
    throw new Error(
      `AMBIGUOUS_HEADING: choose a paragraph ID. Candidates: ${matches.map((p) => p.id).join(", ")}`,
    );
  return matches.length ? matches[0].index : -1;
}

async function currentSnapshot(context, paragraphs, idMode) {
  const snapshot = snapshotParagraphs(paragraphs, idMode);
  const current = await getReviewedTextBatch(
    context,
    paragraphs.items.map((p) => p.getRange()),
  );
  snapshot.forEach((p, i) => {
    p.text = current[i];
  });
  return snapshot;
}

function resolveHeading(snapshot, text, id) {
  assertOneSelector([text, id]);
  const index =
    id === undefined ? findHeadingIndex(snapshot, text) : snapshot.findIndex((p) => p.id === id);
  if (index < 0) throw new Error("Heading not found; reread the document.");
  if (snapshot[index].heading_level === null)
    throw new Error("INVALID_SELECTOR: selected paragraph is not a heading.");
  return index;
}

async function inspectTargets(context, targets, broad = true) {
  let xml;
  try {
    xml = targets.map((t) => t.range.getOoxml());
  } catch {
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: target OOXML is required before editing.");
  }
  const identities = targets.map((t) => {
    if (!/^(table|header|footer):/.test(t.key)) return null;
    if (!t.range.paragraphs)
      throw new Error(
        "TARGET_INSPECTION_UNAVAILABLE: container paragraph identities are required.",
      );
    t.range.paragraphs.load("items/uniqueLocalId");
    return t.range.paragraphs;
  });
  const texts = await getReviewedTextBatch(
    context,
    targets.map((t) => t.range),
  );
  await context.sync();
  return targets.map((t, i) => {
    const cell = /^table:/.test(t.key);
    assertTextTargetSafe(xml[i].value, { broad: broad && t.broad !== false, cell });
    const ids = identities[i]?.items.map((p) => p.uniqueLocalId);
    if (ids && (!ids.length || ids.some((id) => typeof id !== "string" || !id)))
      throw new Error(
        "TARGET_INSPECTION_UNAVAILABLE: container paragraph identities are unavailable.",
      );
    return {
      key: t.key,
      text: texts[i],
      ...(ids ? { paragraph_ids: ids } : {}),
      ooxml:
        cell || !broad
          ? stableCellOoxml(xml[i].value, { paragraphIdentity: cell ? !!ids?.length : true })
          : stableOoxml(xml[i].value, { fragment: true, paragraphIdentity: true }),
    };
  });
}

async function guardStoryFields(context, protectReview = false) {
  let result;
  try {
    result = context.document.body.getOoxml();
  } catch {
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: containing story OOXML is required.");
  }
  await context.sync();
  assertNoCrossParagraphFields(result.value, { protectReview });
}

// Insertions after a paragraph inside a multi-paragraph field (such as a TOC)
// would land in the field result and be discarded or garble it on update.
// Word may export one trailing empty paragraph; if the OOXML paragraphs still
// cannot be aligned with body.paragraphs, refuse whenever such a field exists.
async function assertInsertionOutsideFields(context, paragraphs, anchorIdx) {
  let result;
  try {
    result = context.document.body.getOoxml();
  } catch {
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: body OOXML is required to place an insertion.");
  }
  await context.sync();
  const openAfter = paragraphFieldOpenAfter(result.value);
  if (!openAfter)
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: body field structure could not be inspected.");
  const count = paragraphs.items.length;
  const aligned = openAfter.length === count || openAfter.length === count + 1;
  if (aligned ? openAfter[anchorIdx] : openAfter.some(Boolean))
    throw new Error(
      "PROTECTED_FIELD: the insertion point is inside a multi-paragraph field (such as a TOC); anchor after the field instead.",
    );
}

async function authorizeTargets(context, targets, control) {
  // Header/footer bodies are inspected independently below; all other broad
  // targets belong to the main body and may be inside a multi-paragraph field.
  if (targets.some((t) => !/^(header|footer):/.test(t.key))) await guardStoryFields(context, true);
  const state = await inspectTargets(context, targets);
  if (control?.expected && JSON.stringify(state) !== JSON.stringify(control.expected))
    throw new Error(
      "STALE_PROPOSAL: target identity, text, formatting or review structure changed.",
    );
  armWrite(context, control);
  return state;
}

async function verifyTexts(context, ranges, expected, { cell = false } = {}) {
  const actual = await getReviewedTextBatch(context, ranges);
  // Body ranges include the host's final paragraph mark; paragraph ranges do
  // not consistently include it. Normalize only a terminal paragraph mark.
  const normalize = (s) => (cell ? s.replace(/\t$/, "") : s.replace(/\r$/, ""));
  if (
    actual.length !== expected.length ||
    actual.some((s, i) => normalize(s) !== (cell ? expected[i] : normalize(expected[i])))
  )
    throw new Error(
      "RESULT_UNKNOWN: target Current-text readback differs; inspect Word before another edit.",
    );
}

async function broadTargets(context, name, args) {
  if (name === "office_set_table_cell") {
    for (const key of ["table_index", "row", "column"])
      if (!Number.isInteger(args[key]) || args[key] < 0)
        throw new Error(`INVALID_SELECTOR: ${key} must be nonnegative.`);
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();
    const table = tables.items[args.table_index];
    if (!table) throw new Error("Table not found.");
    return [
      {
        key: `table:${args.table_index}:${args.row}:${args.column}`,
        range: table.getCell(args.row, args.column).body,
      },
    ];
  }
  if (name === "office_header_footer") {
    if (!["header", "footer"].includes(args.which))
      throw new Error("which must be header or footer");
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();
    return sections.items.map((s, i) => ({
      key: `${args.which}:${i}:Primary`,
      range:
        args.which === "header"
          ? s.getHeader(Word.HeaderFooterType.primary)
          : s.getFooter(Word.HeaderFooterType.primary),
    }));
  }
  const { paragraphs, idMode } = await getParagraphsWithIds(context);
  if (idMode !== "uniqueLocalId")
    throw new Error("Stable paragraph IDs are required for broad writes.");
  const snapshot = await currentSnapshot(context, paragraphs, idMode);
  let indices;
  if (name === "office_replace_section") {
    const start = resolveHeading(snapshot, args.heading, args.heading_id);
    // The heading participates in stale checking although only its body is replaced.
    indices = Array.from({ length: findSectionEnd(snapshot, start) - start }, (_, i) => start + i);
  } else {
    const ids = args.ids || args.paragraph_ids;
    assertDistinctIds(ids);
    indices = ids.map((id) => {
      const i = findIndexById(paragraphs, id, idMode);
      if (i < 0) throw new Error(`PARAGRAPH_NOT_FOUND: ${id}`);
      return i;
    });
  }
  return indices.map((i, at) => ({
    key: snapshot[i].id,
    range: paragraphs.items[i].getRange(),
    broad: name !== "office_replace_section" || at !== 0,
  }));
}

export async function prepareSafeWordWrite(name, args = {}) {
  const tool = {
    office_replace_paragraphs: toolReplaceParagraphs,
    office_replace_section: toolReplaceSection,
    office_delete_paragraphs: toolDeleteParagraphs,
    office_set_table_cell: toolSetTableCell,
    office_header_footer: toolHeaderFooter,
  }[name];
  if (!tool) return null;
  const capturedArgs = structuredClone(args);
  const expected = await Word.run(async (c) => {
    if (name !== "office_header_footer") await guardStoryFields(c, true);
    return inspectTargets(c, await broadTargets(c, name, capturedArgs));
  });
  return {
    preview: {
      operation: name,
      args: capturedArgs,
      targets: expected.map(({ key, text }) => ({ key, current_text: text })),
      warning:
        "Exact target replacement. Rich runs, existing review markup, fields and embedded structures are refused. Native layout still requires review.",
    },
    apply: ({ guard, beforeWrite }) => tool(capturedArgs, { expected, guard, beforeWrite }),
  };
}

// Pin text-based headings to a live ID before presenting a generic operation's
// confirmation. The pane's existing body precondition then detects any change.
export async function resolveWordHeadingArgs(name, args = {}) {
  if (name === "office_insert_paragraphs" && args.after?.heading !== undefined) {
    assertOneSelector([args.after.id, args.after.heading]);
    return Word.run(async (c) => {
      const { paragraphs, idMode } = await getParagraphsWithIds(c);
      if (idMode !== "uniqueLocalId")
        throw new Error("Stable heading IDs are required for a controlled insertion.");
      const snapshot = await currentSnapshot(c, paragraphs, idMode);
      const i = resolveHeading(snapshot, args.after.heading);
      return { ...args, after: { id: snapshot[i].id } };
    });
  }
  if (
    ["office_clear_highlights", "office_clear_comments"].includes(name) &&
    (args.heading_section !== undefined || args.heading_id !== undefined)
  ) {
    assertOneSelector([
      args.paragraph_ids,
      args.heading_section,
      args.heading_id,
      args.all === true ? true : undefined,
    ]);
    return Word.run(async (c) => {
      const { paragraphs, idMode } = await getParagraphsWithIds(c);
      if (idMode !== "uniqueLocalId")
        throw new Error("Stable heading IDs are required for scoped cleanup.");
      const snapshot = await currentSnapshot(c, paragraphs, idMode);
      const i = resolveHeading(snapshot, args.heading_section, args.heading_id);
      return { ...args, heading_section: undefined, heading_id: snapshot[i].id };
    });
  }
  return args;
}

// Find the canonical body paragraph for the section that contains the
// insert position `fromIdx`.
//
//   1. Walk BACKWARD from fromIdx-1 to the nearest heading — the top of
//      the section the insert sits in.
//   2. From that heading, walk FORWARD for the first body paragraph in the
//      section, stopping at the next heading.
//   3. If there's no heading above the insert, or the section has no body
//      paragraph above the next heading, fall back to the first non-heading
//      paragraph anywhere in the doc.
//
// Returns `{ idx, inSection }`:
//   - `idx`       — index into `snapshot` of the reference paragraph (or -1).
//   - `inSection` — true when `idx` was found *inside* the insert's own
//     section (steps 1–2); false when it came from the doc-wide fallback
//     (step 3).
//
// Why `inSection` matters: the style NAME is a safe thing to copy even from
// the fallback (low blast radius — old behavior). But the reference
// paragraph's *direct* formatting (indent / alignment / spacing) is only
// representative if it's from the same section. The doc-wide fallback might
// be a centered title block or an indented address line, and cloning its
// direct formatting onto a body insert would be visibly wrong. Callers use
// `inSection` to gate format-cloning while still always taking the style
// name.
function findBodyReferenceIndex(snapshot, fromIdx) {
  // Step 1: find the section's heading by walking backward.
  let sectionStart = -1;
  for (let i = fromIdx - 1; i >= 0; i--) {
    if (snapshot[i] && snapshot[i].heading_level !== null) {
      sectionStart = i;
      break;
    }
  }

  // Step 2: from the heading, the first body paragraph in THIS section.
  // Stop at the next heading.
  if (sectionStart !== -1) {
    for (let i = sectionStart + 1; i < snapshot.length; i++) {
      if (!snapshot[i]) continue;
      if (snapshot[i].heading_level !== null) break;
      return { idx: i, inSection: true };
    }
  }

  // Step 3: doc-wide fallback — style name only, NOT a format-clone source.
  for (let i = 0; i < snapshot.length; i++) {
    if (snapshot[i] && snapshot[i].heading_level === null) {
      return { idx: i, inSection: false };
    }
  }
  return { idx: -1, inSection: false };
}

function inferBodyStyle(snapshot, fromIdx) {
  const { idx } = findBodyReferenceIndex(snapshot, fromIdx);
  return idx === -1 ? null : snapshot[idx].style || null;
}

// Paragraph-format properties cloned from a reference body paragraph onto
// inserted ones. Many templates set these directly (not via the named
// style), so a style-name copy alone leaves inserts left-aligned with no
// first-line indent and no inter-paragraph spacing.
const CLONED_PARA_FORMAT_PROPS = [
  "alignment",
  "firstLineIndent",
  "leftIndent",
  "rightIndent",
  "lineSpacing",
  "spaceBefore",
  "spaceAfter",
  "lineUnitBefore",
  "lineUnitAfter",
];

// Queue a load of the format props off a reference paragraph. Call before
// the next context.sync().
function loadParagraphFormat(refPara) {
  refPara.load(CLONED_PARA_FORMAT_PROPS.join(", "));
}

// Snapshot a synced reference paragraph's format values into a plain
// object. Decouples the captured formatting from the live proxy —
// necessary for office_replace_section, where the reference paragraph is
// deleted before the new paragraphs are inserted.
function snapshotParagraphFormat(refPara) {
  const out = {};
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    try {
      const v = refPara[prop];
      if (v !== undefined && v !== null) out[prop] = v;
    } catch {
      /* prop unsupported on this build */
    }
  }
  return out;
}

// Apply a snapshotted format object onto a target paragraph.
function applyParagraphFormat(targetPara, fmt) {
  if (!fmt) return;
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    const v = fmt[prop];
    if (v !== undefined && v !== null) {
      try {
        targetPara[prop] = v;
      } catch {
        /* prop unsupported on this build */
      }
    }
  }
}

// Find the end of a section starting at headingIdx (exclusive): next paragraph
// with a heading style of same-or-higher rank (i.e. lower or equal level).
function findSectionEnd(snapshot, headingIdx) {
  const startLevel = snapshot[headingIdx].heading_level;
  for (let i = headingIdx + 1; i < snapshot.length; i++) {
    const lvl = snapshot[i].heading_level;
    if (lvl !== null && lvl <= startLevel) return i;
  }
  return snapshot.length;
}

// ---------------------------------------------------------------------------
// Tool: office_get_selection
// ---------------------------------------------------------------------------
export async function toolGetSelection() {
  return await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    const selParas = sel.paragraphs;
    context.document.load("changeTrackingMode");

    // Try to load uniqueLocalId directly on the selected paragraphs — that
    // gives us stable IDs without any text-matching against the full doc
    // (which is unsafe whenever the doc has repeated/boilerplate paragraphs).
    // Pack ALL per-item properties (incl. listItemOrNullObject/listString
    // for rendered numbers) into one load before the first sync — a second
    // load/sync on the same paragraphs collection would silently un-load
    // the previously-loaded scalars and make every subsequent .text read
    // throw.
    let idMode = "uniqueLocalId";
    try {
      selParas.load(
        "items/text, items/style, items/uniqueLocalId, items/listItemOrNullObject/listString",
      );
      await context.sync();
      if (selParas.items.length > 0 && !selParas.items[0].uniqueLocalId) {
        throw new Error("uniqueLocalId not populated");
      }
    } catch {
      selParas.load("items/text, items/style");
      await context.sync();
      idMode = "index";
    }
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";

    // When track changes is on, .text concatenates tracked deletions with
    // their inserted replacements — the agent then sees "clean + stale"
    // and tries to "fix" the duplicate by re-replacing it, stacking more
    // tracked changes every turn. Use getReviewedText(current) on the
    // selection AND each selected paragraph so the agent sees what the
    // user sees with track changes accepted. This also runs when recording
    // is Off, because the document may still contain earlier revisions.
    let reviewedSelText = null;
    let reviewedPerPara = null;
    {
      const ranges = [sel, ...selParas.items.map((p) => p.getRange())];
      const reviewed = await getReviewedTextBatch(context, ranges);
      if (reviewed) {
        reviewedSelText = reviewed[0];
        reviewedPerPara = reviewed.slice(1);
      }
    }

    const selSnapshot = selParas.items.map((p, i) => {
      const text = reviewedPerPara ? reviewedPerPara[i] : p.text;
      const number = listNumberOf(p, p.text);
      return {
        // When uniqueLocalId is unavailable, return null rather than guessing
        // an index by text-matching — duplicate paragraphs in any doc make
        // text-match unreliable. The agent should fall back to text references.
        id: idMode === "uniqueLocalId" ? p.uniqueLocalId : null,
        style: p.style,
        text,
        ...(number ? { number } : {}),
      };
    });

    const finalText = reviewedSelText != null ? reviewedSelText : sel.text;
    return {
      text: finalText,
      coverage: { ...BODY_COVERAGE, scope: "selection" },
      is_empty: !finalText || finalText.length === 0,
      paragraphs: selSnapshot,
      addressing: idMode,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_read_paragraphs
// ---------------------------------------------------------------------------
export async function toolReadParagraphs({
  ids,
  heading_section,
  heading_id,
  range,
  preview,
} = {}) {
  assertOneSelector([ids, heading_section, heading_id, range], { optional: true });
  if (ids !== undefined) assertDistinctIds(ids);
  return await Word.run(async (context) => {
    const { paragraphs, idMode, trackChangesOn } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);
    if (heading_section !== undefined || heading_id !== undefined) {
      const current = await getReviewedTextBatch(
        context,
        paragraphs.items.map((p) => p.getRange()),
      );
      snapshot.forEach((p, i) => {
        p.text = current[i];
      });
    }

    let picked;
    // Truncation is on by default in the no-args ("dump the whole doc")
    // mode — that's an orientation pass, not a content read. Targeted forms
    // (`ids`, `heading_section`, `range`) always return full text. The agent
    // can also explicitly force `preview: false` in the no-args mode for a
    // full dump (heavy, but escape-hatch valid).
    let truncate = false;
    // Scope and preview are distinct from text version. All modes read Current.
    let isTargeted = false;
    if (ids && ids.length > 0) {
      const set = new Set(ids);
      picked = snapshot.filter((p) => set.has(p.id));
      const missing = ids.filter((id) => !picked.some((p) => p.id === id));
      if (missing.length)
        throw new Error(`PARAGRAPH_NOT_FOUND: ${missing.join(", ")}; no partial read returned.`);
      isTargeted = true;
    } else if (heading_section !== undefined || heading_id !== undefined) {
      const startIdx = resolveHeading(snapshot, heading_section, heading_id);
      const endIdx = findSectionEnd(snapshot, startIdx);
      picked = snapshot.slice(startIdx, endIdx);
      isTargeted = true;
    } else if (range !== undefined) {
      if (
        !Array.isArray(range) ||
        range.length !== 2 ||
        range.some((n) => !Number.isInteger(n) || n < 0) ||
        range[0] >= range[1] ||
        range[1] > snapshot.length
      )
        throw new Error(
          "INVALID_SELECTOR: range must be [start, end), with 0 <= start < end <= paragraph count.",
        );
      const [s, e] = range;
      picked = snapshot.slice(s, e);
      isTargeted = true;
    } else {
      // Default: outline-ish view. Word documents in the wild use varied
      // heading conventions (built-in Heading N, custom style names, all-caps
      // direct formatting), so we don't pre-filter — return every paragraph
      // with its style and a truncated text preview. The agent picks out the
      // headings.
      picked = snapshot;
      truncate = preview !== false;
    }

    // With track changes on, raw .text concatenates a tracked deletion's
    // original run with its inserted replacement — the agent reads back
    // "clean + stale" and tries to "fix" the duplicate, stacking more
    // tracked changes every turn. getReviewedText(current) returns what
    // the user sees with the deletions accepted, which is what the agent
    // should reason about. Preview truncates this Current text only afterwards.
    let reviewedByPicked = null;
    if (picked.length > 0) {
      const ranges = picked.map((p) => paragraphs.items[p.index].getRange());
      reviewedByPicked = await getReviewedTextBatch(context, ranges);
    }

    // Preview length: long enough to surface most in-paragraph references
    // (numbers, defined terms, citations) while still bounding the response
    // size on large docs. Was 117 originally — too aggressive; the agent
    // treated previews as full reads and missed mid-paragraph content.
    const PREVIEW_LEN = 500;

    return {
      paragraphs: picked.map((p, k) => {
        const full = reviewedByPicked && reviewedByPicked[k] != null ? reviewedByPicked[k] : p.text;
        const isTruncated = truncate && full.length > PREVIEW_LEN;
        return {
          id: p.id,
          style: p.style,
          ...(p.number ? { number: p.number } : {}),
          text: isTruncated ? full.slice(0, PREVIEW_LEN - 1) + "…" : full,
          // Structured truncation flag so the agent doesn't have to detect
          // "…" suffixes. When `truncated` is true, the paragraph has more
          // content; re-read it via `ids: [<id>]` (always full text) to get
          // it. `full_length` lets the agent decide if a re-read is worth
          // the round-trip.
          ...(isTruncated ? { truncated: true, full_length: full.length } : {}),
        };
      }),
      total_in_doc: snapshot.length,
      coverage: {
        ...BODY_COVERAGE,
        scope: isTargeted ? "selected_body_paragraphs" : "body_paragraphs",
        preview: truncate,
      },
      addressing: idMode,
      // When truncation is active across the response, surface it at the
      // top level too. Cheap signal for the agent's next-action decision.
      ...(truncate ? { preview_mode: true, preview_limit: PREVIEW_LEN } : {}),
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_insert_paragraphs
// ---------------------------------------------------------------------------
export async function toolInsertParagraphs(
  { after, content, track_changes, style_per_para, provenance_comment },
  control,
) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const snapshot = await currentSnapshot(context, paragraphs, idMode);

        let anchorIdx;
        assertOneSelector([after?.id, after?.heading]);
        if (after?.id) {
          anchorIdx = findIndexById(paragraphs, after.id, idMode);
          if (anchorIdx === -1) {
            throw new Error(`Anchor paragraph not found: ${after.id}`);
          }
        } else if (after?.heading) {
          anchorIdx = findHeadingIndex(snapshot, after.heading);
          if (anchorIdx === -1) throw new Error(`Heading not found: "${after.heading}"`);
        } else {
          throw new Error("after must specify either id or heading");
        }

        await assertInsertionOutsideFields(context, paragraphs, anchorIdx);
        const anchor = paragraphs.items[anchorIdx];
        // If no explicit style is given, infer the section's body style so we
        // don't inherit a heading's formatting. Start one past the anchor (the
        // first paragraph the new content will sit alongside).
        const useInferred = !(style_per_para && style_per_para.length === content.length);
        const ref = useInferred
          ? findBodyReferenceIndex(snapshot, anchorIdx + 1)
          : { idx: -1, inSection: false };
        const refIdx = ref.idx;
        const inferredStyle = refIdx === -1 ? null : snapshot[refIdx].style || null;

        // Clone the reference body paragraph's *direct* formatting (indent,
        // alignment, spacing) — templates apply these manually, so the style
        // name alone isn't enough. ONLY when the reference is from the
        // insert's own section: the doc-wide fallback could be a title block
        // or address line, and cloning its direct formatting onto a body
        // insert would be visibly wrong. The style name is still taken from
        // the fallback (low blast radius); only format-cloning is gated.
        let refFmt = null;
        if (refIdx !== -1 && ref.inSection) {
          const refPara = paragraphs.items[refIdx];
          loadParagraphFormat(refPara);
          await context.sync();
          refFmt = snapshotParagraphFormat(refPara);
        }

        let cursor = anchor;
        const inserted = [];
        armWrite(context, control);
        for (let i = 0; i < content.length; i++) {
          const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
          const explicit = style_per_para && style_per_para[i];
          if (explicit) p.style = explicit;
          else if (inferredStyle) p.style = inferredStyle;
          if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
          inserted.push(p);
          cursor = p;
        }

        if (provenance_comment && inserted.length > 0) {
          inserted[0].getRange().insertComment(provenance_comment);
        }

        await context.sync();

        let newIds;
        if (idMode === "uniqueLocalId") {
          for (const p of inserted) p.load("uniqueLocalId");
          await context.sync();
          newIds = inserted.map((p) => p.uniqueLocalId);
        } else {
          newIds = inserted.map((_, k) => fallbackId(anchorIdx + 1 + k));
        }

        return {
          inserted_count: inserted.length,
          new_para_ids: newIds,
          addressing: idMode,
        };
      },
      control,
    );
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_paragraphs
// ---------------------------------------------------------------------------
export async function toolReplaceParagraphs(
  { ids, content, track_changes, style_per_para, provenance_comment },
  control,
) {
  assertDistinctIds(ids);
  if (!ids || ids.length === 0) throw new Error("ids must be non-empty");
  if (!content || content.length === 0) throw new Error("content must be non-empty");
  if (ids.length !== content.length) {
    throw new Error(
      `ids.length (${ids.length}) must equal content.length (${content.length}). To grow or shrink a section, use office_insert_paragraphs or office_replace_section.`,
    );
  }

  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        if (idMode !== "uniqueLocalId")
          throw new Error("Stable paragraph IDs are required for broad writes.");

        const targets = ids.map((id) => {
          const idx = findIndexById(paragraphs, id, idMode);
          if (idx === -1) {
            throw new Error(`Paragraph not found: ${id}`);
          }
          return { id, idx, paragraph: paragraphs.items[idx] };
        });

        await authorizeTargets(
          context,
          targets.map((t) => ({ key: t.id, range: t.paragraph.getRange() })),
          control,
        );

        // `insertText(..., "Replace")` swaps the text content while keeping the
        // paragraph element (so style is preserved unless we override below).
        for (let i = 0; i < targets.length; i++) {
          const t = targets[i];
          t.paragraph.insertText(content[i], Word.InsertLocation.replace);
          if (style_per_para && style_per_para[i]) {
            t.paragraph.style = style_per_para[i];
          }
        }

        if (provenance_comment && targets.length > 0) {
          targets[0].paragraph.getRange().insertComment(provenance_comment);
        }

        await context.sync();
        await verifyTexts(
          context,
          targets.map((t) => t.paragraph.getRange()),
          content,
        );

        return {
          replaced_count: targets.length,
          paragraph_ids: targets.map((t) => t.id),
        };
      },
      control,
    );
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_text — surgical sub-paragraph search/replace
// ---------------------------------------------------------------------------
// The pane supplies a human-confirmed plan, never an approval bit from the model.
export async function prepareTextReplacement(args) {
  assertDistinctIds(args.paragraph_ids, "paragraph_ids");
  if (typeof args.find !== "string" || !args.find || typeof args.replace !== "string")
    throw new Error("A nonempty find string and replacement string are required.");
  assertLiteralSearchText(args.find);
  return Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const targets = args.paragraph_ids.map((id) => {
      const i = findIndexById(paragraphs, id, idMode);
      if (i < 0) throw new Error(`PARAGRAPH_NOT_FOUND: ${id}`);
      return { key: id, range: paragraphs.items[i].getRange(), paragraph: paragraphs.items[i] };
    });
    await guardStoryFields(context);
    const states = await inspectTargets(context, targets, false);
    const probes = targets.map((t) => {
      const results = t.paragraph.search(args.find, {
        matchCase: !!args.match_case,
        matchWholeWord: !!args.whole_word,
      });
      results.load("items/text");
      return results;
    });
    await context.sync();
    const rows = [];
    for (let i = 0; i < targets.length; i++) {
      if (probes[i].items.length !== 1)
        throw new Error(
          `AMBIGUOUS_RANGE: expected one Word search match in ${states[i].key}; found ${probes[i].items.length}.`,
        );
      const hit = probes[i].items[0];
      const [match, throughMatch] = await getReviewedTextBatch(context, [
        hit,
        targets[i].paragraph.getRange("Start").expandTo(hit),
      ]);
      if (match !== hit.text || !throughMatch.endsWith(match))
        throw new Error("AMBIGUOUS_RANGE: search hit includes non-Current text.");
      rows.push({
        id: states[i].key,
        text: states[i].text,
        search_matches: [{ text: match, index: throughMatch.length - match.length }],
      });
    }
    return planReplacement({ addressing: idMode, paragraphs: rows }, args).map((p, i) => ({
      ...p,
      target_ooxml: states[i].ooxml,
    }));
  });
}

export async function toolApplyTextReplacement(args, plan, { beforeWrite, guard }) {
  assertLiteralSearchText(args.find);
  return Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const targets = plan.map((p) => {
      const index = findIndexById(paragraphs, p.id, idMode);
      if (index < 0) throw new Error(`STALE_PROPOSAL: paragraph ${p.id} disappeared.`);
      return paragraphs.items[index];
    });
    const texts = await getReviewedTextBatch(
      context,
      targets.map((p) => p.getRange()),
    );
    assertPreconditions(
      { addressing: idMode, paragraphs: plan.map((p, i) => ({ id: p.id, text: texts[i] })) },
      plan,
    );
    await guardStoryFields(context);
    const state = await inspectTargets(
      context,
      targets.map((p, i) => ({ key: plan[i].id, range: p.getRange() })),
      false,
    );
    if (plan.some((p, i) => p.target_ooxml && p.target_ooxml !== state[i].ooxml))
      throw new Error("STALE_PROPOSAL: target formatting or review markup changed.");
    const probes = targets.map((p) => {
      const r = p.search(args.find, {
        matchCase: !!args.match_case,
        matchWholeWord: !!args.whole_word,
      });
      r.load("items/text");
      return r;
    });
    await context.sync();
    for (let i = 0; i < probes.length; i++) {
      if (probes[i].items.length !== 1 || probes[i].items[0].text !== plan[i].match)
        throw new Error("AMBIGUOUS_RANGE: Word search does not match the confirmed proposal.");
    }
    guard();
    const previousMode = context.document.changeTrackingMode;
    let attempted = false;
    let trackingAttempted = false;
    let result;
    let failure;
    try {
      // Recheck the live Current text after the last asynchronous gap.
      const latest = await getReviewedTextBatch(
        context,
        targets.map((p) => p.getRange()),
      );
      assertPreconditions(
        { addressing: idMode, paragraphs: plan.map((p, i) => ({ id: p.id, text: latest[i] })) },
        plan,
      );
      await guardStoryFields(context);
      const finalState = await inspectTargets(
        context,
        targets.map((p, i) => ({ key: plan[i].id, range: p.getRange() })),
        false,
      );
      if (finalState.some((s, i) => s.ooxml !== state[i].ooxml || s.text !== plan[i].before))
        throw new Error("STALE_PROPOSAL: target changed before mutation.");
      guard();
      beforeWrite();
      attempted = true;
      // A word or phrase replacement always leaves revision marks, including
      // when Setup is "Don't turn it on". Native Word serializes tracking mode
      // into settings.xml, so change it only after all preconditions, in the
      // same queued batch as the replacement; our own mode change must not
      // invalidate the approved XML snapshot. The previous mode is restored below.
      trackingAttempted = true;
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      probes.forEach((r) => r.items[0].insertText(args.replace, Word.InsertLocation.replace));
      await context.sync();
      const actual = await getReviewedTextBatch(
        context,
        targets.map((p) => p.getRange()),
      );
      if (actual.some((text, i) => text !== plan[i].after))
        throw new Error("Current-text readback did not match the approved result.");
      result = {
        total_replacements: plan.length,
        verification: "current_text_readback_matched",
        track_changes: true,
        paragraphs: plan.map((p, i) => ({ id: p.id, text: actual[i] })),
      };
    } catch (error) {
      failure = new Error((attempted ? "RESULT_UNKNOWN: " : "") + error.message);
    } finally {
      if (trackingAttempted) {
        try {
          context.document.changeTrackingMode = previousMode;
          await context.sync();
        } catch (error) {
          failure = new Error(
            `RESULT_UNKNOWN: tracking-mode restore failed; inspect Word. ${failure?.message || error.message}`,
          );
        }
      }
    }
    if (failure) throw failure;
    return result;
  });
}

// ---------------------------------------------------------------------------
// Tool: office_delete_paragraphs — remove paragraphs entirely (including
// their paragraph marks). office_replace_text with replace:"" only blanks
// text — the paragraph mark survives, so a stale/duplicated paragraph stays
// as an empty line. This is the only path that actually collapses the doc.
// ---------------------------------------------------------------------------
export async function toolDeleteParagraphs({ paragraph_ids, track_changes }, control) {
  assertDistinctIds(paragraph_ids);
  if (!Array.isArray(paragraph_ids) || paragraph_ids.length === 0) {
    throw new Error("paragraph_ids must be a non-empty array");
  }
  return await Word.run(async (context) =>
    withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        if (idMode !== "uniqueLocalId")
          throw new Error("Stable paragraph IDs are required for broad writes.");
        // Resolve every target to its paragraph proxy BEFORE deleting any.
        // Each delete() shifts subsequent indices, but the already-resolved
        // proxy objects stay valid — so collect first, delete second.
        const targets = [];
        for (const id of paragraph_ids) {
          const idx = findIndexById(paragraphs, id, idMode);
          if (idx === -1) throw new Error(`Paragraph not found: ${id}`);
          targets.push(paragraphs.items[idx]);
        }
        await authorizeTargets(
          context,
          targets.map((p, i) => ({ key: paragraph_ids[i], range: p.getRange() })),
          control,
        );
        for (const p of targets) p.delete();
        await context.sync();
        // Tracked deletions may retain paragraph objects; verify Current text,
        // not raw proxy existence or raw revision text.
        const latest = await getParagraphsWithIds(context);
        const remaining = paragraph_ids
          .map((id) => findIndexById(latest.paragraphs, id, latest.idMode))
          .filter((i) => i >= 0);
        if (remaining.length)
          await verifyTexts(
            context,
            remaining.map((i) => latest.paragraphs.items[i].getRange()),
            remaining.map(() => ""),
          );
        return { deleted: targets.length, verification: "target_current_text_readback" };
      },
      control,
    ),
  );
}

// ---------------------------------------------------------------------------
// Tool: office_replace_section
// ---------------------------------------------------------------------------
export async function toolReplaceSection(
  { heading, heading_id, content, track_changes, style_per_para, provenance_comment },
  control,
) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        let para1 = await getParagraphsWithIds(context);
        let paragraphs = para1.paragraphs;
        let idMode = para1.idMode;
        if (idMode !== "uniqueLocalId")
          throw new Error("Stable paragraph IDs are required for broad writes.");
        let snapshot = await currentSnapshot(context, paragraphs, idMode);

        const startIdx = resolveHeading(snapshot, heading, heading_id);
        const endIdx = findSectionEnd(snapshot, startIdx);

        // Capture the existing section's body style AND direct formatting
        // BEFORE deleting, so inserted paragraphs match. Without this, new
        // paragraphs inherit the heading's style (centered/bold/large) and
        // lose any manually-applied indent/justify/spacing.
        const ref =
          style_per_para && style_per_para.length === content.length
            ? { idx: -1, inSection: false }
            : findBodyReferenceIndex(snapshot, startIdx + 1);
        const refIdx = ref.idx;
        const bodyStyle = refIdx === -1 ? null : snapshot[refIdx].style || null;
        // Format-clone only when the reference is from this section (see
        // findBodyReferenceIndex). Style name is still taken from the
        // doc-wide fallback; direct formatting is not.
        let refFmt = null;
        if (refIdx !== -1 && ref.inSection) {
          const refPara = paragraphs.items[refIdx];
          loadParagraphFormat(refPara);
          await context.sync();
          refFmt = snapshotParagraphFormat(refPara);
        }

        const toDelete = paragraphs.items.slice(startIdx + 1, endIdx);
        const deletedCount = toDelete.length;
        const headingKey = snapshot[startIdx].id;
        await authorizeTargets(
          context,
          paragraphs.items.slice(startIdx, endIdx).map((p, i) => ({
            key: snapshot[startIdx + i].id,
            range: p.getRange(),
            broad: i !== 0,
          })),
          control,
        );
        for (const p of toDelete) p.delete();
        await context.sync();

        // Re-fetch. Heading index may have shifted by 0 (we only deleted *after* it)
        // but be safe and re-locate.
        const para2 = await getParagraphsWithIds(context);
        paragraphs = para2.paragraphs;
        idMode = para2.idMode;
        snapshot = await currentSnapshot(context, paragraphs, idMode);
        const retainedDeleted = snapshot.filter((p) =>
          toDelete.some((old) => old.uniqueLocalId === p.id),
        );
        if (retainedDeleted.some((p) => p.text.replace(/\r$/, "") !== ""))
          throw new Error("RESULT_UNKNOWN: old section text remains Current after deletion.");
        const newHeadingIdx = resolveHeading(snapshot, undefined, headingKey);
        if (newHeadingIdx === -1) throw new Error("Lost the heading after delete; aborting");

        const headingPara = paragraphs.items[newHeadingIdx];
        let cursor = headingPara;
        const inserted = [];
        // Deletion has already committed, but this new batch is still revocable.
        // A Stop/expiry/document switch during sync or re-reading must not insert
        // more content. Keep the attempted receipt and report a partial unknown.
        armWrite(context, control);
        for (let i = 0; i < content.length; i++) {
          const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
          const explicit = style_per_para && style_per_para[i];
          if (explicit) p.style = explicit;
          else if (bodyStyle) p.style = bodyStyle;
          if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
          inserted.push(p);
          cursor = p;
        }

        if (provenance_comment && inserted.length > 0) {
          inserted[0].getRange().insertComment(provenance_comment);
        }

        await context.sync();
        await verifyTexts(
          context,
          inserted.map((p) => p.getRange()),
          content,
        );

        let newIds;
        if (idMode === "uniqueLocalId") {
          for (const p of inserted) p.load("uniqueLocalId");
          await context.sync();
          newIds = inserted.map((p) => p.uniqueLocalId);
        } else {
          newIds = inserted.map((_, k) => fallbackId(newHeadingIdx + 1 + k));
        }

        return {
          deleted_paragraphs: deletedCount,
          inserted_count: inserted.length,
          new_para_ids: newIds,
          addressing: idMode,
        };
      },
      control,
    );
  });
}

// ---------------------------------------------------------------------------
// Tool: office_highlight
// ---------------------------------------------------------------------------
const SEVERITY_COLOR = {
  error: "Red",
  warning: "Yellow",
  info: "Turquoise",
  uncertain: "Pink",
};

export async function toolHighlight({ targets }, control) {
  if (!Array.isArray(targets) || targets.length === 0) {
    throw new Error("targets must be a non-empty array");
  }
  for (const t of targets) if (t?.query) assertLiteralSearchText(t.query, "query");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    // Phase 1: queue all the searches. Office.js batches operations until
    // sync(), so doing this in one pass is faster than one-sync-per-target.
    const perTarget = [];
    const pendingSearches = []; // { idx, target, searchResults }
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      const idx = findIndexById(paragraphs, t.paragraph_id, idMode);
      if (idx === -1) {
        perTarget[i] = { paragraph_id: t.paragraph_id, ok: false, error: "paragraph_not_found" };
        continue;
      }
      const paragraph = paragraphs.items[idx];
      if (t.query) {
        const searchResults = paragraph.search(t.query, { matchCase: false });
        searchResults.load("items");
        pendingSearches.push({ i, t, paragraph, searchResults });
      } else {
        pendingSearches.push({ i, t, paragraph, wholeParagraph: true });
      }
    }
    await context.sync();
    control?.guard();

    // Phase 2: apply highlights that still exist. A missing target must not
    // be reported as a fully completed edit, and a cancel during the search
    // sync must not reach this assignment.
    const writable = [];
    for (const p of pendingSearches) {
      if (!p.wholeParagraph && p.searchResults.items.length === 0) {
        perTarget[p.i] = {
          paragraph_id: p.t.paragraph_id,
          ok: false,
          matches: 0,
          severity: p.t.severity || "warning",
          error: "query_not_found",
          query: p.t.query,
        };
      } else writable.push(p);
    }
    let totalHighlights = 0;
    if (writable.length) {
      await checkWordFormattingPrecondition(context, "office_highlight", { targets }, control, {
        paragraphs,
        idMode,
      });
      control?.guard();
      control?.beforeWrite();
      for (const p of writable) {
        const color = SEVERITY_COLOR[p.t.severity || "warning"] || "Yellow";
        try {
          if (p.wholeParagraph) {
            p.paragraph.getRange().font.highlightColor = color;
            perTarget[p.i] = {
              paragraph_id: p.t.paragraph_id,
              ok: true,
              matches: 1,
              severity: p.t.severity || "warning",
            };
            totalHighlights += 1;
          } else {
            const items = p.searchResults.items;
            for (const r of items) r.font.highlightColor = color;
            perTarget[p.i] = {
              paragraph_id: p.t.paragraph_id,
              ok: true,
              matches: items.length,
              severity: p.t.severity || "warning",
            };
            totalHighlights += items.length;
          }
        } catch (e) {
          perTarget[p.i] = { paragraph_id: p.t.paragraph_id, ok: false, error: e.message };
        }
      }
      await context.sync();
    }
    const applied = perTarget.filter((t) => t?.ok);
    const notApplied = perTarget.filter((t) => t && t.ok === false);
    if (!applied.length) throw new Error("HIGHLIGHT_FAILED: none of the targets were highlighted.");
    return {
      status: notApplied.length ? "partial" : "completed",
      total_highlights: totalHighlights,
      per_target: perTarget,
      applied,
      not_applied: notApplied,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_highlights
// ---------------------------------------------------------------------------
export async function toolClearHighlights(args, control) {
  const { paragraph_ids, heading_section, heading_id, all } = args || {};
  assertOneSelector([paragraph_ids, heading_section, heading_id, all === true ? true : undefined]);
  if (paragraph_ids !== undefined) assertDistinctIds(paragraph_ids, "paragraph_ids");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = await currentSnapshot(context, paragraphs, idMode);

    let targets = [];
    if (all === true) {
      targets = paragraphs.items;
    } else if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx === -1) throw new Error(`PARAGRAPH_NOT_FOUND: ${id}`);
        targets.push(paragraphs.items[idx]);
      }
    } else if (heading_section !== undefined || heading_id !== undefined) {
      const startIdx = resolveHeading(snapshot, heading_section, heading_id);
      const endIdx = findSectionEnd(snapshot, startIdx);
      targets = paragraphs.items.slice(startIdx, endIdx);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    // Office.js batches all property assignments and applies them at
    // context.sync(). A single bad paragraph (empty paragraph, paragraph
    // in a content control, paragraph with only tracked deletions, etc.)
    // fails the whole batch with a generic InvalidArgument. So:
    //   Fast path — try the whole batch in one sync.
    //   Slow path — if the batch fails, sync one paragraph at a time and
    //               record which paragraphs rejected the change.
    // null is the Office.js named value for clearing a highlight.
    if (targets.length) {
      await checkWordFormattingPrecondition(context, "office_clear_highlights", args, control, {
        paragraphs,
        idMode,
      });
      armWrite(context, control);
    }
    try {
      for (const p of targets) {
        p.getRange().font.highlightColor = null;
      }
      await context.sync();
      return { status: "completed", cleared_paragraphs: targets.length, failed_paragraphs: 0 };
    } catch (batchErr) {
      let cleared = 0;
      const failed = [];
      for (let i = 0; i < targets.length; i++) {
        try {
          targets[i].getRange().font.highlightColor = null;
          await context.sync();
          cleared++;
        } catch (e) {
          failed.push({ index: i, error: e?.message ?? String(e) });
        }
      }
      if (!cleared) throw new Error("RESULT_UNKNOWN: highlight cleanup did not apply.");
      return {
        status: failed.length ? "partial" : "completed",
        cleared_paragraphs: cleared,
        failed_paragraphs: failed.length,
        failures: failed,
      };
    }
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_comments
// ---------------------------------------------------------------------------
export async function toolClearComments(args, control) {
  const { paragraph_ids, heading_section, heading_id, all } = args || {};
  assertOneSelector([paragraph_ids, heading_section, heading_id, all === true ? true : undefined]);
  if (paragraph_ids !== undefined) assertDistinctIds(paragraph_ids, "paragraph_ids");

  return await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    if (all === true) {
      const total = comments.items.length;
      armWrite(context, control);
      for (const c of comments.items) c.delete();
      await context.sync();
      return { cleared_comments: total };
    }

    // Scoped path: determine which paragraph each comment is anchored on.
    // Use the MAIN DOCUMENT anchor, not the comment's own contentRange.
    // Match session-local IDs; duplicate text is not an identity.
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);
    if (heading_section !== undefined || heading_id !== undefined) {
      const current = await getReviewedTextBatch(
        context,
        paragraphs.items.map((p) => p.getRange()),
      );
      snapshot.forEach((p, i) => {
        p.text = current[i];
      });
    }

    let allowedIdxs;
    if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      allowedIdxs = new Set();
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx === -1) throw new Error(`PARAGRAPH_NOT_FOUND: ${id}`);
        allowedIdxs.add(idx);
      }
    } else if (heading_section !== undefined || heading_id !== undefined) {
      const startIdx = resolveHeading(snapshot, heading_section, heading_id);
      const endIdx = findSectionEnd(snapshot, startIdx);
      allowedIdxs = new Set();
      for (let i = startIdx; i < endIdx; i++) allowedIdxs.add(i);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    const probes = comments.items.map((c) => {
      if (idMode !== "uniqueLocalId")
        throw new Error("Stable paragraph IDs are required for scoped comment cleanup.");
      const paras = c.getRange().paragraphs;
      paras.load("items/uniqueLocalId");
      return { comment: c, paras };
    });
    await context.sync();
    armWrite(context, control);

    let deletedCount = 0;
    for (const { comment, paras } of probes) {
      if (paras.items.length === 0) continue;
      // A cross-boundary comment must not be deleted from a narrower scope.
      const anchorIndices = paras.items.map((p) =>
        snapshot.findIndex((s) => s.id === p.uniqueLocalId),
      );
      if (anchorIndices.every((idx) => idx !== -1 && allowedIdxs.has(idx))) {
        comment.delete();
        deletedCount += 1;
      }
    }
    await context.sync();

    return { cleared_comments: deletedCount };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_add_comment
// ---------------------------------------------------------------------------
export async function toolAddComment({ paragraph_id, query, text }) {
  if (!text || !text.trim()) throw new Error("text must be non-empty");
  if (query) assertLiteralSearchText(query, "query");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const idx = findIndexById(paragraphs, paragraph_id, idMode);
    if (idx === -1) {
      throw new Error(`Paragraph not found: ${paragraph_id}`);
    }
    const paragraph = paragraphs.items[idx];

    let target;
    if (query) {
      const searchResults = paragraph.search(query, { matchCase: false });
      searchResults.load("items");
      await context.sync();
      if (searchResults.items.length !== 1) {
        throw new Error(
          `AMBIGUOUS_RANGE: query has ${searchResults.items.length} matches in paragraph ${paragraph_id}; choose a unique phrase.`,
        );
      }
      target = searchResults.items[0];
    } else {
      target = paragraph.getRange();
    }

    target.insertComment(text);
    await context.sync();

    return { paragraph_id, anchored_on: query || "whole_paragraph", comment_added: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 1 editing tools — formatting / structure / whole-doc read
// ---------------------------------------------------------------------------

// Resolve a list of paragraph IDs to live Paragraph objects (shared by the
// formatting tools below). Throws on the first unknown ID.
function resolveTargets(paragraphs, idMode, ids) {
  return ids.map((id) => {
    const idx = findIndexById(paragraphs, id, idMode);
    if (idx === -1) throw new Error(`Paragraph not found: ${id}`);
    return { id, paragraph: paragraphs.items[idx] };
  });
}

const FORMATTING_TOOLS = new Set([
  "office_set_font",
  "office_apply_style",
  "office_set_paragraph_formatting",
  "office_highlight",
  "office_clear_highlights",
  "office_set_list",
]);

async function formattingPrecondition(context, name, args, loadedParagraphs) {
  // Reloading this collection can invalidate previously captured native
  // paragraph references on Word for Mac (Document._GetObjectByReferenceId).
  // Reuse the tool's loaded collection while reading fresh target OOXML.
  const { paragraphs, idMode } = loadedParagraphs ?? (await getParagraphsWithIds(context));
  let ids;
  if (name === "office_highlight") {
    if (!Array.isArray(args.targets) || !args.targets.length)
      throw new Error("targets must be a non-empty array");
    ids = [...new Set(args.targets.map((target) => target.paragraph_id))];
  } else if (name === "office_clear_highlights") {
    assertOneSelector([
      args.paragraph_ids,
      args.heading_section,
      args.heading_id,
      args.all === true ? true : undefined,
    ]);
    if (args.all === true) ids = paragraphs.items.map((p, i) => getId(p, i, idMode));
    else if (args.paragraph_ids) ids = args.paragraph_ids;
    else {
      const snapshot = await currentSnapshot(context, paragraphs, idMode);
      const start = resolveHeading(snapshot, args.heading_section, args.heading_id);
      ids = snapshot.slice(start, findSectionEnd(snapshot, start)).map((p) => p.id);
    }
  } else {
    assertDistinctIds(args.ids);
    ids = args.ids;
  }
  const targets = ids.map((id) => {
    const index = findIndexById(paragraphs, id, idMode);
    if (index < 0) return { key: id, missing: true };
    const paragraph = paragraphs.items[index];
    try {
      return { key: id, style: paragraph.style, xml: paragraph.getRange().getOoxml() };
    } catch {
      throw new Error("TARGET_INSPECTION_UNAVAILABLE: target formatting OOXML is required.");
    }
  });
  await context.sync();
  return {
    addressing: idMode,
    targets: targets.map(({ xml, ...target }) =>
      xml
        ? {
            ...target,
            ooxml: stableCellOoxml(xml.value, { paragraphIdentity: idMode === "uniqueLocalId" }),
          }
        : target,
    ),
  };
}

export async function captureWordFormattingPrecondition(name, args = {}) {
  if (!FORMATTING_TOOLS.has(name)) return null;
  return Word.run((context) => formattingPrecondition(context, name, args));
}

async function checkWordFormattingPrecondition(context, name, args, control, loadedParagraphs) {
  if (!control?.formatExpected) return;
  const latest = await formattingPrecondition(context, name, args, loadedParagraphs);
  if (JSON.stringify(latest) !== JSON.stringify(control.formatExpected))
    throw new Error(
      "STALE_PROPOSAL: target style, formatting or review markup changed after approval.",
    );
}

// Tool: office_apply_style — restyle existing paragraphs in place (without
// rewriting their text, unlike insert/replace which take a style alongside
// new content).
export async function toolApplyStyle({ ids, style, track_changes }, control) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  if (typeof style !== "string" || !style.trim())
    throw new Error("style must be a non-empty string");
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const targets = resolveTargets(paragraphs, idMode, ids);
        await checkWordFormattingPrecondition(context, "office_apply_style", { ids }, control, {
          paragraphs,
          idMode,
        });
        armWrite(context, control);
        for (const t of targets) t.paragraph.style = style;
        await context.sync();
        return { styled_count: targets.length, paragraph_ids: targets.map((t) => t.id), style };
      },
      control,
    );
  });
}

// Tool: office_set_font — character formatting on whole paragraphs.
const UNDERLINE = (v) =>
  v === true ? "Single" : v === false ? "None" : typeof v === "string" ? v : undefined;
export async function toolSetFont(
  { ids, bold, italic, underline, size, color, name, track_changes },
  control,
) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const targets = resolveTargets(paragraphs, idMode, ids);
        await checkWordFormattingPrecondition(context, "office_set_font", { ids }, control, {
          paragraphs,
          idMode,
        });
        armWrite(context, control);
        for (const t of targets) {
          const f = t.paragraph.getRange().font;
          if (bold !== undefined) f.bold = bold;
          if (italic !== undefined) f.italic = italic;
          if (underline !== undefined) f.underline = UNDERLINE(underline);
          if (size !== undefined) f.size = size;
          if (color !== undefined) f.color = color;
          if (name !== undefined) f.name = name;
        }
        await context.sync();
        return { formatted_count: targets.length, paragraph_ids: targets.map((t) => t.id) };
      },
      control,
    );
  });
}

// Tool: office_set_paragraph_formatting — alignment / indent / spacing.
const ALIGNMENT = { left: "Left", center: "Centered", right: "Right", justify: "Justified" };
export async function toolSetParagraphFormatting(
  { ids, alignment, left_indent, space_before, space_after, line_spacing, track_changes },
  control,
) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  if (alignment !== undefined && !ALIGNMENT[alignment]) {
    throw new Error(`alignment must be one of: ${Object.keys(ALIGNMENT).join(", ")}`);
  }
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const targets = resolveTargets(paragraphs, idMode, ids);
        await checkWordFormattingPrecondition(
          context,
          "office_set_paragraph_formatting",
          { ids },
          control,
          { paragraphs, idMode },
        );
        armWrite(context, control);
        for (const t of targets) {
          const p = t.paragraph;
          if (alignment !== undefined) p.alignment = ALIGNMENT[alignment];
          if (left_indent !== undefined) p.leftIndent = left_indent;
          if (space_before !== undefined) p.spaceBefore = space_before;
          if (space_after !== undefined) p.spaceAfter = space_after;
          if (line_spacing !== undefined) p.lineSpacing = line_spacing;
        }
        await context.sync();
        return { formatted_count: targets.length, paragraph_ids: targets.map((t) => t.id) };
      },
      control,
    );
  });
}

// Tool: office_insert_table — insert a table, optionally after a given
// paragraph (default: appended at the end of the document body).
export async function toolInsertTable(
  { rows, after_paragraph_id, header, track_changes },
  control,
) {
  if (!Array.isArray(rows) || rows.length === 0 || !Array.isArray(rows[0])) {
    throw new Error("`rows` must be a non-empty 2D array of cell strings.");
  }
  const cols = rows[0].length;
  for (let i = 1; i < rows.length; i++) {
    if (!Array.isArray(rows[i]) || rows[i].length !== cols) {
      throw new Error(`Ragged rows: row ${i} has ${rows[i]?.length} cells but row 0 has ${cols}.`);
    }
  }
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        let table;
        if (after_paragraph_id) {
          const { paragraphs, idMode } = await getParagraphsWithIds(context);
          const idx = findIndexById(paragraphs, after_paragraph_id, idMode);
          if (idx === -1) throw new Error(`Paragraph not found: ${after_paragraph_id}`);
          await assertInsertionOutsideFields(context, paragraphs, idx);
          const anchor = paragraphs.items[idx].getRange(Word.RangeLocation.after);
          armWrite(context, control);
          table = anchor.insertTable(rows.length, cols, Word.InsertLocation.after, rows);
        } else {
          armWrite(context, control);
          table = context.document.body.insertTable(
            rows.length,
            cols,
            Word.InsertLocation.end,
            rows,
          );
        }
        if (header) table.headerRowCount = 1;
        await context.sync();
        return { rows: rows.length, columns: cols, header: !!header };
      },
      control,
    );
  });
}

// Tool: office_set_table_cell — overwrite one cell in an existing table.
// Tables are addressed by 0-based document order (body.tables).
export async function toolSetTableCell({ table_index, row, column, text, track_changes }, control) {
  if (!Number.isInteger(table_index) || table_index < 0)
    throw new Error("`table_index` must be a non-negative integer.");
  if (!Number.isInteger(row) || row < 0) throw new Error("`row` must be a non-negative integer.");
  if (!Number.isInteger(column) || column < 0)
    throw new Error("`column` must be a non-negative integer.");
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const tables = context.document.body.tables;
        tables.load("items");
        await context.sync();
        if (table_index >= tables.items.length) {
          throw new Error(`Table ${table_index} not found (document has ${tables.items.length}).`);
        }
        const cell = tables.items[table_index].getCell(row, column);
        await authorizeTargets(
          context,
          [{ key: `table:${table_index}:${row}:${column}`, range: cell.body }],
          control,
        );
        cell.body.clear();
        cell.body.insertText(text ?? "", Word.InsertLocation.start);
        await context.sync();
        await verifyTexts(context, [cell.body], [text ?? ""], { cell: true });
        return {
          table_index,
          row,
          column,
          set: true,
          verification: "target_current_text_readback",
        };
      },
      control,
    );
  });
}

// Tool: office_get_document_text — the full body text in one call (the
// paragraph-range read tools are for targeted reads; this is the whole doc).
// When track changes is on, use getReviewedText(current) so the agent sees
// the doc as the user sees it with deletions accepted — raw body.text
// concatenates tracked deletions with their replacements (see
// toolReadParagraphs for the full rationale).
export async function toolGetDocumentText() {
  return await Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    context.document.load("changeTrackingMode");
    await context.sync();
    const trackChangesOn = String(context.document.changeTrackingMode) !== "Off";

    let text = body.text || "";
    {
      const reviewed = await getReviewedTextBatch(context, [body]);
      if (reviewed && reviewed[0] != null) text = reviewed[0];
    }

    const words = text.trim() ? text.trim().split(/\s+/).length : 0;
    return {
      text,
      character_count: text.length,
      word_count: words,
      word_count_method: "whitespace_tokens_not_language_aware",
      coverage: BODY_COVERAGE,
    };
  });
}

// Tool: office_get_outline — the heading tree (id, level, text) for
// orientation without dumping the whole document.
export async function toolGetOutline() {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const headingIndices = paragraphs.items
      .map((p, i) => ((headingLevelOf(p) ?? 0) > 0 ? i : -1))
      .filter((i) => i >= 0);
    const current = await getReviewedTextBatch(
      context,
      headingIndices.map((i) => paragraphs.items[i].getRange()),
    );
    const outline = [];
    for (let n = 0; n < headingIndices.length; n++) {
      const i = headingIndices[n];
      const p = paragraphs.items[i];
      const level = headingLevelOf(p);
      if (current[n].trim()) {
        outline.push({ id: getId(p, i, idMode), level, text: current[n].trim() });
      }
    }
    return {
      heading_count: outline.length,
      outline,
      text_version: "Current",
      coverage: BODY_COVERAGE,
    };
  });
}

// ---------------------------------------------------------------------------
// Tier 2 editing tools — lists / media / links / find / comments / chrome
// ---------------------------------------------------------------------------

// Tool: office_set_list — turn existing paragraphs into a bulleted or
// numbered list (the first becomes a new list; the rest attach to it).
export async function toolSetList({ ids, ordered, track_changes }, control) {
  if (!Array.isArray(ids) || ids.length === 0) throw new Error("ids must be non-empty");
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const { paragraphs, idMode } = await getParagraphsWithIds(context);
        const targets = resolveTargets(paragraphs, idMode, ids);
        await checkWordFormattingPrecondition(context, "office_set_list", { ids }, control, {
          paragraphs,
          idMode,
        });
        armWrite(context, control);
        const list = targets[0].paragraph.startNewList();
        list.load("id");
        await context.sync();
        control?.guard();
        for (let i = 1; i < targets.length; i++) {
          targets[i].paragraph.attachToList(list.id, 0);
        }
        if (ordered) {
          list.setLevelNumbering(0, Word.ListNumbering.arabic, ["%1."]);
        } else {
          list.setLevelBullet(0, Word.ListBullet.solid);
        }
        await context.sync();
        return { list_count: targets.length, ordered: !!ordered, paragraph_ids: ids };
      },
      control,
    );
  });
}

// Tool: office_insert_image — insert an inline picture from base64 (the
// agent obtains it with e.g. `base64 -i path`). Appended at the end of the
// body, or after a given paragraph.
export async function toolInsertImage(
  { base64, after_paragraph_id, alt_text, track_changes },
  control,
) {
  if (typeof base64 !== "string" || !base64.trim()) {
    throw new Error("`base64` (a base64-encoded image) is required.");
  }
  const clean = base64.replace(/^data:image\/[a-z]+;base64,/i, "").trim();
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        let pic;
        if (after_paragraph_id) {
          const { paragraphs, idMode } = await getParagraphsWithIds(context);
          const idx = findIndexById(paragraphs, after_paragraph_id, idMode);
          if (idx === -1) throw new Error(`Paragraph not found: ${after_paragraph_id}`);
          await assertInsertionOutsideFields(context, paragraphs, idx);
          armWrite(context, control);
          pic = paragraphs.items[idx]
            .getRange(Word.RangeLocation.after)
            .insertInlinePictureFromBase64(clean, Word.InsertLocation.after);
        } else {
          armWrite(context, control);
          pic = context.document.body.insertInlinePictureFromBase64(clean, Word.InsertLocation.end);
        }
        if (alt_text) pic.altTextDescription = alt_text;
        await context.sync();
        return { inserted: true, anchored: after_paragraph_id || "end_of_document" };
      },
      control,
    );
  });
}

// Resolve a paragraph + optional sub-string query to a Range (shared by
// hyperlink/bookmark). Returns the whole paragraph range when no query.
async function rangeForParagraphQuery(context, paragraph_id, query) {
  const { paragraphs, idMode } = await getParagraphsWithIds(context);
  const idx = findIndexById(paragraphs, paragraph_id, idMode);
  if (idx === -1) throw new Error(`Paragraph not found: ${paragraph_id}`);
  const paragraph = paragraphs.items[idx];
  if (!query) return paragraph.getRange();
  assertLiteralSearchText(query, "query");
  const found = paragraph.search(query, { matchCase: false });
  found.load("items");
  await context.sync();
  if (found.items.length === 0) {
    throw new Error(`Query "${query}" not found in paragraph ${paragraph_id}`);
  }
  if (found.items.length > 1)
    throw new Error(
      `AMBIGUOUS_RANGE: query has ${found.items.length} matches in paragraph ${paragraph_id}; choose a unique phrase.`,
    );
  return found.items[0];
}

// External links are limited to web/mail schemes; "#name" targets a bookmark.
const SAFE_HYPERLINK = /^(https?:\/\/|mailto:|#)/i;

// Tool: office_insert_hyperlink — make text (a query within a paragraph, or
// the whole paragraph) a hyperlink.
export async function toolInsertHyperlink({ paragraph_id, query, url, track_changes }, control) {
  if (!paragraph_id) throw new Error("`paragraph_id` is required.");
  if (typeof url !== "string" || !url.trim()) throw new Error("`url` is required.");
  if (!SAFE_HYPERLINK.test(url.trim()))
    throw new Error("UNSUPPORTED_URL: use an http(s):, mailto: or #bookmark link.");
  return await Word.run(async (context) => {
    return await withProtectedTracking(
      context,
      track_changes,
      async () => {
        const range = await rangeForParagraphQuery(context, paragraph_id, query);
        armWrite(context, control);
        range.hyperlink = url;
        await context.sync();
        return { paragraph_id, anchored_on: query || "whole_paragraph", url };
      },
      control,
    );
  });
}

// Tool: office_insert_bookmark — drop a named bookmark on a paragraph or a
// text query within it.
export async function toolInsertBookmark({ paragraph_id, query, name }, control) {
  if (!paragraph_id) throw new Error("`paragraph_id` is required.");
  if (typeof name !== "string" || !name.trim()) throw new Error("`name` is required.");
  return await Word.run(async (context) => {
    const range = await rangeForParagraphQuery(context, paragraph_id, query);
    try {
      armWrite(context, control);
      range.insertBookmark(name);
      await context.sync();
    } catch (e) {
      throw new Error(`Bookmarks are not supported on this Word build (${e?.message ?? e}).`);
    }
    return { paragraph_id, anchored_on: query || "whole_paragraph", bookmark: name };
  });
}

// Tool: office_find — search the document; reports matches with the
// containing paragraph ID so the agent can follow up with an edit tool.
export async function toolFind({ query, match_case, whole_word, wildcards }) {
  if (typeof query !== "string" || !query.length) throw new Error("`query` is required.");
  return await Word.run(async (context) => {
    const results = context.document.body.search(query, {
      matchCase: !!match_case,
      matchWholeWord: !!whole_word,
      matchWildcards: !!wildcards,
    });
    results.load("items/text");
    await context.sync();
    const items = results.items.slice(0, 200);
    for (const r of items) r.paragraphs.load("items/uniqueLocalId");
    await context.sync();
    const current = await getReviewedTextBatch(context, items);
    const matches = items.map((r, i) => {
      const p = r.paragraphs.items[0];
      return {
        text: current[i],
        paragraph_id: p ? p.uniqueLocalId : null,
        text_version: "Current",
        raw_match_differs: current[i] !== r.text,
      };
    });
    return {
      query,
      match_count: results.items.length,
      matches,
      truncated: results.items.length > items.length,
      coverage:
        "Word search may match tracked/deleted text. Match text is Current; raw_match_differs requires revision inspection. At most 200 matches; narrow the search for further coverage.",
    };
  });
}

async function getCommentsLoaded(context, fields) {
  const comments = context.document.body.getComments();
  comments.load(fields);
  await context.sync();
  return comments;
}

// Tool: office_list_comments — every comment with id/author/text/resolved.
export async function toolListComments(args = {}) {
  return toolReadComments(args);
}

// Tool: office_reply_to_comment — add a reply to a comment thread by id.
export async function toolReplyToComment({ comment_id, text }) {
  if (!comment_id) throw new Error("`comment_id` is required.");
  if (typeof text !== "string" || !text.trim()) throw new Error("`text` is required.");
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id");
    const c = comments.items.find((x) => x.id === comment_id);
    if (!c) throw new Error(`Comment not found: ${comment_id}`);
    c.reply(text);
    await context.sync();
    return { comment_id, replied: true };
  });
}

// Tool: office_resolve_comment — mark a comment resolved (or reopen it).
export async function toolResolveComment({ comment_id, resolved = true }) {
  if (!comment_id) throw new Error("`comment_id` is required.");
  return await Word.run(async (context) => {
    const comments = await getCommentsLoaded(context, "items/id, items/resolved");
    const c = comments.items.find((x) => x.id === comment_id);
    if (!c) throw new Error(`Comment not found: ${comment_id}`);
    c.resolved = !!resolved;
    await context.sync();
    return { comment_id, resolved: !!resolved };
  });
}

// Tool: office_header_footer — set the primary header or footer text on
// every section (replaces existing content).
export async function toolHeaderFooter({ which, text }, control) {
  if (which !== "header" && which !== "footer") {
    throw new Error("`which` must be 'header' or 'footer'.");
  }
  return await Word.run(async (context) => {
    try {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      const targets = sections.items.map((section, i) => ({
        key: `${which}:${i}:Primary`,
        range:
          which === "header"
            ? section.getHeader(Word.HeaderFooterType.primary)
            : section.getFooter(Word.HeaderFooterType.primary),
      }));
      await authorizeTargets(context, targets, control);
      for (const { range } of targets) {
        range.clear();
        range.insertText(text ?? "", Word.InsertLocation.start);
      }
      await context.sync();
      await verifyTexts(
        context,
        targets.map((t) => t.range),
        targets.map(() => text ?? ""),
      );
      return {
        which,
        sections: sections.items.length,
        set: true,
        verification: "target_current_text_readback",
      };
    } catch (error) {
      if (context.__draftspectWriteAttempted && !String(error.message).includes("RESULT_UNKNOWN"))
        throw new Error("RESULT_UNKNOWN: " + error.message);
      throw error;
    }
  });
}
