/* global Word, Office */
import { stableOoxml, stableCellOoxml, verifyAcceptedInsertion } from "./word-ooxml-state.js";
import {
  assertTextTargetSafe,
  assertNoCrossParagraphFields,
  assertLiteralSearchText,
} from "./word-write-safety.js";
// Live Word review tools. Proposals retain plain data only, never Office proxies
// across Word.run contexts. OOXML is evidence/precondition data, not rewritten.
const FONT = [
  "bold",
  "italic",
  "underline",
  "size",
  "color",
  "name",
  "highlightColor",
  "strikeThrough",
  "subscript",
  "superscript",
];
const PARA = [
  "style",
  "alignment",
  "leftIndent",
  "rightIndent",
  "firstLineIndent",
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
];
const revisionSnapshots = new Map();
const formatSnapshots = new Map();
const own = (obj, keys) => Object.fromEntries(keys.map((k) => [k, obj[k] ?? null]));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const integer = (v, label) => {
  if (!Number.isInteger(v) || v < 0) throw new Error(`${label} must be a nonnegative integer`);
  return v;
};
function supported(version, name = "WordApi") {
  return !!globalThis.Office?.context?.requirements?.isSetSupported(name, version);
}
function requireApi(version, name = "WordApi") {
  if (!supported(version, name))
    throw new Error(
      `UNSUPPORTED_API: ${name} ${version} is required; this part was not inspected or changed.`,
    );
}
async function collection(context, value, fields = "items") {
  value.load(fields);
  await context.sync();
  return value.items;
}
async function reviewed(context, ranges, version = "Current") {
  requireApi("1.4");
  if (!ranges.length) return [];
  const saved = context.document.getSelection();
  const values = ranges.map((r) => r.getReviewedText(version));
  try {
    await context.sync();
  } finally {
    try {
      saved.select();
      await context.sync();
    } catch {
      /* best-effort host scroll workaround */
    }
  }
  if (values.some((r) => typeof r.value !== "string"))
    throw new Error("REVIEWED_TEXT_UNAVAILABLE: no raw text fallback");
  return values.map((r) => r.value);
}
async function xml(context, range) {
  const result = range.getOoxml();
  await context.sync();
  if (typeof result.value !== "string")
    throw new Error("OOXML_UNAVAILABLE: cannot verify this range");
  return result.value;
}
async function paragraphIds(context, range) {
  if (!supported("1.6")) return null;
  const ps = await collection(context, range.paragraphs, "items/uniqueLocalId");
  const ids = ps.map((p) => p.uniqueLocalId);
  if (ids.some((id) => typeof id !== "string" || !id))
    throw new Error("PARAGRAPH_ID_UNAVAILABLE: cannot verify native paragraph identity");
  return ids;
}
const xmlState = (raw, ids, fragment = false) => {
  const paragraphIdentity = Array.isArray(ids) && ids.length > 0;
  return fragment && paragraphIdentity
    ? stableCellOoxml(raw, { paragraphIdentity })
    : stableOoxml(raw, { fragment, paragraphIdentity });
};
const rangeState = (data) => ({
  ...data,
  target_ooxml: xmlState(data.target_ooxml, data.paragraph_ids, true),
  body_ooxml: xmlState(data.body_ooxml, data.body_paragraph_ids),
});
const commentRowsState = (rows) =>
  rows.map((row) => ({
    ...row,
    anchor: { ...row.anchor, ooxml: xmlState(row.anchor.ooxml, row.anchor.paragraph_ids, true) },
  }));
const revisionState = (data) => ({
  rows: data.rows.map((row) => ({ ...row, ooxml: xmlState(row.ooxml, row.paragraph_ids, true) })),
  story_xml: xmlState(data.story_xml, data.story_paragraph_ids, true),
  body_xml: xmlState(data.body_xml, data.body_paragraph_ids),
  story_paragraph_ids: data.story_paragraph_ids,
  body_paragraph_ids: data.body_paragraph_ids,
});

export async function resolveStory(context, target = {}) {
  const kind = target.story || "body";
  let body;
  if (kind === "body") body = context.document.body;
  else if (kind === "header" || kind === "footer") {
    const sections = await collection(context, context.document.sections);
    const section = sections[integer(target.section, "section")];
    if (!section) throw new Error("Section not found; reread the document map");
    const variant = target.variant || "Primary";
    if (!["Primary", "FirstPage", "EvenPages"].includes(variant))
      throw new Error("Unknown header/footer variant");
    body = kind === "header" ? section.getHeader(variant) : section.getFooter(variant);
  } else if (kind === "footnote" || kind === "endnote") {
    requireApi("1.5");
    const notes = await collection(context, context.document.body[kind + "s"]);
    body = notes[integer(target.note, "note")]?.body;
    if (!body) throw new Error("Note not found; reread the document map");
  } else throw new Error("Unknown story kind");
  if (target.shape_path?.length) {
    requireApi("1.2", "WordApiDesktop");
    let shapes = body.shapes;
    for (let i = 0; i < target.shape_path.length; i++) {
      const items = await collection(context, shapes, "items/type");
      const shape = items[integer(target.shape_path[i], "shape index")];
      if (!shape) throw new Error("Shape no longer exists");
      if (i < target.shape_path.length - 1) {
        if (shape.type !== "Group") throw new Error("Intermediate shape must be a group");
        shapes = shape.shapeGroup.shapes;
      } else {
        if (!["TextBox", "GeometricShape"].includes(shape.type))
          throw new Error("This shape has no supported text body");
        body = shape.body;
      }
    }
  }
  return body;
}

function storyOnly(target) {
  if (
    target.story === "selection" ||
    ["table", "row", "column", "paragraph_id", "query"].some((k) => target[k] !== undefined)
  )
    throw new Error(
      "This tool reads an entire story. Use story/section/variant/note/shape_path only; use office_read_formatting for a smaller range.",
    );
}

async function resolveRange(context, target = {}) {
  let range;
  if (target.story === "selection") range = context.document.getSelection();
  else {
    const body = await resolveStory(context, target);
    range = body.getRange();
    if (target.table !== undefined) {
      const tables = await collection(context, body.tables);
      const table = tables[integer(target.table, "table")];
      if (!table) throw new Error("Table not found");
      if (target.row !== undefined || target.column !== undefined) {
        range = table
          .getCell(integer(target.row, "row"), integer(target.column, "column"))
          .body.getRange();
      } else range = table.getRange();
    }
    if (target.paragraph_id) {
      requireApi("1.6");
      const paragraphs = await collection(context, range.paragraphs, "items/uniqueLocalId");
      const p = paragraphs.find((p) => p.uniqueLocalId === target.paragraph_id);
      if (!p)
        throw new Error(
          "Paragraph ID expired or is outside the target; reread it in this Word session",
        );
      range = p.getRange();
    }
  }
  if (target.query !== undefined) {
    if (typeof target.query !== "string" || !target.query)
      throw new Error("query must be nonempty");
    assertLiteralSearchText(target.query, "query");
    const hits = await collection(
      context,
      range.search(target.query, { matchCase: true }),
      "items/text",
    );
    if (hits.length !== 1 || hits[0].text !== target.query)
      throw new Error(
        `AMBIGUOUS_RANGE: expected one exact match, found ${hits.length}. Use a longer query or a paragraph ID.`,
      );
    range = hits[0];
    const [current] = await reviewed(context, [range]);
    if (current !== target.query)
      throw new Error("The match is not solely Current text; inspect its existing revisions first");
  }
  return range;
}

export async function toolDocumentMap() {
  const features = {
    reviewed_text_comments: supported("1.4"),
    notes: supported("1.5"),
    stable_session_paragraph_ids_and_revisions: supported("1.6"),
    shapes: supported("1.2", "WordApiDesktop"),
  };
  const targets = [{ story: "body" }],
    gaps = [],
    objects = [];
  // Optional failures stay isolated; an unsupported note must not hide the body.
  await Word.run(async (context) => {
    const sections = await collection(context, context.document.sections);
    sections.forEach((_, section) => {
      for (const story of ["header", "footer"])
        for (const variant of ["Primary", "FirstPage", "EvenPages"])
          targets.push({ story, section, variant });
    });
  });
  for (const story of ["footnote", "endnote"]) {
    try {
      requireApi("1.5");
      await Word.run(async (context) => {
        const notes = await collection(context, context.document.body[story + "s"]);
        notes.forEach((_, note) => targets.push({ story, note }));
      });
    } catch (error) {
      gaps.push({ story, error: error.message });
    }
  }
  if (features.shapes) {
    for (const base of [...targets]) {
      try {
        await Word.run(async (context) => {
          const walk = async (shapes, path = []) => {
            const items = await collection(context, shapes, "items/type,items/name");
            for (let i = 0; i < items.length; i++) {
              const s = items[i],
                shape_path = [...path, i];
              objects.push({ target: { ...base, shape_path }, type: s.type, name: s.name });
              if (["TextBox", "GeometricShape"].includes(s.type))
                targets.push({ ...base, shape_path });
              else if (s.type === "Group" && path.length < 8)
                await walk(s.shapeGroup.shapes, shape_path);
              else if (s.type !== "Picture")
                gaps.push({
                  target: { ...base, shape_path },
                  error: "Object content not read by the text tools; inspect OOXML/native Word.",
                });
            }
          };
          await walk((await resolveStory(context, base)).shapes);
        });
      } catch (error) {
        gaps.push({ target: base, component: "shapes", error: error.message });
      }
    }
  } else
    gaps.push({
      component: "shapes",
      error: "WordApiDesktop 1.2 unavailable; text boxes/floating shapes not enumerated.",
    });
  const stories = [];
  for (const target of targets) {
    try {
      stories.push(
        await Word.run(async (context) => {
          const body = await resolveStory(context, target);
          const ps = await collection(context, body.paragraphs);
          const ts = await collection(context, body.tables);
          const [text] = await reviewed(context, [body]);
          return {
            target,
            paragraph_count: ps.length,
            table_count: ts.length,
            current_characters: text.length,
            status: "inventoried_not_read",
            empty: !text.trim(),
          };
        }),
      );
    } catch (error) {
      gaps.push({ target, error: error.message });
    }
  }
  return {
    features,
    stories,
    objects,
    coverage: {
      complete_file_review: false,
      gaps,
      remaining: [
        "Read each nonempty story with office_read_story; follow next_offset.",
        "Read comments and revisions separately for each relevant story.",
        "Inspect targeted OOXML for runs, fields, numbering, tables and relationships.",
        "Picture/embedded-object meaning and visual layout require native review.",
      ],
    },
  };
}

export async function toolReadStory({
  target = {},
  offset = 0,
  limit = 100,
  version = "Current",
} = {}) {
  storyOnly(target);
  integer(offset, "offset");
  integer(limit, "limit");
  if (!limit || limit > 500) throw new Error("limit must be 1..500");
  if (!["Current", "Original"].includes(version)) throw new Error("Unknown text version");
  return Word.run(async (context) => {
    const body = await resolveStory(context, target);
    const stable = supported("1.6");
    const ps = await collection(
      context,
      body.paragraphs,
      "items/style" + (stable ? ",items/uniqueLocalId" : ""),
    );
    const selected = ps.slice(offset, offset + limit);
    const texts = await reviewed(
      context,
      selected.map((p) => p.getRange()),
      version,
    );
    return {
      target,
      version,
      total: ps.length,
      offset,
      next_offset: offset + selected.length < ps.length ? offset + selected.length : null,
      addressing: stable ? "session_local_id" : "read_only_index",
      paragraphs: selected.map((p, i) => ({
        index: offset + i,
        id: stable ? p.uniqueLocalId : null,
        style: p.style,
        text: texts[i],
      })),
      coverage:
        "Only the specified story and page; no text truncation. Tables/fields/runs require OOXML inspection. IDs must be reread after reopening Word.",
    };
  });
}

async function captureRange(context, target) {
  const range = await resolveRange(context, target);
  range.font.load(FONT.join(","));
  const ps = await collection(context, range.paragraphs, PARA.map((k) => "items/" + k).join(","));
  const [current] = await reviewed(context, [range]);
  const [original] = await reviewed(context, [range], "Original");
  const scope = target.query ? await resolveRange(context, { ...target, query: undefined }) : range;
  const [scope_current] = await reviewed(context, [scope]);
  return {
    range,
    data: {
      current,
      original,
      scope_current,
      font: own(range.font, FONT),
      paragraphs: ps.map((p) => own(p, PARA)),
      paragraph_ids: await paragraphIds(context, range),
      body_paragraph_ids: await paragraphIds(context, context.document.body),
      target_ooxml: await xml(context, range),
      body_ooxml: await xml(context, context.document.body),
    },
  };
}
export async function toolReadFormatting({
  target = {},
  ooxml_offset = 0,
  ooxml_limit = 30000,
  snapshot_id,
} = {}) {
  integer(ooxml_offset, "ooxml_offset");
  integer(ooxml_limit, "ooxml_limit");
  if (!ooxml_limit || ooxml_limit > 100000) throw new Error("ooxml_limit must be 1..100000");
  return Word.run(async (context) => {
    const { data } = await captureRange(context, target);
    let raw = data.target_ooxml;
    delete data.body_ooxml;
    delete data.body_paragraph_ids;
    delete data.target_ooxml;
    if (ooxml_offset > 0 && !snapshot_id)
      throw new Error("Pass the first chunk's snapshot_id when reading more OOXML");
    if (snapshot_id) {
      const saved = formatSnapshots.get(snapshot_id);
      if (!saved || Date.now() - saved.at > 300000)
        throw new Error("Formatting snapshot expired; restart at offset zero");
      assertUnchanged(saved.target, target);
      assertUnchanged(
        xmlState(saved.raw, saved.data.paragraph_ids, true),
        xmlState(raw, data.paragraph_ids, true),
      );
      assertUnchanged(saved.data, data);
      // Every page belongs to the first exact export, never a mixture of later
      // exports whose comparison-equivalent metadata has different byte lengths.
      raw = saved.raw;
    } else {
      snapshot_id = globalThis.crypto.randomUUID();
      formatSnapshots.set(snapshot_id, {
        at: Date.now(),
        target: structuredClone(target),
        raw,
        data,
      });
      while (formatSnapshots.size > 4) formatSnapshots.delete(formatSnapshots.keys().next().value);
    }
    return {
      target,
      snapshot_id,
      ...data,
      ooxml: raw.slice(ooxml_offset, ooxml_offset + ooxml_limit),
      ooxml_offset,
      ooxml_length: raw.length,
      next_ooxml_offset:
        ooxml_offset + ooxml_limit < raw.length ? ooxml_offset + ooxml_limit : null,
      coverage:
        "Font values may be mixed/null. OOXML preserves individual runs, styles/numbering references, fields, tables and revision markup; it is not Current-only text or proof of rendered layout. Pass snapshot_id and the same target for later chunks; changed state is refused.",
    };
  });
}

async function commentData(context, body) {
  requireApi("1.4");
  const comments = await collection(
    context,
    body.getComments(),
    "items/id,items/authorName,items/creationDate,items/content,items/resolved",
  );
  const probes = comments.map((c) => {
    const anchor = c.getRange();
    if (supported("1.6")) anchor.paragraphs.load("items/uniqueLocalId");
    c.replies.load("items/id,items/authorName,items/creationDate,items/content");
    return {
      c,
      anchor,
      current: anchor.getReviewedText("Current"),
      original: anchor.getReviewedText("Original"),
      ooxml: anchor.getOoxml(),
    };
  });
  await context.sync();
  const rows = probes.map(({ c, anchor, current, original, ooxml }) => {
    if ([current.value, original.value, ooxml.value].some((v) => typeof v !== "string"))
      throw new Error("Comment anchor was not populated");
    return {
      id: c.id,
      author: c.authorName,
      created: String(c.creationDate),
      text: c.content,
      resolved: c.resolved,
      anchor: {
        paragraph_ids: supported("1.6") ? anchor.paragraphs.items.map((p) => p.uniqueLocalId) : [],
        current: current.value,
        original: original.value,
        ooxml: ooxml.value,
      },
      replies: c.replies.items.map((r) => ({
        id: r.id,
        author: r.authorName,
        created: String(r.creationDate),
        text: r.content,
      })),
    };
  });
  return { comments, rows };
}
export async function toolReadComments({ target = {} } = {}) {
  storyOnly(target);
  return Word.run(async (context) => {
    const { rows } = await commentData(context, await resolveStory(context, target));
    return {
      target,
      comment_count: rows.length,
      comments: rows,
      coverage:
        "Threads and anchors in this story. Modern comments/native display still require Word verification.",
    };
  });
}

async function revisionData(context, target) {
  requireApi("1.6");
  const body = await resolveStory(context, target);
  const changes = await collection(
    context,
    body.getTrackedChanges(),
    "items/author,items/date,items/text,items/type",
  );
  const probes = changes.map((c) => {
    const range = c.getRange();
    range.paragraphs.load("items/uniqueLocalId");
    return {
      c,
      range,
      current: range.getReviewedText("Current"),
      original: range.getReviewedText("Original"),
      ooxml: range.getOoxml(),
    };
  });
  await context.sync();
  const rows = probes.map(({ c, range, current, original, ooxml }) => {
    if ([current.value, original.value, ooxml.value].some((v) => typeof v !== "string"))
      throw new Error("Revision content was not populated");
    return {
      author: c.author,
      date: String(c.date),
      type: c.type,
      text: c.text,
      current: current.value,
      original: original.value,
      paragraph_ids: range.paragraphs.items.map((p) => p.uniqueLocalId),
      ooxml: ooxml.value,
    };
  });
  return {
    changes,
    rows,
    story_xml: await xml(context, body),
    body_xml: await xml(context, context.document.body),
    story_paragraph_ids: await paragraphIds(context, body),
    body_paragraph_ids: await paragraphIds(context, context.document.body),
  };
}
export async function toolReadRevisions({ target = {} } = {}) {
  storyOnly(target);
  return Word.run(async (context) => {
    const data = await revisionData(context, target);
    const snapshot_id = globalThis.crypto.randomUUID();
    revisionSnapshots.set(snapshot_id, {
      target: structuredClone(target),
      data: { rows: data.rows, story_xml: data.story_xml, body_xml: data.body_xml },
      state: revisionState(data),
      at: Date.now(),
    });
    while (revisionSnapshots.size > 8)
      revisionSnapshots.delete(revisionSnapshots.keys().next().value);
    return {
      target,
      snapshot_id,
      revisions: data.rows.map((r, index) => ({ index, ...r })),
      enumeration_complete: false,
      coverage:
        "Native API revision rows in this story, not an exhaustive OOXML inventory. Word may collapse paired insertion/deletion replacements; accepting one can expose the existing counterpart. Empty Deleted.text is not proof of no deleted text: use Original and OOXML. Unsupported/move revisions may fail the host API; report that gap. Snapshot expires after five minutes or a document change.",
    };
  });
}

const SNAPSHOT_FIELDS = new Set([
  "current",
  "original",
  "scope_current",
  "font",
  "paragraphs",
  "paragraph_ids",
  "body_paragraph_ids",
  "story_paragraph_ids",
  "anchor_paragraph_ids",
  "target_ooxml",
  "body_ooxml",
  "rows",
  "body",
  "story",
  "anchor",
  "story_xml",
  "body_xml",
  "ooxml",
  "id",
  "author",
  "created",
  "text",
  "resolved",
  "replies",
  "date",
  "type",
  ...FONT,
  ...PARA,
]);
function changedSnapshotField(before, after, path = "snapshot") {
  if (same(before, after)) return null;
  if (Array.isArray(before) && Array.isArray(after)) {
    if (before.length !== after.length) return path + ".length";
    for (let i = 0; i < before.length; i++) {
      const field = changedSnapshotField(before[i], after[i], `${path}[${i}]`);
      if (field) return field;
    }
  } else if (before && after && typeof before === "object" && typeof after === "object") {
    for (const key of new Set([...Object.keys(before), ...Object.keys(after)])) {
      if (!same(before[key], after[key]))
        return SNAPSHOT_FIELDS.has(key)
          ? changedSnapshotField(before[key], after[key], `${path}.${key}`)
          : path;
    }
  }
  return path;
}
function assertUnchanged(before, after) {
  if (!same(before, after))
    throw new Error(
      `STALE_PROPOSAL: text, formatting, anchor or review state changed (${changedSnapshotField(before, after)}); read and propose again.`,
    );
}
function validatePatch(patch, keys) {
  if (!patch || typeof patch !== "object" || !Object.keys(patch).length)
    throw new Error("An explicit nonempty format patch is required");
  for (const k of Object.keys(patch)) {
    if (!keys.includes(k) || patch[k] === null || patch[k] === undefined)
      throw new Error("Unsupported format property: " + k);
    const expected = ["bold", "italic", "strikeThrough", "subscript", "superscript"].includes(k)
      ? "boolean"
      : [
            "size",
            "leftIndent",
            "rightIndent",
            "firstLineIndent",
            "spaceBefore",
            "spaceAfter",
            "lineSpacing",
          ].includes(k)
        ? "number"
        : "string";
    if (typeof patch[k] !== expected || (expected === "number" && !Number.isFinite(patch[k])))
      throw new Error("Invalid format value for " + k);
    if (["size", "lineSpacing"].includes(k) && patch[k] <= 0)
      throw new Error("Positive format value required for " + k);
  }
}

export async function prepareReviewWrite(name, args) {
  // Return a plain preview plus a closure that re-resolves the live proxies.
  if (name === "office_edit_range") {
    const target = structuredClone(args.target || {});
    if (!["replace_text", "format"].includes(args.action)) throw new Error("Unknown range action");
    if (
      args.action === "replace_text" &&
      (!target.query || typeof args.text !== "string" || /[\r\n]/.test(args.text))
    )
      throw new Error("A unique query and a single-paragraph replacement are required");
    if (args.action === "format") {
      if (args.font) validatePatch(args.font, FONT);
      if (args.paragraph) validatePatch(args.paragraph, PARA);
      if (!args.font && !args.paragraph) throw new Error("Choose font or paragraph properties");
    }
    const guardFields = async (c, range) => {
      if (args.action !== "replace_text") return;
      const story = target.story === "selection" ? c.document.body : await resolveStory(c, target);
      assertNoCrossParagraphFields(await xml(c, story));
      const ps = await collection(c, range.paragraphs);
      for (const p of ps) assertTextTargetSafe(await xml(c, p.getRange()));
    };
    let trackingOn = false;
    const before = await Word.run(async (c) => {
      c.document.load("changeTrackingMode");
      const captured = await captureRange(c, target);
      await guardFields(c, captured.range);
      await c.sync();
      trackingOn = String(c.document.changeTrackingMode) !== "Off";
      return captured.data;
    });
    return {
      preview: {
        action: args.action,
        target,
        before: { text: before.current, font: before.font, paragraphs: before.paragraphs },
        after:
          args.action === "replace_text"
            ? args.text
            : { font: args.font, paragraph: args.paragraph },
        // "Don't turn it on" leaves Word's current mode alone. Already-on
        // tracking still records the edit, so the card must not say otherwise.
        tracked: args.track_changes !== false || trackingOn,
        warning: args.paragraph
          ? "Paragraph properties apply to every paragraph touched by this range."
          : undefined,
      },
      apply: ({ guard, beforeWrite }) =>
        Word.run(async (context) => {
          let live = await captureRange(context, target);
          assertUnchanged(rangeState(before), rangeState(live.data));
          context.document.load("changeTrackingMode");
          await context.sync();
          const previous = context.document.changeTrackingMode;
          try {
            live = await captureRange(context, target);
            assertUnchanged(rangeState(before), rangeState(live.data));
            await guardFields(context, live.range);
            guard();
            beforeWrite();
            // false means "do not turn tracking on", not "turn it off".
            // Forcing Off here drops revisions for an edit the user already
            // had under Track Changes. TrackAll is only applied when the
            // caller asked to record the edit.
            if (args.track_changes !== false) context.document.changeTrackingMode = "TrackAll";
            if (args.action === "replace_text")
              live.range = live.range.insertText(args.text, "Replace");
            else {
              if (args.font) Object.assign(live.range.font, args.font);
              if (args.paragraph) {
                live.range.paragraphs.items.forEach((p) => Object.assign(p, args.paragraph));
              }
            }
            await context.sync();
            // Keep the exact proxy range for font readback; replacement text is
            // verified against the encompassing live story/paragraph below.
            if (args.action === "format") {
              live.range.font.load(FONT.join(","));
              const ps = await collection(
                context,
                live.range.paragraphs,
                PARA.map((p) => "items/" + p).join(","),
              );
              for (const [k, v] of Object.entries(args.font || {}))
                if (!same(live.range.font[k], v))
                  throw new Error("RESULT_UNKNOWN: font readback differs for " + k);
              for (const p of ps)
                for (const [k, v] of Object.entries(args.paragraph || {}))
                  if (!same(p[k], v))
                    throw new Error("RESULT_UNKNOWN: paragraph readback differs for " + k);
              const [current] = await reviewed(context, [live.range]);
              if (current !== before.current)
                throw new Error("RESULT_UNKNOWN: formatting unexpectedly changed text");
            } else {
              const [current] = await reviewed(context, [live.range]);
              if (current !== args.text)
                throw new Error("RESULT_UNKNOWN: replacement range readback differs; inspect Word");
              const scope = await resolveRange(context, { ...target, query: undefined });
              const [actualScope] = await reviewed(context, [scope]);
              if (
                before.scope_current.split(target.query).length !== 2 ||
                actualScope !== before.scope_current.replace(target.query, args.text)
              )
                throw new Error("RESULT_UNKNOWN: surrounding text readback differs; inspect Word");
            }
            return {
              verified: true,
              verification: "requested_properties_and_current_text",
              target,
              native_layout: "not_verified",
            };
          } finally {
            try {
              context.document.changeTrackingMode = previous;
              await context.sync();
            } catch {
              throw new Error("RESULT_UNKNOWN: could not restore tracking mode");
            }
          }
        }),
    };
  }
  if (name === "office_manage_comment") {
    const target = structuredClone(args.target || {});
    if (target.story === "selection")
      throw new Error("Use an explicit story and paragraph/query for a comment anchor");
    if (!["add", "edit", "reply", "resolve", "reopen", "delete"].includes(args.action))
      throw new Error("Unknown comment action");
    if (
      ["add", "edit", "reply"].includes(args.action) &&
      (typeof args.text !== "string" || !args.text.trim())
    )
      throw new Error("Nonempty comment text is required");
    const capture = async (context) => {
      const body = await resolveStory(context, target);
      const data = await commentData(context, body);
      const anchor = args.action === "add" ? await resolveRange(context, target) : null;
      const bodyIds = await paragraphIds(context, context.document.body);
      const storyIds = await paragraphIds(context, body);
      const anchorIds = anchor ? await paragraphIds(context, anchor) : null;
      return {
        ...data,
        anchor,
        state: {
          rows: commentRowsState(data.rows),
          body: xmlState(await xml(context, context.document.body), bodyIds),
          story: xmlState(await xml(context, body), storyIds, true),
          anchor: anchor ? xmlState(await xml(context, anchor), anchorIds, true) : null,
          body_paragraph_ids: bodyIds,
          story_paragraph_ids: storyIds,
          anchor_paragraph_ids: anchorIds,
        },
      };
    };
    const before = await Word.run(async (c) => {
      const data = await capture(c);
      return {
        rows: data.rows,
        state: data.state,
        anchor_text: data.anchor ? (await reviewed(c, [data.anchor]))[0] : null,
      };
    });
    const row = before.rows.find((c) => c.id === args.comment_id);
    if (args.action !== "add" && !row) throw new Error("Comment is missing from the chosen story");
    return {
      preview: {
        action: args.action,
        target,
        anchor_text: before.anchor_text,
        comment: row,
        text: args.text,
        warning:
          args.action === "delete"
            ? "Deletes the comment AND all replies; this is not tracked text."
            : "Comments are not tracked text revisions.",
      },
      apply: ({ guard, beforeWrite }) =>
        Word.run(async (context) => {
          const latest = await capture(context);
          assertUnchanged(before.state, latest.state);
          const c = latest.comments.find((c) => c.id === args.comment_id);
          guard();
          beforeWrite();
          let added;
          if (args.action === "add") {
            added = latest.anchor.insertComment(args.text);
            added.load("id");
          } else if (args.action === "edit") c.content = args.text;
          else if (args.action === "reply") c.reply(args.text);
          else if (args.action === "resolve" || args.action === "reopen")
            c.resolved = args.action === "resolve";
          else c.delete();
          await context.sync();
          const after = await commentData(context, await resolveStory(context, target));
          const actual = after.rows.find((c) => c.id === (added?.id || args.comment_id));
          const ok =
            args.action === "delete"
              ? !actual
              : args.action === "add" || args.action === "edit"
                ? actual?.text === args.text
                : args.action === "reply"
                  ? actual?.replies.length === row.replies.length + 1 &&
                    actual.replies.at(-1).text === args.text
                  : actual?.resolved === (args.action === "resolve");
          if (!ok) throw new Error("RESULT_UNKNOWN: comment readback differs");
          return { verified: true, action: args.action, comment: actual || null };
        }),
    };
  }
  if (name === "office_manage_revisions") {
    if (!["accept", "reject"].includes(args.action)) throw new Error("Choose accept or reject");
    const saved = revisionSnapshots.get(args.snapshot_id);
    if (!saved || Date.now() - saved.at > 300000)
      throw new Error("Revision snapshot expired; read revisions again");
    const indices = args.indices;
    if (!Array.isArray(indices) || !indices.length || new Set(indices).size !== indices.length)
      throw new Error("Explicit distinct revision indices are required");
    indices.forEach((i) => {
      integer(i, "revision index");
      if (!saved.data.rows[i]) throw new Error("Revision index not in snapshot");
    });
    return {
      preview: {
        action: args.action,
        target: saved.target,
        revisions: indices.map((i) => saved.data.rows[i]),
        warning: "Accept/reject removes existing revision marks. There is no automatic rollback.",
      },
      apply: ({ guard, beforeWrite }) =>
        Word.run(async (context) => {
          const latest = await revisionData(context, saved.target);
          assertUnchanged(saved.state, revisionState(latest));
          const invariantVersion = args.action === "accept" ? "Current" : "Original";
          const [beforeText] = await reviewed(
            context,
            [await resolveStory(context, saved.target)],
            invariantVersion,
          );
          guard();
          beforeWrite();
          [...indices].sort((a, b) => b - a).forEach((i) => latest.changes[i][args.action]());
          await context.sync();
          revisionSnapshots.delete(args.snapshot_id);
          const after = await revisionData(context, saved.target);
          const expected = saved.data.rows.filter((_, i) => !indices.includes(i));
          const semantic = (rows) =>
            rows.map(({ ooxml, paragraph_ids, current, original, ...r }) => r);
          const sameApiRows = same(semantic(expected), semantic(after.rows));
          const insertionProof =
            !sameApiRows &&
            args.action === "accept" &&
            (saved.target.story || "body") === "body" &&
            same(latest.body_paragraph_ids, after.body_paragraph_ids) &&
            same(latest.story_paragraph_ids, after.story_paragraph_ids) &&
            verifyAcceptedInsertion(
              latest.body_xml,
              after.body_xml,
              indices.map((i) => latest.rows[i]),
            );
          if (!sameApiRows && !insertionProof)
            throw new Error(
              `RESULT_UNKNOWN: remaining revision list differs (${changedSnapshotField(semantic(expected), semantic(after.rows))}; expected_count=${expected.length}; actual_count=${after.rows.length}); inspect Word before continuing`,
            );
          const [afterText] = await reviewed(
            context,
            [await resolveStory(context, saved.target)],
            invariantVersion,
          );
          if (beforeText !== afterText)
            throw new Error(
              "RESULT_UNKNOWN: accept/reject changed the invariant " +
                invariantVersion +
                " story text",
            );
          return {
            verified: true,
            action: args.action,
            affected: indices.length,
            remaining: after.rows.length,
            verification: insertionProof
              ? "selected_insertion_wrapper_removed_full_ooxml_and_Current_story_text"
              : "remaining_revision_metadata_and_" + invariantVersion + "_story_text",
            api_revision_list_changed: !sameApiRows,
            current_text: (await reviewed(context, [await resolveStory(context, saved.target)]))[0],
            native_layout: "not_verified",
          };
        }),
    };
  }
  throw new Error("Unknown guarded review operation");
}
