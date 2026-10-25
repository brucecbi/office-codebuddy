// Pure redline-audit assembler for one Word story. No Word.run, no writes,
// no model calls, and no Node-only imports: the taskpane loads this module.
// Live collection stays in word-review.js and reuses its existing readers.

export const REDLINE_AUDIT_SCHEMA_VERSION = 1;
export const REDLINE_AUDIT_ADAPTER = "office-codebuddy-word-review-p1a";
export const REDLINE_AUDIT_TTL_MS = 300_000;
export const MODEL_OOXML_ACCESS =
  "Default model result omits raw OOXML only. For exact XML use office_list_comments, office_read_revisions, or office_audit_redline with include_ooxml:true, or office_read_formatting with include_ooxml:true and small paged ooxml_limit (e.g. 8000). Native verification still used the complete snapshot.";
export const FULL_SNAPSHOT_SHA256_COVERS =
  "Full taskpane audit result, including raw ooxml strings, excluding full_snapshot_sha256 and evidence_sha256.";
export const EVIDENCE_SHA256_COVERS =
  "Default MCP projection: raw ooxml strings removed, ooxml_omitted and ooxml_chunk_length recorded, ooxml_access included when any ooxml was omitted, excluding full_snapshot_sha256 and evidence_sha256. This is not a file hash.";

const KINDS = new Set(["live_word", "synthetic_fixture", "range_ooxml", "disk_docx"]);
const PRIOR_STATUS = new Set(["expired", "not_current", "not_verified", "matches_probe"]);
const K = new Uint32Array([
  0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
  0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
  0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
  0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
  0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
  0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
  0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
  0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
]);

function rotr(x, n) {
  return (x >>> n) | (x << (32 - n));
}

// Sync SHA-256 so evidence hashes work in the Office webview without depending
// on crypto.subtle, which is unavailable outside a secure context.
export function sha256Hex(text) {
  const data = new TextEncoder().encode(String(text));
  const bitLen = data.length * 8;
  const withPad = new Uint8Array(((data.length + 9 + 63) >> 6) << 6);
  withPad.set(data);
  withPad[data.length] = 0x80;
  const view = new DataView(withPad.buffer);
  view.setUint32(withPad.length - 8, Math.floor(bitLen / 2 ** 32));
  view.setUint32(withPad.length - 4, bitLen >>> 0);
  let h0 = 0x6a09e667,
    h1 = 0xbb67ae85,
    h2 = 0x3c6ef372,
    h3 = 0xa54ff53a,
    h4 = 0x510e527f,
    h5 = 0x9b05688c,
    h6 = 0x1f83d9ab,
    h7 = 0x5be0cd19;
  const w = new Uint32Array(64);
  for (let i = 0; i < withPad.length; i += 64) {
    for (let t = 0; t < 16; t++) w[t] = view.getUint32(i + t * 4);
    for (let t = 16; t < 64; t++) {
      const s0 = rotr(w[t - 15], 7) ^ rotr(w[t - 15], 18) ^ (w[t - 15] >>> 3);
      const s1 = rotr(w[t - 2], 17) ^ rotr(w[t - 2], 19) ^ (w[t - 2] >>> 10);
      w[t] = (w[t - 16] + s0 + w[t - 7] + s1) >>> 0;
    }
    let a = h0,
      b = h1,
      c = h2,
      d = h3,
      e = h4,
      f = h5,
      g = h6,
      h = h7;
    for (let t = 0; t < 64; t++) {
      const s1 = rotr(e, 6) ^ rotr(e, 11) ^ rotr(e, 25);
      const ch = (e & f) ^ (~e & g);
      const temp1 = (h + s1 + ch + K[t] + w[t]) >>> 0;
      const s0 = rotr(a, 2) ^ rotr(a, 13) ^ rotr(a, 22);
      const maj = (a & b) ^ (a & c) ^ (b & c);
      const temp2 = (s0 + maj) >>> 0;
      h = g;
      g = f;
      f = e;
      e = (d + temp1) >>> 0;
      d = c;
      c = b;
      b = a;
      a = (temp1 + temp2) >>> 0;
    }
    h0 = (h0 + a) >>> 0;
    h1 = (h1 + b) >>> 0;
    h2 = (h2 + c) >>> 0;
    h3 = (h3 + d) >>> 0;
    h4 = (h4 + e) >>> 0;
    h5 = (h5 + f) >>> 0;
    h6 = (h6 + g) >>> 0;
    h7 = (h7 + h) >>> 0;
  }
  return [h0, h1, h2, h3, h4, h5, h6, h7].map((x) => x.toString(16).padStart(8, "0")).join("");
}

// Same omission the Office MCP bridge applies to compact review results.
// The default model hash is computed from this projection, not from the raw snapshot.
export function projectModelOoxml(result) {
  let omitted = 0;
  function project(value) {
    if (Array.isArray(value)) return value.map(project);
    if (!value || typeof value !== "object") return value;
    const entries = Object.entries(value)
      .filter(([key, item]) => {
        if (key !== "ooxml" || typeof item !== "string") return true;
        omitted++;
        return false;
      })
      .map(([key, item]) => [key, project(item)]);
    if (typeof value.ooxml === "string") {
      entries.push(["ooxml_omitted", true], ["ooxml_chunk_length", value.ooxml.length]);
    }
    return Object.fromEntries(entries);
  }
  const projected = project(result);
  return omitted ? { ...projected, ooxml_access: MODEL_OOXML_ACCESS } : projected;
}

export function canonicalJson(value) {
  if (typeof value === "string" || typeof value === "boolean" || value === null)
    return JSON.stringify(value);
  if (typeof value === "number") {
    if (!Number.isFinite(value))
      throw new Error("Evidence canonicalization cannot hash non-finite numbers");
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) return `[${value.map((item) => canonicalJson(item)).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value)
      .filter((key) => value[key] !== undefined)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`)
      .join(",")}}`;
  }
  throw new Error("Evidence canonicalization cannot hash this value");
}

export function storyScope(target = {}) {
  const source = target && typeof target === "object" ? target : {};
  const story = source.story || "body";
  const normalized = {
    story,
    section: source.section ?? null,
    variant: source.variant ?? null,
    note: source.note ?? null,
    shape_path: source.shape_path ?? null,
  };
  const label = [story];
  if (source.section !== undefined) label.push(`section ${source.section}`);
  if (source.variant) label.push(source.variant);
  if (source.note !== undefined) label.push(`note ${source.note}`);
  if (source.shape_path?.length) label.push(`shape ${source.shape_path.join(".")}`);
  return {
    key: canonicalJson(normalized),
    label: label.join(" / "),
    target: {
      story,
      ...(source.section !== undefined ? { section: source.section } : {}),
      ...(source.variant ? { variant: source.variant } : {}),
      ...(source.note !== undefined ? { note: source.note } : {}),
      ...(source.shape_path?.length ? { shape_path: [...source.shape_path] } : {}),
    },
  };
}

export function countRevisionMarkup(ooxml) {
  if (typeof ooxml !== "string" || !ooxml) return null;
  const ins = ooxml.match(/<(?:[\w.]+:)?ins(?=[\s/>])/g);
  const del = ooxml.match(/<(?:[\w.]+:)?del(?=[\s/>])/g);
  return { ins: ins ? ins.length : 0, del: del ? del.length : 0 };
}

function addMarkup(total, part) {
  if (!part) return total;
  if (!total) return { ins: part.ins, del: part.del };
  return { ins: total.ins + part.ins, del: total.del + part.del };
}

export function semanticRevision(row) {
  return {
    author: row?.author ?? null,
    date: row?.date == null ? null : String(row.date),
    type: row?.type ?? null,
    text: row?.text ?? null,
    current: row?.current ?? null,
    original: row?.original ?? null,
    paragraph_ids: Array.isArray(row?.paragraph_ids) ? [...row.paragraph_ids] : [],
  };
}

export function semanticComment(row) {
  return {
    id: row?.id ?? null,
    author: row?.author ?? null,
    resolved: row?.resolved ?? null,
    text: row?.text ?? null,
    created: row?.created == null ? null : String(row.created),
    replies: Array.isArray(row?.replies)
      ? row.replies.map((reply) => ({
          id: reply?.id ?? null,
          author: reply?.author ?? null,
          created: reply?.created == null ? null : String(reply.created),
          text: reply?.text ?? null,
        }))
      : [],
    anchor_current: row?.anchor?.current ?? null,
    anchor_original: row?.anchor?.original ?? null,
    paragraph_ids: Array.isArray(row?.anchor?.paragraph_ids) ? [...row.anchor.paragraph_ids] : [],
  };
}

function failureStatus(error, unsupported) {
  const reason = typeof error === "string" && error ? error : "Read failed";
  const status =
    unsupported || /UNSUPPORTED_API|Unknown story kind|Unknown header\/footer/.test(reason)
      ? "unsupported"
      : "read_failed";
  return { status, reason };
}

function classifyGap(gap) {
  const reason = gap?.error || "Document map reported a gap";
  const unsupported =
    /UNSUPPORTED_API|unavailable|not enumerated|not read by the text tools|Unknown story/i.test(
      reason,
    );
  return { status: unsupported ? "unsupported" : "read_failed", reason };
}

function gapLabel(gap) {
  if (gap?.component) return `map:${gap.component}`;
  if (gap?.story && !gap.target) return `map:${gap.story}`;
  if (gap?.target) return storyScope(gap.target).label;
  if (gap?.story) return `map:${gap.story}`;
  return "map";
}

function locateExact(paragraphs, query) {
  const hits = [];
  for (const paragraph of paragraphs) {
    const text = typeof paragraph.text === "string" ? paragraph.text : "";
    let from = 0;
    while (query && from <= text.length) {
      const at = text.indexOf(query, from);
      if (at < 0) break;
      hits.push({
        paragraph_id: paragraph.id ?? null,
        paragraph_index: paragraph.index ?? null,
        start: at,
      });
      from = at + query.length;
    }
  }
  return hits;
}

function combinePages(pages) {
  if (!Array.isArray(pages) || !pages.length) {
    return {
      status: "not_read",
      reason: "No page was read",
      paragraphs: [],
      paging_complete: false,
      next_cursor: null,
      total: null,
      addressing: null,
    };
  }
  let expected = pages[0].offset ?? 0;
  const paragraphs = [];
  for (const page of pages) {
    if (page?.error) {
      const failed = failureStatus(page.error, page.unsupported);
      return {
        ...failed,
        paragraphs,
        paging_complete: false,
        next_cursor: page.next_offset ?? null,
        total: page.total ?? null,
        addressing: page.addressing ?? null,
      };
    }
    const offset = page.offset ?? 0;
    if (offset !== expected) {
      return {
        status: "read_failed",
        reason: `Story page offset ${offset} did not continue from ${expected}`,
        paragraphs,
        paging_complete: false,
        next_cursor: offset,
        total: page.total ?? null,
        addressing: page.addressing ?? null,
      };
    }
    for (const paragraph of page.paragraphs || []) paragraphs.push(paragraph);
    expected = page.next_offset ?? offset + (page.paragraphs?.length || 0);
  }
  const last = pages.at(-1);
  const paging_complete = last.next_offset == null;
  const started_at_zero = (pages[0].offset ?? 0) === 0;
  return {
    status: "read",
    reason: paging_complete
      ? started_at_zero
        ? "All returned pages were read"
        : "The read did not start at offset 0, so earlier paragraphs were not covered"
      : "A later page remains. next_offset is still set, so this is not full coverage",
    paragraphs,
    paging_complete: paging_complete && started_at_zero,
    next_cursor: last.next_offset ?? null,
    total: last.total ?? null,
    addressing: last.addressing ?? null,
  };
}

function probeComparable(probe) {
  return !!(
    probe &&
    !probe.revision_error &&
    !probe.comment_error &&
    !probe.paragraph_error &&
    Array.isArray(probe.revisions) &&
    Array.isArray(probe.comments) &&
    Array.isArray(probe.paragraphs)
  );
}

function observationsMatch(after, input) {
  const revisionReads = input.revision_reads || [];
  const commentReads = input.comment_reads || [];
  const storyReads = input.story_reads || [];
  if (revisionReads.some((read) => read?.error) || commentReads.some((read) => read?.error))
    return false;
  const revisions = revisionReads.flatMap((read) => (read.revisions || []).map(semanticRevision));
  if (canonicalJson(revisions) !== canonicalJson(after.revisions)) return false;
  const comments = commentReads.flatMap((read) => (read.comments || []).map(semanticComment));
  if (canonicalJson(comments) !== canonicalJson(after.comments)) return false;
  for (const read of revisionReads) {
    if (
      typeof read.story_stable_sha256 === "string" &&
      typeof after.story_stable_sha256 === "string" &&
      read.story_stable_sha256 !== after.story_stable_sha256
    )
      return false;
  }
  for (const read of storyReads) {
    if (read?.error || read?.pages?.some((page) => page?.error)) return false;
    if (read?.version !== "Current") continue;
    const pages = combinePages(read.pages || []);
    if (pages.status !== "read") return false;
    const collected = pages.paragraphs.map((paragraph) => ({
      id: paragraph.id ?? null,
      text: paragraph.text ?? null,
    }));
    if (pages.paging_complete) {
      // A finished read is the whole Current list: count, order, and text.
      if (canonicalJson(collected) !== canonicalJson(after.paragraphs)) return false;
      continue;
    }
    // An unfinished page still has to match the slice already read. The
    // coverage gap for the unread remainder stays on the coverage report.
    const start = Number.isInteger(pages.paragraphs[0]?.index)
      ? pages.paragraphs[0].index
      : (read.pages?.[0]?.offset ?? 0);
    const probed = after.paragraphs.slice(start, start + collected.length);
    if (canonicalJson(collected) !== canonicalJson(probed)) return false;
  }
  return true;
}

function consistencyOf(input) {
  const supplied = input.consistency || {};
  const before = supplied.before ?? null;
  const after = supplied.after ?? null;
  let state = "unknown";
  if (probeComparable(before) && probeComparable(after)) {
    state = canonicalJson(before) === canonicalJson(after) ? "limited_reread_match" : "changed";
    if (state === "limited_reread_match" && !observationsMatch(after, input)) state = "unknown";
  }
  let prior = null;
  if (supplied.prior_snapshot) {
    const status = PRIOR_STATUS.has(supplied.prior_snapshot.status)
      ? supplied.prior_snapshot.status
      : "not_verified";
    prior = {
      id: supplied.prior_snapshot.id ?? null,
      status,
      used_as_verified_state: false,
    };
  }
  return { state, prior, evidence_probe_matched: state === "limited_reread_match" };
}

const BASE_LIMITS = [
  "This audit is not a full-file pass.",
  "Multiple Word context.sync() calls are not one transaction. A matching before/after probe does not prove unread parts stayed stable.",
  "Native tracked-change enumeration is not an exhaustive OOXML inventory. Zero native rows are not proof of no revisions.",
  "Empty Deleted.text is not proof that deleted text is absent.",
  "Range and story getOoxml() results are not the original DOCX ZIP.",
  "Word author fields are source metadata, not verified identities, and are stored separately from the file submitter.",
  "Session-local paragraph IDs and this snapshot id are not permanent cross-file identifiers.",
  "Comment text is review discussion, not an applied revision.",
  "Equal markup counts are not a completeness verdict.",
  "This snapshot cannot authorize a write, including accept, reject, or resolve.",
];

function authoredText(report) {
  return [
    report.summary.headline,
    ...report.summary.limits,
    ...report.claims.map((claim) => claim.statement),
    ...report.findings.map((finding) => finding.detail),
  ].join("\n");
}

function claimsAbsence(report) {
  const text = authoredText(report).replace(/not proof of no revisions/gi, "");
  return /(?:\bno revisions\b|\bno tracked changes\b|\bdocument is clean\b|无修订|没有修订)/i.test(
    text,
  );
}

export function buildRedlineAudit(input = {}) {
  if (!input.snapshot_id || typeof input.snapshot_id !== "string")
    throw new Error("snapshot_id is required");
  if (!input.captured_at || typeof input.captured_at !== "string")
    throw new Error("captured_at is required");
  const kind = input.source?.kind;
  if (!KINDS.has(kind)) throw new Error("source.kind must identify the evidence source");
  const snapshotId = input.snapshot_id;
  const requested = storyScope(input.requested_target || {});
  const limits = [...BASE_LIMITS];
  let file_sha256 = null;
  let file_sha256_source = null;
  let hash_scope = "collected_evidence_only";
  if (kind === "disk_docx" && /^[a-f0-9]{64}$/.test(input.source.file_sha256 || "")) {
    file_sha256 = input.source.file_sha256;
    file_sha256_source = "caller_supplied_not_rehashed";
    hash_scope = "caller_supplied_file_bytes_plus_collected_evidence";
    limits.push(
      "The file hash was supplied by the caller and was not recomputed from DOCX bytes in P1A.",
    );
  } else if (input.source?.file_sha256) {
    limits.push(
      "A file hash was supplied for a source that is not a disk DOCX, so it was not recorded. The evidence hash does not prove unread bytes are unchanged.",
    );
  }
  const submitter =
    typeof input.source?.submitter === "string" && input.source.submitter.length
      ? input.source.submitter
      : null;

  const entries = [];
  const gaps = [];
  let coverageSeq = 0;
  const addEntry = (fields) => {
    const row = {
      coverage_id: `cov:${snapshotId}:${coverageSeq++}`,
      capability: fields.capability,
      scope_label: fields.scope_label,
      target: fields.target ?? null,
      version: fields.version ?? null,
      status: fields.status,
      paging_complete: fields.paging_complete ?? null,
      next_cursor: fields.next_cursor ?? null,
      reason: fields.reason,
      evidence_ids: fields.evidence_ids || [],
      absence_confirmed: fields.status === "confirmed_absent",
    };
    entries.push(row);
    if (row.status !== "read" || row.paging_complete === false) {
      gaps.push({
        coverage_id: row.coverage_id,
        capability: row.capability,
        scope_label: row.scope_label,
        status: row.status,
        reason: row.paging_complete === false && row.status === "read" ? row.reason : row.reason,
      });
    }
    return row;
  };

  const map = input.document_map;
  if (!map) {
    addEntry({
      capability: "document_map",
      scope_label: "document",
      status: "not_read",
      reason: "The document map was not read for this audit call.",
    });
  } else if (map.error) {
    const failed = failureStatus(map.error, map.unsupported);
    addEntry({
      capability: "document_map",
      scope_label: "document",
      status: failed.status,
      reason: failed.reason,
    });
  } else {
    addEntry({
      capability: "document_map",
      scope_label: "document",
      status: "read",
      paging_complete: null,
      reason:
        "Inventory only. complete_file_review on the map is not a content audit and is not promoted.",
      evidence_ids: [],
    });
    for (const gap of map.coverage?.gaps || []) {
      const classified = classifyGap(gap);
      addEntry({
        capability: "map_gap",
        scope_label: gapLabel(gap),
        target: gap.target || null,
        status: classified.status,
        reason: classified.reason,
      });
    }
    for (const story of map.stories || []) {
      const scope = storyScope(story.target || {});
      if (scope.key === requested.key) continue;
      addEntry({
        capability: "story_text",
        scope_label: scope.label,
        target: scope.target,
        status: "not_read",
        paging_complete: false,
        reason:
          "Inventoried by the document map only. Paragraphs, revisions, and comments in this story were not read by this audit call.",
      });
    }
  }

  const storyParagraphs = { Current: [], Original: [] };
  const storyPages = [];
  for (const read of input.story_reads || []) {
    const scope = storyScope(read.target || requested.target);
    const version = read.version || null;
    const successfulPages = (read.pages || []).filter((page) => page && !page.error);
    const pageIds = [];
    if (!read.error) {
      for (const page of successfulPages) {
        const offset = page.offset ?? 0;
        const evidence_id = `page:${snapshotId}:${version || "unversioned"}:${offset}`;
        pageIds.push(evidence_id);
        storyPages.push({
          evidence_id,
          snapshot_id: snapshotId,
          source_story: scope.target,
          version,
          offset,
          next_offset: page.next_offset ?? null,
          total: page.total ?? null,
          addressing: page.addressing ?? null,
          paragraphs: (page.paragraphs || []).map((paragraph) => ({
            index: paragraph.index ?? null,
            id: paragraph.id ?? null,
            style: paragraph.style ?? null,
            text: paragraph.text ?? null,
          })),
        });
      }
    }
    if (read.error) {
      const failed = failureStatus(read.error, read.unsupported);
      addEntry({
        capability: "story_text",
        scope_label: scope.label,
        target: scope.target,
        version,
        status: failed.status,
        paging_complete: false,
        reason: failed.reason,
        evidence_ids: pageIds,
      });
      continue;
    }
    const pages = combinePages(read.pages || []);
    if (version === "Current" || version === "Original")
      storyParagraphs[version] = pages.paragraphs;
    addEntry({
      capability: "story_text",
      scope_label: scope.label,
      target: scope.target,
      version,
      status: pages.status,
      paging_complete: pages.status === "read" ? pages.paging_complete : false,
      next_cursor: pages.next_cursor,
      reason: pages.reason,
      evidence_ids: pageIds,
    });
  }

  const revisions = [];
  const ooxmlEvidence = [];
  const mappings = [];
  let revisionSeq = 0;
  for (const read of input.revision_reads || []) {
    const scope = storyScope(read.target || requested.target);
    if (read.error) {
      const failed = failureStatus(read.error, read.unsupported);
      addEntry({
        capability: "revisions",
        scope_label: scope.label,
        target: scope.target,
        status: failed.status,
        paging_complete: null,
        reason: `${failed.reason} This failure does not establish that revisions are absent.`,
      });
      mappings.push({
        scope_label: scope.label,
        native_row_count: null,
        enumeration_complete: false,
        story_markup: null,
        range_markup: null,
        counts_equal: null,
        count_equality_is_completeness: false,
        status: "not_read",
        note: "The revision read failed or is unsupported. No absence conclusion is available.",
      });
      continue;
    }
    const rows = Array.isArray(read.revisions) ? read.revisions : [];
    const ids = [];
    let rangeMarkup = null;
    let sawRangeMarkup = false;
    rows.forEach((row, index) => {
      const markup = countRevisionMarkup(row.ooxml);
      if (markup) {
        sawRangeMarkup = true;
        rangeMarkup = addMarkup(rangeMarkup, markup);
      }
      const emptyDeleted = row.type === "Deleted" && (row.text === "" || row.text == null);
      const pairedMarkup = !!(markup && markup.ins > 0 && markup.del > 0);
      const limitsForRow = [];
      if (emptyDeleted) limitsForRow.push("empty_deleted_text_is_not_absence");
      if (pairedMarkup) limitsForRow.push("range_ooxml_contains_both_insertion_and_deletion");
      const evidence_id = `rev:${snapshotId}:${revisionSeq}:${index}`;
      ids.push(evidence_id);
      revisions.push({
        evidence_id,
        snapshot_id: snapshotId,
        source_story: scope.target,
        source_part: "word_tracked_changes_collection",
        native_index: index,
        native_type: row.type ?? null,
        word_author: row.author ?? null,
        author_kind: "word_author_metadata",
        date: row.date == null ? null : String(row.date),
        native_text: row.text ?? null,
        native_text_empty: row.text === "" || row.text == null,
        current: row.current ?? null,
        original: row.original ?? null,
        paragraph_ids: Array.isArray(row.paragraph_ids) ? [...row.paragraph_ids] : [],
        paragraph_id_scope: "word_session_local",
        permanent_id: null,
        ooxml: typeof row.ooxml === "string" && row.ooxml ? row.ooxml : null,
        ooxml_kind: "range_getOoxml_not_docx_zip",
        markup,
        mapping_status:
          emptyDeleted || pairedMarkup ? "partial_native_pair" : "native_row_recorded",
        limits: limitsForRow,
      });
    });
    revisionSeq += 1;
    const storyMarkup = countRevisionMarkup(read.story_ooxml);
    const storyTotal = storyMarkup ? storyMarkup.ins + storyMarkup.del : null;
    const rangeTotal = sawRangeMarkup && rangeMarkup ? rangeMarkup.ins + rangeMarkup.del : null;
    let counts_equal = null;
    if (storyTotal !== null) counts_equal = storyTotal === rows.length;
    else if (rangeTotal !== null) counts_equal = rangeTotal === rows.length;
    if (storyTotal !== null && rangeTotal !== null && storyTotal !== rangeTotal)
      counts_equal = false;
    const enumeration_complete = read.enumeration_complete === true;
    const mapping = {
      scope_label: scope.label,
      native_row_count: rows.length,
      enumeration_complete,
      story_markup: storyMarkup,
      range_markup: sawRangeMarkup ? rangeMarkup : null,
      story_ooxml_sha256: typeof read.story_ooxml === "string" ? sha256Hex(read.story_ooxml) : null,
      story_ooxml_characters: typeof read.story_ooxml === "string" ? read.story_ooxml.length : null,
      story_ooxml_kind: typeof read.story_ooxml === "string" ? "story_getOoxml_not_docx_zip" : null,
      counts_equal,
      count_equality_is_completeness: false,
      status: counts_equal === true ? "recorded_not_a_verdict" : "unresolved",
      note: "Markup counts are recorded only. Equality is not a completeness verdict, and a missing native row is not proof that deleted text is absent.",
    };
    const markupTotal = storyMarkup ? storyMarkup.ins + storyMarkup.del : 0;
    if (typeof read.story_ooxml === "string" && read.story_ooxml && markupTotal > 0) {
      const evidence_id = `xml:${snapshotId}:story:${revisionSeq - 1}`;
      ids.push(evidence_id);
      ooxmlEvidence.push({
        evidence_id,
        snapshot_id: snapshotId,
        source_story: scope.target,
        kind: "story_getOoxml_not_docx_zip",
        ooxml: read.story_ooxml,
        markup: storyMarkup,
        native_row_count: rows.length,
        retains_markup_without_native_row: rows.length === 0 || markupTotal !== rows.length,
      });
    }
    mappings.push(mapping);
    const keptDeletion = rows.length === 0 && storyMarkup && storyMarkup.del > 0;
    const reason = enumeration_complete
      ? "The caller marked enumeration complete. P1A still does not treat that flag as a full OOXML inventory."
      : keptDeletion
        ? "The native collection returned zero rows. Story OOXML still contains deletion markup and that XML is kept as evidence. This is not proof of no revisions."
        : rows.length
          ? "Native tracked-change rows were read. enumeration_complete is false, so this is not an exhaustive inventory."
          : "The native collection returned zero rows. enumeration_complete is false, so this is not proof of no revisions.";
    addEntry({
      capability: "revisions",
      scope_label: scope.label,
      target: scope.target,
      status: "read",
      paging_complete: null,
      reason,
      evidence_ids: ids,
    });
    // A successful native read is still not a complete inventory in P1A.
    gaps.push({
      coverage_id: entries.at(-1).coverage_id,
      capability: "revisions",
      scope_label: scope.label,
      status: "read",
      reason,
    });
  }

  const comments = [];
  for (const read of input.comment_reads || []) {
    const scope = storyScope(read.target || requested.target);
    if (read.error) {
      const failed = failureStatus(read.error, read.unsupported);
      addEntry({
        capability: "comments",
        scope_label: scope.label,
        target: scope.target,
        status: failed.status,
        reason: `${failed.reason} This failure does not establish that comments or revisions are absent.`,
      });
      continue;
    }
    const ids = [];
    for (const [index, row] of (read.comments || []).entries()) {
      const anchor = row.anchor;
      const anchor_missing =
        !anchor || typeof anchor.current !== "string" || typeof anchor.original !== "string";
      const anchor_identity_gap = anchor_missing || !anchor.paragraph_ids?.length;
      const evidence_id = `cmt:${snapshotId}:${row.id || `missing-${index}`}`;
      ids.push(evidence_id);
      const replies = (row.replies || []).map((reply, replyIndex) => ({
        evidence_id: `reply:${snapshotId}:${reply.id || `${evidence_id}:${replyIndex}`}`,
        word_reply_id: reply.id ?? null,
        word_author: reply.author ?? null,
        author_kind: "word_author_metadata",
        created: reply.created == null ? null : String(reply.created),
        text: reply.text ?? null,
      }));
      comments.push({
        evidence_id,
        snapshot_id: snapshotId,
        source_story: scope.target,
        source_part: "word_comment_collection",
        word_comment_id: row.id ?? null,
        word_author: row.author ?? null,
        author_kind: "word_author_metadata",
        created: row.created == null ? null : String(row.created),
        text: row.text ?? null,
        resolved: row.resolved ?? null,
        anchor: anchor
          ? {
              paragraph_ids: Array.isArray(anchor.paragraph_ids) ? [...anchor.paragraph_ids] : [],
              paragraph_id_scope: "word_session_local",
              current: anchor.current ?? null,
              original: anchor.original ?? null,
              ooxml: typeof anchor.ooxml === "string" && anchor.ooxml ? anchor.ooxml : null,
              ooxml_kind: "range_getOoxml_not_docx_zip",
            }
          : null,
        anchor_missing,
        anchor_identity_gap,
        replies,
        comment_is_not_a_text_revision: true,
      });
    }
    const commentReason =
      "Comment threads, replies, resolved state, and anchors that this read returned. A comment count is not the thread, and comment text is not an applied revision. This read is not proof that every comment in the file was enumerated.";
    addEntry({
      capability: "comments",
      scope_label: scope.label,
      target: scope.target,
      status: "read",
      reason: commentReason,
      evidence_ids: ids,
    });
    gaps.push({
      coverage_id: entries.at(-1).coverage_id,
      capability: "comments",
      scope_label: scope.label,
      status: "read",
      reason: commentReason,
    });
  }

  const formattingReads = input.formatting_reads || [];
  if (!formattingReads.length) {
    addEntry({
      capability: "range_ooxml",
      scope_label: requested.label,
      target: requested.target,
      status: "not_read",
      paging_complete: false,
      reason:
        "Range OOXML was not requested for this audit call. Revision-range OOXML, when present on a native row, is separate and is still not a DOCX ZIP.",
    });
  }
  for (const read of formattingReads) {
    const scope = storyScope(read.target || requested.target);
    if (read.error) {
      const failed = failureStatus(read.error, read.unsupported);
      addEntry({
        capability: "range_ooxml",
        scope_label: scope.label,
        target: scope.target,
        status: failed.status,
        paging_complete: false,
        reason: failed.reason,
      });
      continue;
    }
    const pages = read.pages || [];
    const failed = pages.find((page) => page?.error);
    if (!pages.length) {
      addEntry({
        capability: "range_ooxml",
        scope_label: scope.label,
        target: scope.target,
        status: "not_read",
        paging_complete: false,
        reason: "No OOXML page was read",
      });
      continue;
    }
    if (failed) {
      const classified = failureStatus(failed.error, failed.unsupported);
      addEntry({
        capability: "range_ooxml",
        scope_label: scope.label,
        target: scope.target,
        status: classified.status,
        paging_complete: false,
        reason: classified.reason,
      });
      continue;
    }
    const last = pages.at(-1);
    const started = (pages[0].ooxml_offset ?? 0) === 0;
    const paging_complete = started && last.next_ooxml_offset == null;
    const chunkIds = [];
    for (const page of pages) {
      if (typeof page.ooxml !== "string") continue;
      const offset = page.ooxml_offset ?? 0;
      const evidence_id = `xml:${snapshotId}:range:${offset}`;
      chunkIds.push(evidence_id);
      ooxmlEvidence.push({
        evidence_id,
        snapshot_id: snapshotId,
        source_story: scope.target,
        kind: "range_getOoxml_not_docx_zip",
        ooxml: page.ooxml,
        ooxml_offset: offset,
        ooxml_length: page.ooxml_length ?? page.ooxml.length,
        next_ooxml_offset: page.next_ooxml_offset ?? null,
      });
    }
    addEntry({
      capability: "range_ooxml",
      scope_label: scope.label,
      target: scope.target,
      status: "read",
      paging_complete,
      next_cursor: last.next_ooxml_offset ?? null,
      reason: paging_complete
        ? "Returned OOXML pages cover the captured range export. This is still not a DOCX ZIP."
        : "OOXML paging is incomplete. The unread remainder is not covered, and this export is not a DOCX ZIP.",
      evidence_ids: chunkIds,
    });
  }

  const anchors = [];
  for (const query of input.anchor_queries || []) {
    if (typeof query !== "string" || !query) continue;
    for (const view of ["Current", "Original"]) {
      const paragraphs = storyParagraphs[view];
      if (
        !paragraphs.length &&
        !(input.story_reads || []).some((read) => read.version === view && !read.error)
      ) {
        anchors.push({
          query,
          view,
          status: "not_read",
          hits: [],
          reason: `${view} paragraphs were not read, so this query was not searched.`,
        });
        continue;
      }
      const hits = locateExact(paragraphs, query);
      anchors.push({
        query,
        view,
        status: hits.length ? "found" : "not_found",
        hits,
        reason: hits.length
          ? "Exact substring matches in paragraphs this audit read. Identical text in different paragraphs stays separate. No Unicode normalization was applied."
          : "Exact text was not found in the paragraphs this audit read. Unread pages were not searched, and no Unicode normalization was applied.",
      });
    }
  }

  const consistency = consistencyOf(input);
  if (consistency.state === "changed") {
    limits.push(
      "The before/after probe changed during this read. Do not treat the collected evidence as verified state; read again.",
    );
  } else if (consistency.state === "unknown") {
    limits.push(
      "Consistency could not be confirmed. Do not treat the collected evidence as verified state.",
    );
  } else {
    limits.push(
      "The before/after probe matched. That is a limited reread check, not a transaction and not verified write state.",
    );
  }
  if (consistency.prior && consistency.prior.status !== "matches_probe") {
    limits.push(
      "The previous snapshot token is expired, changed, or not verified, and was not used as verified state.",
    );
  }

  const claims = [];
  let claimSeq = 0;
  const claim = (fields) => {
    const row = {
      claim_id: `claim:${snapshotId}:${claimSeq++}`,
      statement: fields.statement,
      evidence_ids: fields.evidence_ids || [],
      coverage_ids: fields.coverage_ids || [],
      status: fields.status,
      unresolved: fields.unresolved || [],
    };
    if (row.status === "supported" && !row.evidence_ids.length)
      throw new Error("A supported audit claim requires evidence ids");
    claims.push(row);
  };
  for (const revision of revisions) {
    claim({
      status: "supported",
      evidence_ids: [revision.evidence_id],
      statement: `A native tracked-change row was observed at index ${revision.native_index} with type ${revision.native_type}. Word author metadata is not a verified identity. enumeration_complete is not treated as a full inventory.`,
      unresolved: [
        ...revision.limits,
        "relationship between this native row and every OOXML markup element",
      ],
    });
  }
  for (const entry of entries.filter(
    (row) => row.capability === "revisions" && row.status !== "read",
  )) {
    claim({
      status: "unconfirmed",
      coverage_ids: [entry.coverage_id],
      statement:
        "No revision-absence conclusion is available because this revision read did not succeed.",
      unresolved: ["revision inventory"],
    });
  }
  for (const mapping of mappings.filter(
    (row) => row.status === "unresolved" && row.native_row_count !== null,
  )) {
    claim({
      status: "unconfirmed",
      evidence_ids: [
        ...revisions
          .filter((revision) => storyScope(revision.source_story).label === mapping.scope_label)
          .map((revision) => revision.evidence_id),
        ...ooxmlEvidence
          .filter((item) => storyScope(item.source_story).label === mapping.scope_label)
          .map((item) => item.evidence_id),
      ],
      statement:
        "The relationship between native tracked-change rows and OOXML insertion/deletion markup is unresolved. Count equality is not used as the verdict.",
      unresolved: ["markup mapping"],
    });
  }
  for (const comment of comments) {
    claim({
      status: "supported",
      evidence_ids: [comment.evidence_id, ...comment.replies.map((reply) => reply.evidence_id)],
      statement:
        "This comment thread was observed, including its replies and resolved flag when returned. Comment text is not an applied document revision.",
      unresolved: comment.anchor_missing ? ["comment anchor"] : [],
    });
  }

  const findings = [];
  let findingSeq = 0;
  const finding = (fields) => {
    findings.push({
      finding_id: `find:${snapshotId}:${findingSeq++}`,
      category: fields.category,
      evidence_ids: fields.evidence_ids || [],
      coverage_ids: fields.coverage_ids || [],
      action: fields.action,
      detail: fields.detail,
    });
  };
  for (const gap of gaps) {
    const unread = gap.status !== "read";
    finding({
      category: unread ? "unread" : "insufficient_evidence",
      coverage_ids: [gap.coverage_id],
      action: gap.status === "unsupported" ? "human_decision" : "reread",
      detail: `${gap.scope_label} (${gap.capability}) is ${gap.status}. ${gap.reason}`,
    });
  }
  for (const revision of revisions.filter((row) => row.mapping_status === "partial_native_pair")) {
    finding({
      category: "insufficient_evidence",
      evidence_ids: [revision.evidence_id],
      action: "inspect_ooxml",
      detail:
        "This native row is only part of a replacement or has empty Deleted.text. Missing deleted text was kept unresolved and was not treated as absent.",
    });
  }
  for (const mapping of mappings.filter(
    (row) => row.status === "unresolved" && row.native_row_count !== null,
  )) {
    finding({
      category: "mapping_unresolved",
      evidence_ids: [
        ...revisions
          .filter((revision) => storyScope(revision.source_story).label === mapping.scope_label)
          .map((revision) => revision.evidence_id),
        ...ooxmlEvidence
          .filter((item) => storyScope(item.source_story).label === mapping.scope_label)
          .map((item) => item.evidence_id),
      ],
      action: "inspect_ooxml",
      detail: `Native rows ${mapping.native_row_count}; story markup ${JSON.stringify(mapping.story_markup)}; range markup ${JSON.stringify(mapping.range_markup)}. ${mapping.note}`,
    });
  }
  if (consistency.state !== "limited_reread_match") {
    finding({
      category: "stale_or_inconsistent",
      action: "reread",
      detail:
        consistency.state === "changed"
          ? "Evidence changed between the opening and closing probes. Read again before relying on this report. The previous evidence was not kept as verified state."
          : "Consistency could not be confirmed. This report is limited and is not verified state.",
    });
  }
  if (consistency.prior && consistency.prior.status !== "matches_probe") {
    finding({
      category: "stale_or_inconsistent",
      action: "reread",
      detail: `Snapshot ${consistency.prior.id} is ${consistency.prior.status}. It was not used as verified state.`,
    });
  }
  for (const comment of comments.filter((row) => row.anchor_missing || row.anchor_identity_gap)) {
    finding({
      category: "anchor_gap",
      evidence_ids: [comment.evidence_id],
      action: "reread",
      detail: comment.anchor_missing
        ? "The comment object was returned without a usable anchor. The thread was not dropped, and it was not counted as a text revision."
        : "The comment anchor has no session paragraph id. Text alone was not treated as a unique location.",
    });
  }
  for (const anchor of anchors.filter((row) => row.status === "not_found")) {
    finding({
      category: "anchor_gap",
      action: "human_decision",
      detail: `${anchor.view} query was not found by exact match. ${anchor.reason}`,
    });
  }
  const authors = new Map();
  const addAuthor = (name, field) => {
    if (!name) return;
    const row = authors.get(name) || { name, revision_rows: 0, comments: 0, replies: 0 };
    row[field] += 1;
    authors.set(name, row);
  };
  for (const revision of revisions) addAuthor(revision.word_author, "revision_rows");
  for (const comment of comments) {
    addAuthor(comment.word_author, "comments");
    for (const reply of comment.replies) addAuthor(reply.word_author, "replies");
  }
  if (authors.size) {
    finding({
      category: "author_unverified",
      evidence_ids: [
        ...revisions.map((revision) => revision.evidence_id),
        ...comments.map((comment) => comment.evidence_id),
      ],
      action: "human_decision",
      detail:
        "Word author names are document metadata, not verified people. The file submitter is a separate field and was not copied from these authors.",
    });
  }
  if (
    revisions.some(
      (revision) =>
        revision.native_type === "Added" ||
        revision.native_type === "Deleted" ||
        revision.current !== revision.original,
    )
  ) {
    finding({
      category: "view_confusion",
      evidence_ids: revisions.map((revision) => revision.evidence_id),
      action: "human_decision",
      detail:
        "Current text drops tracked deletions. Original is another projection, not the full revision history. Do not treat either view as the only text.",
    });
  }

  const revisionEntries = entries.filter((row) => row.capability === "revisions");
  const commentEntries = entries.filter((row) => row.capability === "comments");
  const native_revision_rows =
    !input.revision_reads?.length || revisionEntries.some((row) => row.status !== "read")
      ? null
      : revisions.length;
  const comment_threads =
    !input.comment_reads?.length || commentEntries.some((row) => row.status !== "read")
      ? null
      : comments.length;
  const summary = {
    headline: "Read-only redline audit. This is not a full-file audit pass.",
    complete_file_audit: false,
    coverage_pass: false,
    consistency_state: consistency.state,
    evidence_probe_matched: consistency.evidence_probe_matched,
    verified_state: false,
    submitter,
    authors: [...authors.values()],
    native_revision_rows,
    comment_threads,
    replies: comments.reduce((total, comment) => total + comment.replies.length, 0),
    uncovered: gaps.map((gap) => ({
      scope_label: gap.scope_label,
      capability: gap.capability,
      status: gap.status,
      reason: gap.reason,
    })),
    limits,
  };
  const report = {
    schema_version: REDLINE_AUDIT_SCHEMA_VERSION,
    adapter: REDLINE_AUDIT_ADAPTER,
    snapshot_id: snapshotId,
    office_writes: 0,
    model_invoked: false,
    usable_for_writes: false,
    verified_state: false,
    complete_file_audit: false,
    claims_no_revisions: false,
    snapshot: {
      schema_version: REDLINE_AUDIT_SCHEMA_VERSION,
      snapshot_id: snapshotId,
      captured_at: input.captured_at,
      source: {
        kind,
        source_id: input.source.source_id || "unspecified-source",
        submitter,
        file_sha256,
        file_sha256_source,
        hash_scope,
        full_snapshot_sha256_covers: FULL_SNAPSHOT_SHA256_COVERS,
        evidence_sha256_covers: EVIDENCE_SHA256_COVERS,
        read_scope: {
          requested_target: requested.target,
          adapter: REDLINE_AUDIT_ADAPTER,
        },
        adapter: REDLINE_AUDIT_ADAPTER,
      },
      consistency: {
        state: consistency.state,
        full_file_consistent: false,
        evidence_probe_matched: consistency.evidence_probe_matched,
        verified_state: false,
        prior_snapshot: consistency.prior,
        limits,
      },
    },
    document_inventory:
      map?.error || !map
        ? null
        : {
            complete_file_review: map.coverage?.complete_file_review === true,
            promoted_to_full_audit: false,
            stories: (map.stories || []).map((story) => ({
              target: story.target || null,
              status: story.status || null,
              paragraph_count: story.paragraph_count ?? null,
              empty: story.empty ?? null,
            })),
          },
    coverage: {
      complete_file_audit: false,
      entries,
      gaps,
    },
    revisions,
    comments,
    story_pages: storyPages,
    ooxml_evidence: ooxmlEvidence,
    mappings,
    anchors,
    claims,
    findings,
    decisions: [],
    decision_note:
      "P1A does not record approved, applied, verified, or exported decisions. A later decision must not reuse this snapshot as write authority.",
    summary,
  };
  if (claimsAbsence(report))
    throw new Error("AUDIT_LANGUAGE: refused to describe missing evidence as no revisions");
  // Both hashes exclude themselves. evidence_sha256 covers the default MCP
  // projection; full_snapshot_sha256 covers this object, including raw ooxml.
  const fullSnapshotSha256 = sha256Hex(canonicalJson(report));
  const evidenceSha256 = sha256Hex(canonicalJson(projectModelOoxml(report)));
  report.snapshot.source.full_snapshot_sha256 = fullSnapshotSha256;
  report.snapshot.source.evidence_sha256 = evidenceSha256;
  return report;
}
