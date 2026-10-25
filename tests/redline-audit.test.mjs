import { test } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import {
  buildRedlineAudit,
  canonicalJson,
  projectModelOoxml,
  semanticComment,
  semanticRevision,
  sha256Hex,
} from "../taskpane/shared/redline-audit.js";

function base(extra = {}) {
  return {
    snapshot_id: "snap-1",
    captured_at: "2026-10-24T00:00:00.000Z",
    source: { kind: "synthetic_fixture", source_id: "fixture-1", submitter: null },
    requested_target: { story: "body" },
    document_map: null,
    story_reads: [],
    revision_reads: [],
    comment_reads: [],
    formatting_reads: [],
    anchor_queries: [],
    consistency: { before: null, after: null, prior_snapshot: null },
    ...extra,
  };
}

function probe(parts = {}) {
  return {
    revisions: (parts.revisions || []).map(semanticRevision),
    comments: (parts.comments || []).map(semanticComment),
    paragraphs: parts.paragraphs || [],
    story_stable_sha256: parts.story_stable_sha256 ?? null,
    revision_error: null,
    comment_error: null,
    paragraph_error: null,
  };
}

test("evidence digest matches SHA-256 and does not normalize Unicode", () => {
  assert.equal(sha256Hex(""), "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855");
  assert.equal(
    sha256Hex("abc"),
    "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad",
  );
  const sample = "第一条：价格为１００美元。";
  assert.notEqual(sample, sample.normalize("NFKC"));
  assert.equal(sha256Hex(sample), createHash("sha256").update(sample).digest("hex"));
  assert.notEqual(sha256Hex(sample), sha256Hex(sample.normalize("NFKC")));
});

test("partial native replacements keep empty deletions and do not merge repeated text", () => {
  const paired =
    '<w:p><w:del w:author="甲"><w:r><w:delText>旧条款</w:delText></w:r></w:del><w:ins w:author="甲"><w:r><w:t>同意。</w:t></w:r></w:ins></w:p>';
  const revisions = [
    {
      type: "Added",
      text: "同意。",
      current: "同意。",
      original: "",
      author: "甲",
      date: "2026-01-01",
      paragraph_ids: ["P1"],
      ooxml: paired,
    },
    {
      type: "Deleted",
      text: "",
      current: "",
      original: "应当保留的删除",
      author: "乙",
      date: "2026-01-02",
      paragraph_ids: ["P2"],
      ooxml: "",
    },
    {
      type: "Added",
      text: "同意。",
      current: "同意。",
      original: "",
      author: "乙",
      date: "2026-01-03",
      paragraph_ids: ["P9"],
      ooxml: "<w:p><w:ins><w:t>同意。</w:t></w:ins></w:p>",
    },
  ];
  const story =
    "<w:document><w:del></w:del><w:del></w:del><w:del></w:del><w:ins></w:ins></w:document>";
  const input = base({
    source: {
      kind: "synthetic_fixture",
      source_id: "fixture-1",
      submitter: "Filing Clerk",
      file_sha256: "ab".repeat(32),
    },
    revision_reads: [
      {
        target: { story: "body" },
        revisions,
        enumeration_complete: false,
        story_ooxml: story,
      },
    ],
  });
  const audit = buildRedlineAudit(input);
  assert.equal(audit.snapshot.source.submitter, "Filing Clerk");
  assert.equal(audit.snapshot.source.file_sha256, null);
  assert.equal(audit.snapshot.source.hash_scope, "collected_evidence_only");
  assert.match(audit.snapshot.source.evidence_sha256, /^[a-f0-9]{64}$/);
  assert.notEqual(audit.snapshot.source.evidence_sha256, "ab".repeat(32));
  assert.equal(audit.revisions.length, 3);
  const deleted = audit.revisions.find((row) => row.native_type === "Deleted");
  assert.equal(deleted.native_text, "");
  assert.equal(deleted.original, "应当保留的删除");
  assert.equal(deleted.mapping_status, "partial_native_pair");
  assert.ok(deleted.limits.includes("empty_deleted_text_is_not_absence"));
  assert.equal(audit.revisions[0].mapping_status, "partial_native_pair");
  assert.equal(audit.revisions[0].ooxml, paired);
  const repeated = audit.revisions.filter((row) => row.native_text === "同意。");
  assert.equal(repeated.length, 2);
  assert.notEqual(repeated[0].evidence_id, repeated[1].evidence_id);
  assert.deepEqual(
    repeated.map((row) => row.paragraph_ids),
    [["P1"], ["P9"]],
  );
  assert.ok(audit.revisions.every((row) => row.word_author !== "Filing Clerk"));
  assert.equal(audit.mappings[0].enumeration_complete, false);
  assert.deepEqual(audit.mappings[0].story_markup, { ins: 1, del: 3 });
  assert.equal(audit.mappings[0].counts_equal, false);
  assert.equal(audit.mappings[0].count_equality_is_completeness, false);
  assert.equal(audit.mappings[0].status, "unresolved");
  assert.match(audit.mappings[0].note, /not proof that deleted text is absent/);
  const storyXml = audit.ooxml_evidence.find((item) => item.kind === "story_getOoxml_not_docx_zip");
  assert.equal(storyXml.ooxml, story);
  assert.equal(storyXml.snapshot_id, audit.snapshot_id);
  assert.ok(
    audit.coverage.entries
      .find((entry) => entry.capability === "revisions")
      .evidence_ids.includes(storyXml.evidence_id),
  );
  assert.equal(audit.claims_no_revisions, false);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.usable_for_writes, false);
  assert.equal(audit.office_writes, 0);
  assert.equal(audit.model_invoked, false);
  assert.deepEqual(audit.decisions, []);
  assert.ok(audit.findings.some((finding) => finding.category === "mapping_unresolved"));
  assert.ok(audit.findings.some((finding) => finding.category === "insufficient_evidence"));
  assert.ok(audit.findings.some((finding) => finding.category === "author_unverified"));
  assert.ok(audit.findings.some((finding) => finding.category === "view_confusion"));
  assert.ok(
    audit.claims.every((claim) => claim.status !== "supported" || claim.evidence_ids.length),
  );
  assert.equal(
    buildRedlineAudit(input).snapshot.source.evidence_sha256,
    audit.snapshot.source.evidence_sha256,
  );
});

test("equal markup counts are recorded and still are not a completeness verdict", () => {
  const audit = buildRedlineAudit(
    base({
      revision_reads: [
        {
          target: { story: "body" },
          revisions: [],
          enumeration_complete: false,
          story_ooxml: "<w:p></w:p>",
        },
      ],
    }),
  );
  assert.equal(audit.summary.native_revision_rows, 0);
  assert.equal(audit.mappings[0].counts_equal, true);
  assert.equal(audit.mappings[0].status, "recorded_not_a_verdict");
  assert.equal(audit.mappings[0].count_equality_is_completeness, false);
  assert.equal(audit.mappings[0].enumeration_complete, false);
  assert.match(
    audit.summary.uncovered.map((gap) => gap.reason).join("\n"),
    /not proof of no revisions/,
  );
  assert.equal(audit.claims_no_revisions, false);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.summary.coverage_pass, false);
});

test("recorded native markup that outnumbers rows stays unresolved", () => {
  const data = JSON.parse(
    readFileSync(new URL("./fixtures/word-native-accepted-insertion.json", import.meta.url)),
  );
  const audit = buildRedlineAudit(
    base({
      revision_reads: [
        {
          target: { story: "body" },
          revisions: data.before_rows,
          enumeration_complete: false,
          story_ooxml: data.before,
        },
      ],
    }),
  );
  assert.equal(audit.revisions.length, data.before_rows.length);
  assert.equal(audit.mappings[0].native_row_count, 3);
  assert.equal(audit.mappings[0].story_markup.ins, 2);
  assert.equal(audit.mappings[0].story_markup.del, 3);
  assert.equal(audit.mappings[0].counts_equal, false);
  assert.equal(audit.mappings[0].status, "unresolved");
  assert.equal(audit.mappings[0].count_equality_is_completeness, false);
  const storyXml = audit.ooxml_evidence.find((item) => item.kind === "story_getOoxml_not_docx_zip");
  assert.ok(storyXml.ooxml.includes(data.before.slice(0, 180)));
  assert.equal(storyXml.snapshot_id, "snap-1");
  assert.ok(
    audit.coverage.entries
      .find((entry) => entry.capability === "revisions")
      .evidence_ids.includes(storyXml.evidence_id),
  );
  assert.equal(audit.claims_no_revisions, false);
  assert.ok(
    audit.findings.some((finding) => /not proof that deleted text is absent/.test(finding.detail)),
  );
});

test("unfinished story and OOXML pages are coverage gaps, not a pass", () => {
  const audit = buildRedlineAudit(
    base({
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: 1,
              total: 3,
              paragraphs: [{ index: 0, id: "P1", text: "第一条" }],
            },
          ],
        },
      ],
      formatting_reads: [
        {
          target: { story: "body" },
          pages: [{ ooxml_offset: 0, ooxml_length: 50, next_ooxml_offset: 10 }],
        },
      ],
    }),
  );
  const current = audit.coverage.entries.find(
    (row) => row.capability === "story_text" && row.version === "Current",
  );
  const xml = audit.coverage.entries.find((row) => row.capability === "range_ooxml");
  assert.equal(current.paging_complete, false);
  assert.equal(current.next_cursor, 1);
  assert.equal(xml.paging_complete, false);
  assert.equal(audit.summary.native_revision_rows, null);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.summary.coverage_pass, false);
  assert.ok(audit.findings.some((finding) => finding.category === "unread"));
});

test("unsupported and failed revision reads are not reported as zero revisions", () => {
  const unsupported = buildRedlineAudit(
    base({
      revision_reads: [
        {
          target: { story: "footnote", note: 0 },
          error:
            "UNSUPPORTED_API: WordApi 1.6 is required; this part was not inspected or changed.",
          unsupported: true,
        },
      ],
    }),
  );
  const failed = buildRedlineAudit(
    base({
      revision_reads: [{ target: { story: "body" }, error: "Revision content was not populated" }],
    }),
  );
  for (const audit of [unsupported, failed]) {
    assert.equal(audit.summary.native_revision_rows, null);
    assert.equal(audit.revisions.length, 0);
    assert.equal(audit.claims_no_revisions, false);
    assert.equal(audit.complete_file_audit, false);
    assert.ok(
      audit.claims.some(
        (claim) => claim.status === "unconfirmed" && /absence conclusion/.test(claim.statement),
      ),
    );
  }
  assert.equal(
    unsupported.coverage.entries.find((row) => row.capability === "revisions").status,
    "unsupported",
  );
  assert.equal(
    failed.coverage.entries.find((row) => row.capability === "revisions").status,
    "read_failed",
  );
});

test("a changed probe and an expired token are not verified state", () => {
  const audit = buildRedlineAudit(
    base({
      consistency: {
        before: probe({ paragraphs: [{ id: "P1", text: "甲" }] }),
        after: probe({ paragraphs: [{ id: "P1", text: "乙" }] }),
        prior_snapshot: { id: "old-token", status: "expired", used_as_verified_state: true },
      },
    }),
  );
  assert.equal(audit.snapshot.consistency.state, "changed");
  assert.equal(audit.snapshot.consistency.evidence_probe_matched, false);
  assert.equal(audit.snapshot.consistency.full_file_consistent, false);
  assert.equal(audit.verified_state, false);
  assert.equal(audit.snapshot.consistency.prior_snapshot.used_as_verified_state, false);
  assert.equal(audit.snapshot.consistency.prior_snapshot.status, "expired");
  assert.equal(audit.usable_for_writes, false);
  assert.ok(
    audit.findings.some(
      (finding) => finding.category === "stale_or_inconsistent" && /expired/.test(finding.detail),
    ),
  );
});

test("a matching reread still does not promote an expired token or a map flag", () => {
  const paragraphs = [{ index: 0, id: "P1", text: "稳定条款" }];
  const seen = probe({
    paragraphs: [{ id: "P1", text: "稳定条款" }],
    story_stable_sha256: "abc",
  });
  const input = base({
    document_map: {
      coverage: {
        complete_file_review: true,
        gaps: [
          {
            component: "shapes",
            error: "WordApiDesktop 1.2 unavailable; text boxes/floating shapes not enumerated.",
          },
        ],
      },
      stories: [
        {
          target: { story: "footer", section: 0, variant: "Primary" },
          status: "inventoried_not_read",
          empty: true,
          paragraph_count: 1,
        },
      ],
    },
    story_reads: [
      {
        target: { story: "body" },
        version: "Current",
        pages: [{ offset: 0, next_offset: null, total: 1, paragraphs }],
      },
    ],
    revision_reads: [
      {
        target: { story: "body" },
        revisions: [],
        enumeration_complete: true,
        story_ooxml: "<w:p></w:p>",
        story_stable_sha256: "abc",
      },
    ],
    comment_reads: [{ target: { story: "body" }, comments: [] }],
    consistency: {
      before: seen,
      after: seen,
      prior_snapshot: { id: "old-token", status: "expired", used_as_verified_state: true },
    },
  });
  const audit = buildRedlineAudit(input);
  assert.equal(audit.snapshot.consistency.state, "limited_reread_match");
  assert.equal(audit.snapshot.consistency.evidence_probe_matched, true);
  assert.equal(audit.verified_state, false);
  assert.equal(audit.complete_file_audit, false);
  assert.equal(audit.document_inventory.complete_file_review, true);
  assert.equal(audit.document_inventory.promoted_to_full_audit, false);
  assert.equal(input.document_map.coverage.complete_file_review, true);
  assert.equal(audit.snapshot.consistency.prior_snapshot.used_as_verified_state, false);
  assert.equal(audit.mappings[0].enumeration_complete, true);
  assert.match(
    audit.coverage.entries.find((row) => row.capability === "revisions").reason,
    /does not treat that flag as a full OOXML inventory/,
  );
  const footer = audit.coverage.entries.find(
    (row) => row.scope_label === "footer / section 0 / Primary",
  );
  assert.equal(footer.status, "not_read");
  assert.equal(footer.absence_confirmed, false);
  assert.equal(
    audit.coverage.entries.find((row) => row.scope_label === "map:shapes").status,
    "unsupported",
  );
});

test("comment replies and missing anchors are preserved, including text that mentions revisions", () => {
  const audit = buildRedlineAudit(
    base({
      comment_reads: [
        {
          target: { story: "body" },
          comments: [
            {
              id: "C1",
              author: "甲",
              created: "2026-01-01T00:00:00.000Z",
              text: "请确认",
              resolved: true,
              anchor: {
                paragraph_ids: ["P1"],
                current: "条款",
                original: "条款",
                ooxml: "<w:p/>",
              },
              replies: [
                {
                  id: "R1",
                  author: "乙",
                  created: "2026-01-02T00:00:00.000Z",
                  text: "已看",
                },
              ],
            },
            {
              id: "C2",
              author: "丙",
              created: "2026-01-03T00:00:00.000Z",
              text: "There are no revisions in clause 4",
              resolved: false,
              anchor: null,
              replies: [],
            },
          ],
        },
      ],
    }),
  );
  assert.equal(audit.summary.comment_threads, 2);
  assert.equal(audit.summary.replies, 1);
  assert.equal(audit.comments[0].resolved, true);
  assert.equal(audit.comments[0].replies[0].text, "已看");
  assert.equal(audit.comments[0].comment_is_not_a_text_revision, true);
  assert.equal(audit.comments[1].text, "There are no revisions in clause 4");
  assert.equal(audit.comments[1].anchor_missing, true);
  assert.equal(audit.claims_no_revisions, false);
  assert.ok(audit.findings.some((finding) => finding.category === "anchor_gap"));
  assert.ok(
    audit.claims.some(
      (claim) =>
        claim.status === "supported" && /not an applied document revision/.test(claim.statement),
    ),
  );
});

test("exact anchors keep repeated clauses and refuse lossy Unicode folding", () => {
  const full = "第一条：价格为１００美元。";
  const audit = buildRedlineAudit(
    base({
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: null,
              total: 2,
              paragraphs: [
                { index: 0, id: "P1", text: `${full}同意。` },
                { index: 1, id: "P2", text: "同意。" },
              ],
            },
          ],
        },
      ],
      anchor_queries: ["同意。", full, full.normalize("NFKC")],
    }),
  );
  const current = (query) =>
    audit.anchors.find((row) => row.view === "Current" && row.query === query);
  assert.equal(current("同意。").hits.length, 2);
  assert.deepEqual(
    current("同意。").hits.map((hit) => hit.paragraph_id),
    ["P1", "P2"],
  );
  assert.equal(current(full).status, "found");
  assert.equal(current(full).hits.length, 1);
  assert.equal(current(full.normalize("NFKC")).status, "not_found");
  assert.match(current(full.normalize("NFKC")).reason, /no Unicode normalization/i);
  const original = audit.anchors.find((row) => row.view === "Original" && row.query === "同意。");
  assert.equal(original.status, "not_read");
});

test("a caller-supplied disk hash stays separate from the evidence hash", () => {
  const file_sha256 = "cd".repeat(32);
  const audit = buildRedlineAudit(
    base({
      source: { kind: "disk_docx", source_id: "contract.docx", file_sha256, submitter: "Clerk" },
    }),
  );
  assert.equal(audit.snapshot.source.file_sha256, file_sha256);
  assert.equal(audit.snapshot.source.file_sha256_source, "caller_supplied_not_rehashed");
  assert.equal(
    audit.snapshot.source.hash_scope,
    "caller_supplied_file_bytes_plus_collected_evidence",
  );
  assert.notEqual(audit.snapshot.source.evidence_sha256, file_sha256);
  assert.equal(audit.snapshot.source.submitter, "Clerk");
});

test("the audit tool is registered read-only and omits raw OOXML unless requested", async () => {
  const taskpane = readFileSync(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const tools = readFileSync(new URL("../daemon/office-tools.mjs", import.meta.url), "utf8");
  assert.match(taskpane, /case "office_audit_redline"/);
  assert.match(tools, /"office_audit_redline"/);
  const { createOfficeBridgeMcp } = await import("../daemon/office-tools.mjs");
  const { mcp } = await import("../daemon/runtime/runtime-codebuddy.mjs");
  const raw = "<w:del>" + "旧条款".repeat(40) + "</w:del>";
  const server = createOfficeBridgeMcp(
    {
      callTaskpaneTool: async () => ({
        snapshot_id: "snap",
        complete_file_audit: false,
        claims_no_revisions: false,
        revisions: [{ evidence_id: "rev:snap:0:0", ooxml: raw, native_text: "" }],
      }),
    },
    "word",
    "pane",
    mcp,
  );
  const handler = server.instance._registeredTools.office_audit_redline.handler;
  const hidden = JSON.parse((await handler({}, {})).content[0].text);
  assert.equal(hidden.revisions[0].ooxml, undefined);
  assert.equal(hidden.revisions[0].ooxml_omitted, true);
  assert.equal(hidden.complete_file_audit, false);
  assert.match(hidden.ooxml_access, /omits raw OOXML only/);
  const shown = JSON.parse((await handler({ include_ooxml: true }, {})).content[0].text);
  assert.equal(shown.revisions[0].ooxml, raw);
  assert.equal(
    server.instance._registeredTools.office_audit_redline.inputSchema.safeParse({
      include_ooxml: true,
      submitter: "Filing Clerk",
      anchor_queries: ["同意。"],
    }).success,
    true,
  );
});

function evidenceIndex(audit) {
  const ids = new Set();
  for (const page of audit.story_pages) ids.add(page.evidence_id);
  for (const revision of audit.revisions) ids.add(revision.evidence_id);
  for (const comment of audit.comments) {
    ids.add(comment.evidence_id);
    for (const reply of comment.replies) ids.add(reply.evidence_id);
  }
  for (const item of audit.ooxml_evidence) ids.add(item.evidence_id);
  return ids;
}

test("zero native rows keep story XML that contains a deletion and pages point at real evidence", () => {
  const deleted =
    '<w:p><w:del w:author="甲"><w:r><w:delText>已删除条款</w:delText></w:r></w:del></w:p>';
  const audit = buildRedlineAudit(
    base({
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: null,
              total: 1,
              paragraphs: [{ index: 0, id: "P1", style: "Normal", text: "现行条款" }],
            },
          ],
        },
        {
          target: { story: "body" },
          version: "Original",
          pages: [
            {
              offset: 0,
              next_offset: null,
              total: 1,
              paragraphs: [{ index: 0, id: "P1", style: "Normal", text: "已删除条款" }],
            },
          ],
        },
      ],
      revision_reads: [
        {
          target: { story: "body" },
          revisions: [],
          enumeration_complete: false,
          story_ooxml: deleted,
        },
      ],
      formatting_reads: [
        {
          target: { story: "body" },
          pages: [
            {
              ooxml: "<w:del><w:delText>范围删除</w:delText></w:del>",
              ooxml_offset: 0,
              ooxml_length: 48,
              next_ooxml_offset: null,
            },
          ],
        },
      ],
    }),
  );
  assert.equal(audit.summary.native_revision_rows, 0);
  assert.equal(audit.revisions.length, 0);
  assert.equal(audit.claims_no_revisions, false);
  assert.equal(audit.complete_file_audit, false);
  const storyXml = audit.ooxml_evidence.find((item) => item.kind === "story_getOoxml_not_docx_zip");
  assert.ok(storyXml.ooxml.includes("已删除条款"));
  assert.equal(storyXml.snapshot_id, audit.snapshot_id);
  assert.equal(storyXml.retains_markup_without_native_row, true);
  const rangeXml = audit.ooxml_evidence.find((item) => item.kind === "range_getOoxml_not_docx_zip");
  assert.ok(rangeXml.ooxml.includes("范围删除"));
  assert.equal(rangeXml.snapshot_id, audit.snapshot_id);
  const revisionCoverage = audit.coverage.entries.find((entry) => entry.capability === "revisions");
  assert.ok(revisionCoverage.evidence_ids.includes(storyXml.evidence_id));
  const known = evidenceIndex(audit);
  for (const entry of audit.coverage.entries) {
    for (const id of entry.evidence_ids) assert.ok(known.has(id), id);
  }
  for (const version of ["Current", "Original"]) {
    const coverage = audit.coverage.entries.find(
      (entry) => entry.capability === "story_text" && entry.version === version,
    );
    assert.equal(coverage.evidence_ids.length, 1);
    const page = audit.story_pages.find((item) => item.evidence_id === coverage.evidence_ids[0]);
    assert.equal(page.snapshot_id, audit.snapshot_id);
    assert.equal(page.version, version);
    assert.equal(page.paragraphs.length, 1);
  }
  const rangeCoverage = audit.coverage.entries.find((entry) => entry.capability === "range_ooxml");
  assert.deepEqual(rangeCoverage.evidence_ids, [rangeXml.evidence_id]);
});

test("an unfinished page that disagrees with the closing probe is not a probe match", () => {
  const matched = probe({
    paragraphs: [
      { id: "P1", text: "A" },
      { id: "P2", text: "A2" },
    ],
  });
  const agreed = buildRedlineAudit(
    base({
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: 1,
              total: 2,
              paragraphs: [{ index: 0, id: "P1", text: "A" }],
            },
          ],
        },
      ],
      revision_reads: [{ target: { story: "body" }, revisions: [], enumeration_complete: false }],
      comment_reads: [{ target: { story: "body" }, comments: [] }],
      consistency: { before: matched, after: matched, prior_snapshot: null },
    }),
  );
  assert.equal(agreed.snapshot.consistency.evidence_probe_matched, true);
  assert.equal(agreed.story_pages[0].paragraphs[0].text, "A");
  assert.equal(
    agreed.coverage.entries.find((entry) => entry.version === "Current").paging_complete,
    false,
  );

  const disagreed = buildRedlineAudit(
    base({
      snapshot_id: "snap-mismatch",
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: 1,
              total: 2,
              paragraphs: [{ index: 0, id: "P1", text: "B" }],
            },
          ],
        },
      ],
      revision_reads: [{ target: { story: "body" }, revisions: [], enumeration_complete: false }],
      comment_reads: [{ target: { story: "body" }, comments: [] }],
      consistency: { before: matched, after: matched, prior_snapshot: null },
    }),
  );
  assert.equal(disagreed.snapshot.consistency.evidence_probe_matched, false);
  assert.notEqual(disagreed.snapshot.consistency.state, "limited_reread_match");
  assert.equal(disagreed.story_pages[0].paragraphs[0].text, "B");
  assert.equal(disagreed.story_pages[0].snapshot_id, "snap-mismatch");
  assert.equal(disagreed.complete_file_audit, false);
});

test("default and explicit OOXML projections hash the assembler result the MCP bridge returns", async () => {
  const audit = buildRedlineAudit(
    base({
      revision_reads: [
        {
          target: { story: "body" },
          revisions: [],
          enumeration_complete: false,
          story_ooxml: "<w:p><w:del><w:delText>已删除条款</w:delText></w:del></w:p>",
        },
      ],
      story_reads: [
        {
          target: { story: "body" },
          version: "Current",
          pages: [
            {
              offset: 0,
              next_offset: null,
              total: 1,
              paragraphs: [{ index: 0, id: "P1", text: "现行" }],
            },
          ],
        },
      ],
    }),
  );
  assert.match(audit.snapshot.source.full_snapshot_sha256_covers, /including raw ooxml/);
  assert.match(audit.snapshot.source.evidence_sha256_covers, /raw ooxml strings removed/);
  assert.notEqual(
    audit.snapshot.source.evidence_sha256,
    audit.snapshot.source.full_snapshot_sha256,
  );
  const withoutHashes = (value) => {
    const copy = structuredClone(value);
    delete copy.snapshot.source.evidence_sha256;
    delete copy.snapshot.source.full_snapshot_sha256;
    return copy;
  };
  assert.equal(
    sha256Hex(canonicalJson(withoutHashes(audit))),
    audit.snapshot.source.full_snapshot_sha256,
  );
  assert.equal(
    sha256Hex(canonicalJson(projectModelOoxml(withoutHashes(audit)))),
    audit.snapshot.source.evidence_sha256,
  );

  const { createOfficeBridgeMcp } = await import("../daemon/office-tools.mjs");
  const { mcp } = await import("../daemon/runtime/runtime-codebuddy.mjs");
  const server = createOfficeBridgeMcp(
    { callTaskpaneTool: async () => audit },
    "word",
    "pane",
    mcp,
  );
  const handler = server.instance._registeredTools.office_audit_redline.handler;
  const slim = JSON.parse((await handler({}, {})).content[0].text);
  const full = JSON.parse((await handler({ include_ooxml: true }, {})).content[0].text);
  assert.equal(slim.ooxml_evidence[0].ooxml, undefined);
  assert.equal(slim.ooxml_evidence[0].ooxml_omitted, true);
  assert.equal(slim.story_pages[0].paragraphs[0].text, "现行");
  assert.match(slim.ooxml_access, /omits raw OOXML only/);
  assert.equal(sha256Hex(canonicalJson(withoutHashes(slim))), slim.snapshot.source.evidence_sha256);
  assert.ok(full.ooxml_evidence[0].ooxml.includes("已删除条款"));
  assert.equal(
    sha256Hex(canonicalJson(withoutHashes(full))),
    full.snapshot.source.full_snapshot_sha256,
  );
  assert.equal(full.snapshot.source.evidence_sha256, audit.snapshot.source.evidence_sha256);
  assert.equal(
    full.snapshot.source.full_snapshot_sha256,
    audit.snapshot.source.full_snapshot_sha256,
  );
  assert.notEqual(
    sha256Hex(canonicalJson(withoutHashes(full))),
    full.snapshot.source.evidence_sha256,
  );
});
