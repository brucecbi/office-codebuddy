# Redline review progress

P1A only: a read-only audit foundation for one Word document with multi-author
revisions. Later phases in `docs/REDLINE-REVIEW-PLAN.md` are not started.

## Complete-read probe check — 2026-10-25

`observationsMatch` used to slice the closing probe to the length of the pages
already read, including when the read was finished. A finished Current read of
only P1=`A` therefore matched a probe of P1=`A`, P2=`C`.

A finished read now compares the whole Current paragraph list: count, order,
and text. An unfinished read still compares only the obtained range, and its
coverage entry stays `paging_complete: false`.

Regressions:

- `a complete read shorter than the closing probe does not match`
- `an empty complete read does not match a non-empty probe`
- `a complete read that agrees with the closing probe matches`
- `an unfinished page that disagrees with the closing probe is not a probe match`
  still covers the partial match and the partial mismatch

Cancellation wording is corrected in the same commit. The guard is honored
between audit phases and between pages. `toolDocumentMap` may finish the
`Word.run` calls inside an enumeration that has already started.

| Command                 | Result                                                                                                                                                                                 |
| ----------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm test`              | PASS. 550 passed, 0 failed, 0 skipped. About 5.16 seconds. Node v22.23.4.                                                                                                              |
| `npm run format:check`  | PASS. Prettier 3.8.3, after this section was formatted.                                                                                                                                |
| `npm run check:release` | FAIL. 203 files visited, 20 errors: 16 `hash-mismatch` and 4 `not-allowlisted`. No missing, symlink, or credential finding. The manifest was not regenerated. Not a product-test pass. |
| Mac Word                | NATIVE_NOT_RUN                                                                                                                                                                         |

## Review follow-up — 2026-10-25

Four findings on `cursor/redline-review-p1a-7aa1`. P1B was not started.

1. Current and Original pages are returned as `story_pages`, each bound to the
   audit `snapshot_id`. Range OOXML requested with `include_ooxml: true` is
   returned as `ooxml_evidence`. Story OOXML that contains `w:ins` or `w:del`
   markup is kept the same way, including when the native collection returned
   zero rows. `coverage.evidence_ids` for those reads point at those rows.
   Regression: `zero native rows keep story XML that contains a deletion and
pages point at real evidence`.
2. Unfinished Current pages are compared with the same slice of the closing
   probe. Regression: `an unfinished page that disagrees with the closing
probe is not a probe match`. Before and after probes are `A`, the page read
   is `B`, and `evidence_probe_matched` is false. A matching prefix still
   matches and still does not make `complete_file_audit` true.
3. `full_snapshot_sha256` is the hash of the taskpane result including raw
   `ooxml`, excluding both hash fields. `evidence_sha256` is the hash of the
   default MCP projection: raw `ooxml` strings removed, `ooxml_omitted` and
   `ooxml_chunk_length` recorded, and `ooxml_access` included when anything was
   omitted, again excluding both hash fields. Regression:
   `default and explicit OOXML projections hash the assembler result the MCP
bridge returns`. It runs `buildRedlineAudit` through `createOfficeBridgeMcp`
   for both `{}` and `{ include_ooxml: true }`.
4. The queue `guard` is passed into `toolAuditRedline` and checked between
   phases and before each story or OOXML page. `CANCELLED_OR_EXPIRED` is
   rethrown at the next check. `toolDocumentMap` may still finish `Word.run`
   calls inside an enumeration that has already started. The regression is
   `cancellation between audit phases stops the next read and a later queued
task still runs`.

| Command                 | Result                                                                                                                                                                                 |
| ----------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm test`              | PASS. 547 passed, 0 failed, 0 skipped. About 5.21 seconds. Node v22.23.4.                                                                                                              |
| `npm run format:check`  | PASS after this section was formatted. Prettier 3.8.3.                                                                                                                                 |
| `npm run check:release` | FAIL. 203 files visited, 20 errors: 16 `hash-mismatch` and 4 `not-allowlisted`. No missing, symlink, or credential finding. The manifest was not regenerated. Not a product-test pass. |
| Mac Word                | NATIVE_NOT_RUN                                                                                                                                                                         |

The 2026-10-24 section below is the first P1A run, not this follow-up.

Baseline observed at start: `main` at `c79441f0a8d727e735c393741dace4b0d79e2573`,
clean worktree, Node v22.23.4. That commit was not used as a rollback target.

## What this batch implements

`office_audit_redline` is a read-only Office tool. One call, for one story
(default body):

- reuses the document map, paged `office_read_story` Current and Original reads,
  the existing revision and comment collectors, and, only when asked, paged
  `office_read_formatting`
- returns a schema version 1 snapshot, revision and comment evidence, a
  coverage report, deterministic claims and findings, and a short summary
- does not accept, reject, resolve, edit, or call a model
- does not register an `office_manage_revisions` snapshot

`complete_file_audit` and `summary.coverage_pass` are false. `verified_state`
and `usable_for_writes` are false. `claims_no_revisions` is false.

## Decisions

- The pure assembler is `taskpane/shared/redline-audit.js`. The live entry is
  `toolAuditRedline` in `taskpane/shared/word-review.js`. It calls the existing
  readers instead of a second Word stack.
- Evidence ids are bound to `snapshot_id`. Paragraph ids stay session-local.
  Neither is a permanent cross-file id.
- A live call uses source kind `live_word`. Its hash is
  `collected_evidence_only`. A caller-supplied file hash is ignored unless the
  source kind is `disk_docx`, and even then P1A does not rehash the file.
  Story `getOoxml()` is counted and hashed, then dropped from the result. It is
  not a DOCX ZIP. Per-row `ooxml` remains and the daemon still omits it unless
  `include_ooxml` is true.
- `include_ooxml: true` also reads paged range OOXML. The default is off.
  `max_ooxml_pages` defaults to 1. `story_limit` defaults to 500 and
  `max_story_pages` defaults to 20.
- Consistency is `limited_reread_match`, `changed`, or `unknown`. A matching
  probe is not a transaction and not full-file proof. Expired or changed prior
  tokens are reported on `prior_snapshot.status` (`expired`, `not_current`,
  `not_verified`, `matches_probe`) and `used_as_verified_state` is forced
  false. Audit snapshots expire after five minutes and are capped at eight.
- `enumeration_complete` from the native reader stays false. A true flag
  supplied to the assembler is kept visible and still does not promote a
  full-file pass. `confirmed_absent` is not used for an empty native list or
  for a story that the map only inventoried.
- Markup counts are recorded. Equal counts use status `recorded_not_a_verdict`.
  Unequal or missing counts use `unresolved`. `count_equality_is_completeness`
  is false. An empty `Deleted.text` row is kept, with
  `empty_deleted_text_is_not_absence`.
- Anchor search is exact `indexOf` on paragraphs this call read. No Unicode
  normalization. Unread pages are not searched.
- The same paragraph text in two paragraph ids stays two evidence rows.
  `submitter` is never copied from `word_author`.
- The taskpane hashes with an in-module SHA-256 so the Office webview does not
  need `node:crypto`. Paragraph text is captured before the story OOXML hash so
  the opening and closing probes share one read order.
- Both read-only lists include the tool: `OFFICE_READ_ONLY_TOOLS` and
  `readOnlyOfficeTools`. It takes the existing read queue path, not the write
  confirmation path. The default office-only policy is unchanged.
- `release/approved-files.json` was not regenerated. AGENTS.md reserves that
  for final review.

## Synthetic example

This input is a fixture, not a Mac Word result. Consistency is `unknown` here
because the example does not supply a before/after probe. The live tool does.
Two native rows against story markup of one `w:ins` and two `w:del` stay
unresolved. The empty deletion is kept. The submitter is not an author.

```json
{
  "complete_file_audit": false,
  "claims_no_revisions": false,
  "verified_state": false,
  "usable_for_writes": false,
  "office_writes": 0,
  "model_invoked": false,
  "snapshot": {
    "source": {
      "kind": "synthetic_fixture",
      "submitter": "Filing Clerk",
      "file_sha256": null,
      "hash_scope": "collected_evidence_only"
    },
    "consistency": { "state": "unknown", "full_file_consistent": false, "verified_state": false }
  },
  "summary": {
    "native_revision_rows": 2,
    "comment_threads": 1,
    "replies": 1,
    "coverage_pass": false
  },
  "revisions": [
    {
      "evidence_id": "rev:example-snapshot:0:0",
      "native_type": "Added",
      "word_author": "甲",
      "native_text": "同意。",
      "paragraph_ids": ["P1"],
      "mapping_status": "partial_native_pair"
    },
    {
      "evidence_id": "rev:example-snapshot:0:1",
      "native_type": "Deleted",
      "word_author": "乙",
      "native_text": "",
      "original": "应当保留的删除",
      "paragraph_ids": ["P2"],
      "mapping_status": "partial_native_pair",
      "limits": ["empty_deleted_text_is_not_absence"]
    }
  ],
  "mapping": {
    "native_row_count": 2,
    "enumeration_complete": false,
    "story_markup": { "ins": 1, "del": 2 },
    "counts_equal": false,
    "count_equality_is_completeness": false,
    "status": "unresolved"
  }
}
```

## Verification

Cloud environment: Linux, Node v22.23.4, npm 10.9.9. `npm ci` installed 116
packages from the lockfile. No `npm audit fix`.

| Command                 | Result                                                                                                                                                        |
| ----------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm test`              | PASS. 543 passed, 0 failed, 0 skipped, 0 cancelled. Duration about 4.94 seconds.                                                                              |
| `npm run format:check`  | PASS. Prettier 3.8.3 reported all matched files use its style.                                                                                                |
| `npm run check:release` | FAIL. 203 files visited, 20 errors: 16 `hash-mismatch` and 4 `not-allowlisted`. No missing, symlink, or credential finding. The manifest was not regenerated. |

Not run, and not counted as passes:

- Native macOS Word, Excel, Electron, launcher, TLS trust, and registration:
  `NATIVE_NOT_RUN`
- Real CodeBuddy or Claude model calls: not run
- Multi-file merge, accept/reject automation, clean-copy export, DOCX ZIP
  import, and a general subagent runtime: out of scope
- GitHub Actions: not added, not triggered

## Mac Word acceptance

Use a new synthetic document. Do not use a business contract. Record the macOS
version, Word version, and the commit SHA.

1. Turn on track changes. Have at least three author names if the machine can
   do that. Add an insertion, a deletion, and a replacement. Put the same short
   sentence `同意。` in two different paragraphs. Add a comment, a reply, and
   one resolved comment.
2. Open the task pane and ask for a read-only audit that calls
   `office_audit_redline` with `submitter` set to a name that is not a Word
   author, such as `Filing Clerk`. Do not accept or reject anything.
3. Confirm `complete_file_audit` is false, `verified_state` is false,
   `office_writes` is 0, `snapshot.source.file_sha256` is null, and
   `hash_scope` is `collected_evidence_only`. Confirm `submitter` and
   `revisions[].word_author` are different fields. Confirm an empty
   `Deleted.text` row is still present. Confirm
   `mappings[].enumeration_complete` is false and
   `count_equality_is_completeness` is false. Confirm the reply and resolved
   flag are present and the comment is not described as an applied edit.
   Confirm inventoried headers, footers, or notes that this call did not read
   are `not_read`, not a claim of no revisions. Confirm Word's revisions and
   comments did not change.
4. Call again with `story_limit: 1` and `max_story_pages: 1` on a document with
   more than one paragraph. Confirm the Current story entry has
   `paging_complete: false` and `coverage_pass` is false.
5. Edit the document after a call, then call again with that `prior_snapshot_id`.
   Confirm `prior_snapshot.status` is `not_current` and
   `used_as_verified_state` is false. A made-up id should be `expired`.
6. If this Word build cannot read a footnote or text box, audit that target and
   confirm the status is `unsupported` or `read_failed`, and
   `summary.native_revision_rows` is null rather than 0.

## Known limits

- One story per call. Other stories are inventory gaps, not reads.
- Native revision enumeration is not an OOXML inventory. Range and story
  OOXML are not the DOCX package.
- The before/after probe can miss a change that reverts in between, and it
  does not prove unread bytes.
- No reviewer model, no multi-file merge, no automatic accept/reject, and no
  clean-copy export.
- This cloud run did not execute the Mac checklist. Status: `NATIVE_NOT_RUN`.

## Next batch

After Mac acceptance of this tool, P1B can add a read-only DOCX package
inventory for the revision and comment parts this audit cannot see. Do not
start from this batch. P1B still must not write through a second channel.
