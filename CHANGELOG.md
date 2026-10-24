# Changelog

## 0.2.0-preview.4 — 2026-10-24 — single-document redline audit

- Added read-only `office_audit_redline` for one Word story. It freezes native
  revision and comment evidence, the Current/Original pages it actually read,
  coverage gaps, and a before/after consistency probe. It does not write,
  accept, reject, resolve, or call a model.
- A failed read, an unsupported story, an unfinished page, or zero native rows
  stays explicit. None of those is a full-file pass or a claim that revisions
  are absent. Empty `Deleted.text` is retained. Word authors stay separate from
  the file submitter. The audit snapshot is not write authority.
- Offline coverage is in `docs/REDLINE-REVIEW-PROGRESS.md`. Native macOS Word
  was not run. `release/approved-files.json` was not regenerated.

## 0.2.0-preview.4 — 2026-10-14 — GitHub source publication

- Owner authorized source publication at brucecbi/office-codebuddy. Created the
  independent initial commit and inspected its real author/committer identity;
  no private development history is imported.
- Established the public repository and enabled GitHub Private Vulnerability
  Reporting; verified the reporting entry and updated installation/security docs.
- Retained original MIT attribution, exact pinned dependencies, runtime and assets.
  Rechecked offline tests, formatting, allowlist and actual staged blobs for the
  publication documentation. No App/DMG, Windows acceptance or new native E2E claim.

## 0.2.0-preview.4 — 2026-10-14 — public commit identity approval

- The owner approved a complete public author/committer identity using their
  existing email. GitHub no-reply is optional for this release and is no longer
  a required publication gate.
- Updated the current checklist, privacy scope and test summary; verified the
  approved identity in the candidate's local Git configuration. The first actual
  commit still requires creation and metadata inspection before publishing.
- Preserved upstream attribution and license, production code and locked
  dependencies. Rechecked the source allowlist and staged tree; no remote or push.

## 0.2.0-preview.4 — 2026-10-12 — repository metadata follow-up

- Selected Office CodeBuddy as the public project name and
  brucecbi/office-codebuddy as its repository; updated README, origin,
  contribution/security instructions and package repository/homepage/issue metadata.
- Selected GitHub Private Vulnerability Reporting; documentation records it as
  awaiting repository creation, feature enablement and availability verification.
- Recorded the approved author/committer display name in local preparation;
  the complete GitHub no-reply email remains pending, so no commit is created.
- Retained the internal package identity required by the tested launchers, pinned
  dependencies, production code, add-in identifiers, assets and upstream license.
  Updated the allowlist and local staged-tree evidence; no public push.

## 0.2.0-preview.4 — 2026-10-12 — local Git publication preparation

- Added public-source evidence for the vendor contact email and closed that
  privacy-review question without deleting vendor attribution or altering old
  source archives/reports. Raw private review material remains outside the tree.
- Prepared a new source-only Git tree with an updated manifest/checklist. This
  documentation revision has a separate identity from the reviewed preview.4 ZIP.
  Application implementation, dependencies, public ports/GUIDs and synthetic
  assets are unchanged. No installed-plugin update or public push.
- Commit author/email, actual repository and private security-reporting channel
  remain explicit owner-confirmation gates; no private Git history is imported.
- Removed only redundant final blank lines from .gitignore and AGENTS.md so the
  new repository's initial staged whitespace check passes; rules are unchanged.

## 0.2.0-preview.4 — 2026-10-12 — test isolation follow-up, unpublished

- Created a separate candidate from the reviewed preview.3 native-repair source;
  older candidates, archives and receipts remain unchanged. Only package version
  metadata, two test files and release/testing documentation change. No dependency
  versions, application implementation, public ports or manifest IDs change.
- Session tests now override and restore DRAFTSPECT_STATE_DIR together with the
  fake HOME/USERPROFILE, including callback failures. Regression probes verify
  inherited caller state stays byte-identical across repeated test runs.
- Shutdown tests allocate two distinct kernel-assigned loopback ports per fixture.
  A bounded three-attempt retry handles only a real address-in-use handoff race;
  the deliberately conflicting second-daemon test bypasses that retry and still
  requires an actionable failure. Owned children exit before fixture deletion.
- Added concurrent-live-fixture regression coverage and separately ran two
  shutdown test workers in parallel. Native Office/Electron and provider tests
  are not rerun for this test-only change; prior results stay historical.
- Repackaging remains source-only and conditional on all required checks passing.
  No installed plugin/App update, trust/global-config change, new dependency
  download, model request or public push is part of this follow-up.

## 0.2.0-preview.3 — 2026-10-11 — native repair follow-up, unpublished

- Fixed a native Word formatting failure introduced by the added inner
  precondition: reloading the paragraph collection invalidated captured host
  references (`Document._GetObjectByReferenceId`). All six generic formatting
  channels now reuse their loaded collection and still reread target OOXML
  immediately before writing. A failing proxy-invalidation regression now passes.
- Revalidated the exact repaired tools in isolated native Word/Excel on one Mac:
  font writes, both stale-format gates, old Excel values/formulas, approved range
  edits, relative-formula sort and stale-cell refusal. Earlier native failures
  remain recorded; uncertain writes were inspected, never replayed automatically.
- Completed authorized synthetic Word → CLI → Word serial continuation with
  actual owned SDK-child exit, changed-transcript refusal, explicit second
  confirmation, history-only return and a fresh native read. Tracked insertion
  and comment anchors were checked in Word and in the saved package.
- Corrected the private CLI acceptance assertion: CodeBuddy 2.162's init `tools`
  list is an inventory, while its enabled-tool checker/collector applies
  `--tools ""`. Installed-code tests include positive controls; no extra tools
  were called. This does not claim OS-level isolation or a model-wire capture.
- Updated compatibility, recovery and current test evidence. Only the source
  archive is regenerated under owner authorization; older archives/receipts,
  the installed plugin and public ports/GUIDs are preserved. No push or App update.

## 0.2.0-preview.3 — 2026-10-11 — independent-review repair revision, unpublished

- Retained the version/locked dependencies, public ports and manifest IDs. This is
  a separate source revision; the previous candidate/archive/receipt is unchanged.
  No ZIP was regenerated, pushed or installed.
- Closed stdin and bounded every OpenSSL test fixture command, added LibreSSL's
  explicit `-new` certificate request and private key permissions. Tests no longer
  require OpenSSL 3; system LibreSSL checks are recorded separately.
- Excel approval previews include old values/formulas, limited to 200 cells with
  explicit truncation/empty state. Full formula preconditions are not truncated.
- Generic Word formatting approvals now compare target paragraph style/OOXML both
  after approval and just before writing. The existing body check already included
  style; the missing protection was direct-format/review-markup-only changes.
- All three system prompts label document/selection content as untrusted data,
  not instructions or authorization; user requests and existing guards prevail.
- Serial CLI return compares the release transcript's stable file fingerprint.
  External changes first refuse return; an explicit second confirmation accepts
  only the still-matching new fingerprint. Continued changes refuse again. This
  preserves normal CLI continuation, but does not prove an external writer exited.
- SDK tests resolve the package entry instead of a fixed node_modules layout;
  only tests needing a genuinely absent SDK explicitly skip. Installed API drift
  still fails. The opt-in control-only script uses the public package entry too.
- Removed the internal startup label, corrected the tracked-policy comment and
  selection description, and evicted settled queue promises without deleting
  persisted write receipts or authorizing edit replay.
- Local builder derives numeric Apple bundle versions and full preview metadata
  from package.json; source hashing rejects symlink entries. Bridge-token/context
  writes reject symlinks before opening/truncation and enforce private token mode.
- Added focused negative/positive regressions and refreshed release/test/recovery
  documentation. Native Office/real CLI/model E2E is not rerun in this revision.

## 0.2.0-preview.3 — 2026-10-10 — local candidate, unpublished

- Separated source-only, local thin-wrapper and vendor-containing App license gates;
  retained the full MIT grant, origin and previously confirmed personal ownership.
- Added exact-version evidence and unsent vendor questions. Missing CodeBuddy root
  text does not block this ZIP without SDK code; whole-bundle scope stays unresolved.
- Scoped extract-zip's notice gap to the install-time component if actually shipped;
  recorded the later upstream relicensing without replacing the pinned notice.
- Recorded bundled ripgrep's incomplete notice selector and PCRE2 version, including
  the upstream downstream-binary exception; avoided a blanket zsh GPL conclusion.
- Corrected Anthropic subscription/SDK wording using its October 7 help update and
  separate legal product-authentication requirements. Legacy route stays unaccepted.
- Thin macOS builder now copies LICENSE, UPSTREAM.md and THIRD-PARTY-NOTICES.md
  verbatim into Resources, records their hashes and refuses invalid notice sources.
  No vendor payload or dependency is added; application runtime behavior is unchanged.
- Added packaging tests; updated release/install/notices and dated test evidence.
  Prior candidates and receipts remain preserved; no push or installed App update.

## 0.2.0-preview.2 — 2026-10-10 — local candidate, unpublished

- Validated a fresh lockfile install with new task-local caches on Node 22.23.1;
  dependencies remain at the same pinned versions. No dependency upgrades.
- Corrected setup instructions for Electron 42.11.3's lazy binary download:
  explicitly invoke its bundled installer before compiling the thin launcher.
- Replaced stale Claude-branded descriptions in both Office manifests with
  provider-neutral, Office-only-default descriptions. IDs and public ports stay
  unchanged; this does not introduce side-by-side installation support.
- Added candidate-specific native Electron and synthetic Office evidence, with
  QA-only port/GUID overlays and untested provider/new-machine routes disclosed.
  Historical failed attempts and preceding receipts remain separate and intact.
- Refreshed dependency evidence from the fresh installation. Confirmed no separate
  zsh function files in this exact SDK's bundled zsh directory; its binary and
  vendor-license scope still require artifact-specific redistribution review.
  Pinned extract-zip provenance does not supply the missing BSD license text.
- Updated platform/install/release claims. Source-only preview remains unpublished;
  complete binary redistribution clearance and clean-other-Mac acceptance remain open.

## 0.2.0-preview.1 — 2026-10-10 — local candidate, unpublished

- New reviewed source snapshot; private history and development evidence stay out
  of the public tree. Upstream MIT attribution remains intact.
- Office-only default: no generic built-in model tools, forwarded user MCP, or
  user settings/hooks. Extended mode and each extension source require opt-in.
- Explicit CodeBuddy environment selection; removed enterprise-environment
  fallback and private gateway model defaults. Only the selected SDK is imported.
- Independent preview state; legacy history is not auto-imported. Generated
  default CLI resume commands also disable builtin tools/user settings/extra MCP.
- Preview manifest filenames and conflict checks prevent overwriting an existing
  installation. Both hosts are checked before writes. Windows installation is
  disabled pending separate native acceptance. No startup installation prompt.
- Sanitized synthetic Word bootstrap author metadata; refreshed its integrity hash.
- Preview-only sidebar preferences and permission-accurate context-file guidance.
- Explicit local TLS generation never installs trust or overwrites an existing
  transport. Default MCP-loader and runtime-option tests exercise the real entry paths.
- Swift configuration checks reject missing/unknown account environments before
  startup, and bounded diagnostic fields admit preview version suffixes. No change
  to production monitor timeouts or automatic recovery budgets.
- Dependencies pinned to already resolved versions; no dependency upgrade.
  Node engine floor reflects Electron's locked install dependencies.
- Replaced inherited setup/platform/authentication claims with preview-specific
  documentation, dependency evidence, test summary and release checklist.

Historical native tests of the development baseline are described separately from
candidate tests. Original native application-registration crash causation remains
unconfirmed; monitoring/recovery is not a guarantee against crashes.
