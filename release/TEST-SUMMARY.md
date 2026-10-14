# Candidate test summary — 0.2.0-preview.4 — 2026-10-12

## GitHub publication checks — 2026-10-14

The owner authorized publication execution and source push. The independent
initial commit was created and its actual author/committer identity and 200-file
tree were checked against the approved candidate. GitHub Private Vulnerability
Reporting is enabled and its public entry checked without submitting a report.

Publication documentation is updated, with unchanged runtime, tests, assets and
dependencies. Required offline/format/release and staged-blob checks are repeated
for this final documentation; exact final commit/tree, remote readback and
preservation results are bound in the local publication receipt. Native tests
remain the historical results below; no App/DMG or Windows acceptance is added.

## Public commit identity approval — 2026-10-14

The owner approved the full public author/committer identity with their existing
email. Both identities are verified in the candidate's local Git configuration;
GitHub no-reply is optional and the missing-email gate is closed. Actual first
commit metadata remains a separate pre-publication check. No commit, remote,
reporting-feature activation or push is performed by this local status follow-up.

Only current release/identity documentation and the source allowlist change.
Production, package/dependency files, tests and assets remain unchanged. Required
offline/format/release checks and staged-blob/privacy validation are recorded in
a new dated local receipt; previous native/provider results remain historical.

## Repository and maintainer metadata follow-up

The owner selected Office CodeBuddy and repository brucecbi/office-codebuddy,
plus GitHub Private Vulnerability Reporting. README/package/security/contribution
metadata are updated for that planned destination. The repository and reporting
feature are not created/enabled in this local preparation. The approved Git
display names were recorded locally; the email was pending at the October 12 step
and is now approved as described above.

Production code, launchers, package name, version, dependency lock, ports/GUIDs
and all assets remain unchanged. The required full offline suite, formatting,
release manifest and actual staged-blob checks are rerun for this revision and
bound in its new local receipt. Historical receipts are retained. This change
does not add new native Office/Electron/provider or cross-machine acceptance.

## Local Git publication preparation (documentation-only derivative)

This tree adds the vendor-email public-source verification, privacy scope and
publication checklist to the reviewed source ZIP. Raw private reports are not
included. Production implementation/assets and dependency entries are unchanged.
The manifest and Git index have separate identity checks; no push is authorized.
This preparation reruns the full offline suite (526/526, zero failures/skips),
formatting and release checks with the same pinned dependencies, plus actual
staged-blob hashing and a bounded privacy scan. Private history is not imported.
At this earlier step the email required owner input; the October 14 approval closes
that gate. Final hashes/results are bound in the private local
handoff. Earlier native results below remain historical; this step
does not create new Office/Electron/provider or another-machine acceptance.

## Current follow-up: test isolation only

This is a separate unpublished source candidate from the October 11 preview.3
native-repair ZIP. The production app/daemon/taskpane/scripts/manifests files,
icons and synthetic Office assets remain byte-identical; package version metadata
is updated. Only two test files and current release/testing documentation change.
Dependencies are reused at the exact pinned lockfile versions; no fresh npm ci or
download. Tests use Node 22.23.1, npm 10.9.8 and system LibreSSL 3.3.6 on the same
arm64 Mac. They use private synthetic HOME/state and temporary loopback ports;
no model request, Office registration, installed-plugin update or trust change.

| Layer                                                 | Status  | Actual result and scope                                                                                                                             |
| ----------------------------------------------------- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| Session isolation red regression                      | PASS    | New regression first failed against the old helper (7 pass / 1 fail), proving the inherited-state defect. Historical failed log retained privately  |
| Focused session/shutdown suite                        | PASS    | 14/14, zero failures/skips                                                                                                                          |
| Inherited DRAFTSPECT_STATE_DIR                        | PASS    | Two successive runs, each 8/8; synthetic caller sessions.json and file inventory remain unchanged                                                   |
| Concurrent shutdown test workers                      | PASS    | Two overlapping workers, each 6/6; each parallel-fixture case requires four distinct live bound ports. Intentional same-pair conflict still refuses |
| Full npm test                                         | PASS    | 526/526, zero failures/skips; approximately 3.12 seconds on Node 22.23.1                                                                            |
| npm run format:check                                  | PASS    | Pinned Prettier 3.8.3; final formatting is rechecked after summary/manifest closure                                                                 |
| npm run check:release                                 | PASS    | 198 reviewed allowlisted files; final manifest/source and allowlist-only ZIP integrity are rechecked and bound in the accompanying local receipt    |
| Native Office/Electron / actual provider continuation | NOT_RUN | Test-only change; October 11 results below are historical, not a new preview.4 native run                                                           |
| Fresh install / another Mac / Intel / Windows         | NOT_RUN | Reused locked dependencies; no cross-machine support claim                                                                                          |
| New vendor-containing App clearance                   | NOT_RUN | Prior unresolved payload/license/signing gates remain; no new legal decision or App payload                                                         |

The session helper restores both set and absent state variables after success
and thrown callbacks. Shutdown fixture cleanup awaits owned-child exit before
removing synthetic state. Kernel allocation-to-spawn has a short unavoidable
handoff window; only genuine address conflicts receive bounded test-only retries.
Production recovery policy and port settings remain unchanged.

Raw logs, protected-file checks, package hashes and a separate local acceptance
receipt accompany this source candidate outside the ZIP. They are not an
independent third-party audit, publication authorization or native proof.

## Preserved October 11 preview.3 native repair acceptance (historical)

This exact repair revision was tested on the same arm64 Mac, Word/Excel
16.113.3, Node 22.23.1, CodeBuddy SDK 0.3.243 and installed CLI 2.162.0.
Dependencies were reused from the earlier isolated lockfile installation, not
downloaded again. The owner explicitly authorized synthetic provider continuation
and a new source ZIP after the required tests. No business documents, installed
App, original registrations, certificate trust, global config or remote changed.

| Layer                                                                     | Status  | Actual result and scope                                                                                                                                                                                                                                      |
| ------------------------------------------------------------------------- | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Full offline suite after native Word fix                                  | PASS    | 524/524, zero failures/skips, Node 22.23.1; includes the new reference-lifetime regression                                                                                                                                                                   |
| Word reference-lifetime red/green                                         | PASS    | Failing regression 24/25 before the fix, then 34/34 with red-team Word guards; focused runs used Node 25.8.0                                                                                                                                                 |
| Native Word formatting                                                    | PASS    | Approved font write reads back 11 → 13 with unchanged text; format-only external change refuses the stale approval; separate inner check refuses after an italic-only change, with zero writes                                                               |
| Native Excel preview and writes                                           | PASS    | Literal values, calculated values and relative formulas are shown in the bounded old-content preview; six explicit approved writes/sorts verified; external cell change refuses stale write and preserves the new value                                      |
| Native Word tracked insertion/comment                                     | PASS    | One expressly approved insertion plus one anchored comment; actual native revision/comment reads and saved package checks confirm it occurred once. Native Office author metadata is retained, not globally changed                                          |
| Real serial Word → CLI → Word                                             | PASS    | Actual owned SDK child exited before CLI; installed CLI recalled synthetic history and exited; first return refused changed transcript; explicit stable second confirmation returned historical chat only; next explicit request reread four live paragraphs |
| CLI tool-availability interpretation                                      | PASS    | Installed-code filter and collector tests with positive/negative controls distinguish init inventory from enabled tools. The real CLI continuation executed no tools. No model-wire capture or OS-level isolation claim                                      |
| No automatic replay                                                       | PASS    | History return itself spawned no SDK child or model turn and added no Office calls; final user-authorized turn made only one live read. Exactly one model-origin insertion remains in provider history and saved Word text                                   |
| Saved synthetic fixtures                                                  | PASS    | Word ZIP/registered OOXML/anchors pass; Excel saved values match native stale-refusal readback. This is not byte-identical baseline preservation or an alternate-render pass                                                                                 |
| Final formatting and reviewed source allowlist                            | PASS    | npm format:check and check:release pass; 198 reviewed files plus the self-excluded manifest. The new allowlist-only source ZIP has its separate integrity/hashes gate in the accompanying local receipt                                                      |
| Candidate Electron GUI / desktop entry in this follow-up                  | NOT_RUN | Prior launcher/native evidence is retained separately; this Office-tool fix did not change launcher code                                                                                                                                                     |
| Fresh install / TLS or account onboarding / another Mac / Intel / Windows | NOT_RUN | Existing dependencies, trusted TLS and account were used; no cross-machine installer or Windows support claim                                                                                                                                                |
| Vendor-containing App redistribution                                      | BLOCKED | Earlier exact-payload/vendor/license gates remain unresolved; this source-only follow-up does not clear App/DMG distribution                                                                                                                                 |

### Evidence and preserved failures

- Native Word font attempts initially returned `RESULT_UNKNOWN: GeneralException`.
  A temporary private observer identified `Document._GetObjectByReferenceId`.
  Explicit reads found the text unchanged and size still 11; the UI fenced unknown
  writes. Reloading the same paragraph collection invalidated captured host
  proxies. Reusing it fixes this path without relaxing fresh OOXML or cancellation
  checks. Both failed attempts remain in the private record.
- A private CLI assertion initially failed because init advertised 63 tool names.
  That inventory does not apply `isEnabled`; the installed capability collector
  does. Exact installed-code extraction tests confirm empty-tools denial and
  nonempty positive controls; the actual CLI answer and session history confirm
  context reuse and no additional tool events. The initial failure is preserved,
  not overwritten or bypassed. A diagnostic helper's event-field assertion was
  corrected against the actual transcript schema and its failed run retained.
- Native harness overlays use unique QA GUIDs/ports and copied synthetic bindings,
  plus a private process-exit observer and private QA controls. The repaired
  production Word/Excel tools, approval and handoff logic are tested unchanged.
  This accepts those native paths, not ordinary-menu clean-install activation,
  preview side-by-side support, every Office API, or full agent feature parity.
- Three explicitly submitted Office turns and one CLI continuation used only
  synthetic content; model-internal generation rounds/cost were not measured.
  Excel tests were model-free. No real-provider request was retried automatically.
- Word/Excel were restarted only after their Window menus showed no business
  windows. No caches were cleared. Test registrations/processes are task-owned;
  both QA registrations and owned test processes were removed, with original
  App/alias/registrations/source/old-package hashes and metadata verified. The
  selected protection checks and their limits are in the final separate receipt.
- The new source archive excludes all private evidence, QA overlays/documents,
  provider transcripts, dependencies, binaries and Git history. Old archives and
  acceptance receipts remain unchanged. No push or installed-plugin update.

## Historical October 11 offline repairs (before native follow-up)

This is a new unpublished source revision, not the previous preview.3 ZIP. Version,
locked dependencies, ports/GUIDs, extended-mode semantics and existing installation
are unchanged. The previous source/archive/receipt is preserved. No new ZIP, push,
Office registration, installed-App update, model request or global trust change.
Tests reuse the prior isolated pinned dependencies through a temporary link,
removed before handoff; **not** a fresh npm ci. Node 22.23.1, npm 10.9.8 and system
LibreSSL 3.3.6 only in command PATH, on the previously recorded Mac.

| Layer                                                        | Status  | Actual result and scope                                                                                                                                                                                 |
| ------------------------------------------------------------ | ------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Full installed offline suite                                 | PASS    | 523/523, zero failures/skips; synthetic Office/SDK boundary, temporary files/listeners/owned Node children                                                                                              |
| Factory LibreSSL acceptance                                  | PASS    | Requested two files 21/21 in 0.59 seconds; stdin closed and 20-second command limit; no OpenSSL 3 in PATH                                                                                               |
| Excel old-content approval                                   | PASS    | Literal/calculated values and formulas present; empty state and 200-cell tall/wide cap; full preconditions retained; mock card displays old contents                                                    |
| Word format-only stale approval                              | PASS    | All six generic format channels reject target OOXML-only changes before writes; style-only changes and preparation-race checked; rich/unrelated content positive case                                   |
| Prompt trust boundary                                        | PASS    | Static regression checks all three prompts; no claim that wording alone prevents all prompt injection                                                                                                   |
| CLI return change guard                                      | PASS    | Actual RPC functions plus synthetic transcript writes/inode replacement; first refusal, explicit second confirmation, continued-write invalidation and history-only replay; owned-exit barrier retained |
| SDK absent/corrupt fixture                                   | PASS    | Two focused files: 8 SDK-independent passes, 9 declared skips, no failures; corrupt installed entry rejects instead of skipping; complete installed suite above has no skips                            |
| Token/context/queue/link guards                              | PASS    | Symlink/dangling/directory refusal preserves targets; actual token startup block; private token mode; settled promises evicted, durable write receipts retained                                         |
| Thin launcher artifact                                       | PASS    | Isolated Swift compile, plist lint, ad-hoc signature, native `--check`; 0.2.0 numeric keys and full preview version inspected; synthetic CLI placeholder, no GUI/account                                |
| Final formatting and source allowlist                        | PASS    | npm format:check and check:release pass; 198 allowlisted files plus self-excluded manifest, no new ZIP; exact hashes in the separate receipt                                                            |
| Existing environment and previous candidate protection       | PASS    | Recorded prior hashes/metadata match original preview.3 source/ZIP/receipt, installed App/alias and both Office registrations; production HEAD/status unchanged                                         |
| Native Office / Electron GUI / actual CLI resume             | NOT_RUN | Runtime changed; historical native passes below are not transferred to this repair revision                                                                                                             |
| Fresh install / minimum Node / another Mac / Windows / Intel | NOT_RUN | No dependency download or cross-machine/native support claim                                                                                                                                            |
| Vendor-containing App redistribution                         | BLOCKED | Prior exact-payload/vendor/license gates remain unresolved; this source fix is not binary distribution clearance                                                                                        |

### Preserved failed attempts and corrections

- First restricted-runner LibreSSL/preflight attempt: 19/21, two loopback
  `listen EPERM` failures, no deadlock. Normal runner permission gave the 21/21 pass;
  application security and certificate checks were not disabled.
- Initial new pane regression: 96/97; the synthetic harness had not reached the
  awaiting card after its fixed microtask flush. A second bounded flush corrected
  the test; production timeouts and guards were not relaxed. Targeted rerun 120/120.
- Initial dependency-absence fixture exposed the version adapter's eager SDK
  import; the fixture also omitted two files used by source assertions. Conditional
  test import/explicit SDK-adapter skips and complete fixture inputs repaired the
  check. The production adapter and its exit/version barrier were not changed.
- Patch-assembly validation failures changed no files and were corrected; they
  are private tooling records, not recategorized product passes.

### Review corrections and limits

The original Word body JSON already included paragraph `style`; its missing
protection was direct formatting/review OOXML, not style itself. Node's asynchronous
`execFile` has no `input`/`stdio` option; explicitly ending `child.stdin` plus
LibreSSL's `req -new` covers the actual fixture failure. Apple bundle version keys
remain numeric, with full preview suffix in project metadata, rather than placing
an invalid SemVer suffix in numeric bundle keys. See Apple's
[short-version](https://developer.apple.com/documentation/bundleresources/information-property-list/cfbundleshortversionstring)
and [bundle-version](https://developer.apple.com/documentation/bundleresources/information-property-list/cfbundleversion)
definitions; the generated artifact is checked separately from runtime behavior.

Strictly refusing every release-to-CLI transcript change would also break normal
serial continuation. The owner approved the second-confirmation variant: observed
changes never silently return; continued changes refuse even the old confirmation.
A stable fingerprint is **not** proof an external CLI exited or an exclusive lock.
Concurrent CLI reuse remains unsupported. Historical chat reload never executes
tools/model calls. Live document reread is mandatory after returning.

There is no new native Office rendering or actual provider-backed handoff evidence.
Default SDK absence skips are dependency diagnostics, not a complete test pass.
Format/allowlist tests and launcher `--check` are not native runtime E2E. Local
raw logs, App artifacts, fixture state and private Git metadata stay outside the
approved source tree. No ZIP was regenerated after the changed allowlist.

## Preserved October 10 follow-up: license scope and thin-wrapper notices

The only implementation change is project-notice packaging in the optional local
builder. App/daemon/Office runtime, Swift process logic, manifests, icons and
synthetic Office fixtures are byte-identical to preview.2. Dependency resolution
is unchanged; tests reuse the preceding isolated fresh install through a temporary
link, removed before delivery. This turn is not another fresh npm ci or native
Office E2E, and does not change provider support or previous acceptance limits.

| Layer                                                            | Status  | Actual scope                                                                                                                           |
| ---------------------------------------------------------------- | ------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| Targeted packaging tests                                         | PASS    | 9/9; includes two new tests for verbatim bytes/hashes and missing/empty/symlink refusal                                                |
| Full offline regression                                          | PASS    | 507/507, no failures/skips, Node 22.23.1; mocks/temporary loopback servers, no model requests                                          |
| Thin Swift build/config                                          | PASS    | New isolated App, compile/ad-hoc sign/config check; synthetic executable as CLI placeholder, not an account/GUI acceptance             |
| Notice artifact inspection                                       | PASS    | Three project notice bytes and hashes, exact thin-wrapper file set, dependency inventory hashes and unchanged runtime/lock resolution  |
| Format/release/source ZIP                                        | PASS    | Final formatting, allowlist/hash/member and nested fixture/link/privacy checks, recorded in the accompanying receipt                   |
| Existing environment preservation                                | PASS    | Original repo/App/alias/registration/transport and preview.2 source/ZIP/receipt hashes unchanged                                       |
| Source license scope                                             | PASS    | No unresolved copyright-license blocker identified for exact source-only delivery; not publication approval or general legal clearance |
| Vendor-containing App license clearance                          | BLOCKED | SDK whole-bundle scope and exact payload/compiled notice audit remain unresolved; installer-only omissions assessed only if copied     |
| Actual provider/account agreement                                | NOT_RUN | Public guidance reviewed; no account-specific contract, authentication or model call                                                   |
| Native Office/Electron GUI rerun                                 | NOT_RUN | Runtime unchanged; earlier preview.2 evidence preserved below, not counted as a new native pass                                        |
| Fresh reinstall / another Mac / Windows / Intel / signed release | NOT_RUN | Outside this narrow follow-up; unaccepted installation/distribution claims remain unaccepted                                           |

The separate local receipt is an evidence-bound new file, not a third-party
independent audit or push approval. No vendor files or licenses were silently
changed. Public-source cache misses/timeouts were retained as retrieval limits;
missing text was never presumed permissive. One private patch assembly attempt
failed validation before changing files; the patch was corrected, not recategorized
as a product failure. Format-only tooling initially used Node 25.8.0; required
final checks and the build use Node 22.23.1.

## Preserved preview.2 fresh-install and isolated native checks

Same Mac: macOS 26.6.2 arm64, Word/Excel 16.113.3, Node 22.23.1, npm 10.9.8,
Electron 42.11.3. No push, installed App replacement, certificate trust change,
business-document access or model request. Previous source/ZIP/receipt preserved.

| Layer                                                   | Status           | Exact scope                                                                                                                                                                     |
| ------------------------------------------------------- | ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Fresh dependencies                                      | PASS             | Empty extraction/cache, 115 packages installed from pinned lockfile; inspect-first `--ignore-scripts`, then normal `npm ci`; no reused production node_modules                  |
| Electron binary download                                | PASS             | Explicit bundled installer, new task-local cache, checksum verification retained; rerun after normal ci uses only this turn's cache                                             |
| Offline regression                                      | PASS             | 505/505, zero failures/skips, using fresh packages on Node 22.23.1; mock Office/provider boundary disclosed                                                                     |
| Native Electron single-instance                         | PASS             | 3 scenarios: same-state rejection, different-state isolation, replacement after owner quit; owned children cleaned up                                                           |
| Real Electron main/status window/daemon                 | PASS             | Ready state accurately separates Office/model status; renderer sandbox/context isolation retained; explicit Stop, Start and normal Quit; no Office clients/model                |
| Thin Swift launcher                                     | PASS             | Recompiled from candidate and config/ad-hoc-signature checks; not launched through installed desktop shortcut this round                                                        |
| Word native read/format                                 | PASS             | New synthetic DOCX, real Office.js body read; explicit approval; font 11→13, text unchanged; live readback and saved OOXML format revision verified                             |
| Word disconnect/reconnect                               | PASS             | Side pane reports unavailable connection without claiming confirmed crash; send disabled; reconnect restores sendability without replaying writes/model requests                |
| Excel native read/write/sort                            | PASS             | New synthetic XLSX, real Office.js list/read, 5 explicit approvals, range/formula writes and two sorts; sorted rows and relative formulas verified live and in saved package    |
| Native activation detail                                | PASS with limits | Unique QA GUIDs/ports; Word fixture activation, Excel full quit/restart followed by human ordinary-menu click (automation click was unreliable)                                 |
| Candidate cleanup/preservation                          | PASS             | Both QA registrations removed by exact hash; owned servers/windows closed; original source/refs/status, App/alias, original manifest directories and transport hashes unchanged |
| Full provider-backed installation / CLI resume          | NOT_RUN          | QA replaces model generation with fixed synthetic commands; does not exercise actual account authentication, provider requests or resume                                        |
| Another clean Mac/account / Intel / Windows             | NOT_RUN          | Same-machine isolation does not establish these routes; exact Node minimum 22.12.0 also untested                                                                                |
| Full native feature matrix / `.doc` / sleep-wake / soak | NOT_RUN          | This follow-up is targeted; not a rerun of every comment/revision/format feature or historical failure-injection scenario                                                       |
| Signed/notarized downloadable App                       | NOT_RUN          | No App/DMG published or installed; optional thin launcher remains machine-specific                                                                                              |
| Complete binary licensing                               | BLOCKED          | Exact SDK scope, extract-zip notice/provenance and artifact-specific compiled/vendor obligations remain open                                                                    |

### Isolation and limits

Office native tests use the candidate's production taskpane and tool implementations,
but a model-free QA bridge with explicit test-port/CSP-address and unique-GUID
overlays, plus a synthetic-document identity/marker gate. Existing trusted loopback
TLS was read without copying keys or changing trust. These are native Office tool/
UI results, **not** end-to-end model generation or unmodified public installer
acceptance. Public IDs/ports remain unchanged and side-by-side installation is not
supported. A second clean machine/account and TLS onboarding remain release gates.

Word native API/readback and saved revision markup prove the tested change, not
pixel-perfect layout or lossless formatting of arbitrary documents. Runtime test
documents can acquire local Office author metadata; they and raw test logs/screenshots
are private evidence, never included in the source ZIP. Public fixture files remain
at their reviewed sanitized hashes.

### Failures preserved, not recategorized as passes

- Initial network-restricted fresh install encountered `EPERM`; that attempt was
  stopped and retained. Normal network permission enabled the clean retry. No
  application security setting, certificate check or package integrity was disabled.
- Initial native Electron test driver awaited app readiness at ESM top level,
  preventing entry-module completion. It hit its 90-second test timeout/SIGTERM.
  Correcting only the private driver (static main import plus async run) allowed
  the real candidate main/daemon test to pass. This was not evidence of a new
  production Electron crash or an application runtime fix.
- Already-running Excel did not discover the new QA registration; the first open
  reported a missing add-in. After user-approved full quit/restart, the new ID
  appeared. A human menu click opened the pane; automated menu clicks did not
  reliably activate it. No Office cache or original registration was cleared.
- License-source retrieval encountered a timeout, pinned LICENSE 404 and GitHub
  API 403. Available actual package and pinned public-source evidence were used;
  unavailable material was not inferred to be permissive.

Formatting, allowlist/hash checks and final archive verification are recorded in
the independent receipt accompanying the candidate. Exact source identity is in
approved-files.json; that file excludes itself to avoid a circular hash.

## Preserved preview.1 preparation summary (historical, not current status)

Scope: 0.2.0-preview.1, local source preparation. No remote, publication, native
Office installation, certificate trust change or model request. The public source
does not include raw private execution logs. Exact file identity is recorded in
approved-files.json; the archive and this manifest are hashed in a separate receipt.

Dependencies for these tests were read from an existing local installation at the
lockfile versions through a temporary link. That link is excluded from the archive
and removed from this source directory. This is **not a fresh npm ci test**.

| Layer                                                | Status      | Scope                                                                                                                                   |
| ---------------------------------------------------- | ----------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| Final offline regression                             | PASS        | 505/505, zero failures/skips; mocked Office and injected SDK transports                                                                 |
| Formatting and syntax                                | PASS        | Prettier check; 143 JavaScript module syntax checks; relative documentation links checked                                               |
| Thin Swift launcher build/config check               | PASS        | Compiled, ad-hoc signature/config checks; no installed App update                                                                       |
| Compiled launcher + synthetic child                  | PASS        | 8/8 including isolated native process supervision, explicit environment rejection and preview version metadata; not Electron/Office E2E |
| Bootstrap Word OOXML/ZIP                             | PASS / WARN | Package/schema pass; baseline WARN classified as intended core author metadata change only                                              |
| Exact source/hash/archive review                     | PASS        | Allowlist, hashes, targeted private-pattern scan, ZIP member/extracted hash checks; no dependencies or Git history                      |
| Original checkout/installation preservation          | PASS        | Tracked content, HEAD/refs/status and selected installed App/alias/manifest/transport hashes unchanged                                  |
| Fresh dependency install / minimum Node version      | NOT_RUN     | Existing locked dependencies and Node 25.8.0 used; no new packages downloaded                                                           |
| Candidate native Office / actual provider generation | NOT_RUN     | Avoid disturbing the existing working plugin; no model requests                                                                         |
| Another Mac / Intel / Windows / notarized App        | NOT_RUN     | Outside source-preparation scope                                                                                                        |
| Complete binary redistribution license clearance     | BLOCKED     | Unresolved vendor license scope/text and artifact-specific review; source references only                                               |

## Preserved failures and corrections

- Initial targeted 30/33: two runner loopback `listen EPERM` and one obsolete
  mandatory-unselected-provider assertion. Normal runner permission was used;
  application security remained enabled.
- First full run 492/495: two obsolete gateway dropdown assertions and one obsolete
  implicit legacy-history import assertion. Tests now assert replacement safety rules.
- Later full milestones: 495/495, 499/499 and 503/503; final 505/505 above.
- One private harness invocation used the wrong relative test path. Default CLI
  discovery in the npm build environment required an explicit CLI path.
- The first native synthetic run hit its 5-second harness timeout after the child
  completed. The final run took approximately 5.84 seconds with a 15-second test
  budget and verified normal exit. Production monitor/recovery timeouts did not change.
- The private package inspector needed literal bracket escaping for the ZIP
  member `[Content_Types].xml`; final part-diff and privacy checks passed.

New negative tests cover default MCP-loader non-read behavior, SDK option propagation,
selected-provider import isolation, invalid config fail-closed behavior, conflicting
Office registrations (including symlinks), refusal to overwrite TLS config and
source manifest tampering. Only Word core author metadata changed; all other Word
parts and the entire Excel fixture are unchanged. Rendering/native fixture validation
is NOT_RUN; package validation is not rendering proof.

The preceding development baseline separately had 487/487 offline tests and native
single-Mac Word/Excel/lifecycle/desktop-entry evidence. Those historical results are
not used as candidate native acceptance. Dependency inventory passes establish an
inventory, not complete legal clearance or a vulnerability scan.
