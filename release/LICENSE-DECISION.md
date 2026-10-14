# License decision by delivery scope — 2026-10-10

Factual outbound-scope review for **0.2.0-preview.3**. Not a jurisdiction-specific
legal opinion, provider approval, trademark clearance or authorization to publish.
The owner confirms this is a personal project, with authority over additions and
assets and no additional company OSS policy. Private history/evidence is not shipped.

## Bottom line

**No unresolved copyright-license blocker was identified for this exact source
snapshot**, retaining its MIT grant and upstream attribution. It contains project
code, authorized icons, synthetic Office fixtures and dependency references;
it contains no installed SDK/CLI, Electron/Chromium, Node, node_modules or App.
Dependency references do not redistribute those products. Applicable account/use
terms still govern installation and execution. This conclusion does not authorize
a push or promise supported installation on another machine.

**A self-contained vendor-containing App remains NOT CLEARED.** The local thin
Swift wrapper is a different artifact: its own executable/icon/config/notices,
pointing to separate installations. Its project notices are now included verbatim;
machine-specific paths, portable onboarding and signing remain separate concerns.

## Decision matrix

| Item                                    | Source-only snapshot                                                                       | Thin local Swift wrapper                                                 | App containing vendor payloads                                                                                                                                   |
| --------------------------------------- | ------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Project MIT grant, original attribution | PASS: full original LICENSE/UPSTREAM retained; owner authority confirmed                   | Project notices copied to Resources and hashed; no vendor payload copied | Preserve project notices and modification provenance                                                                                                             |
| CodeBuddy SDK 0.3.243                   | Referenced, not shipped; missing SDK root LICENSE is not a source-copy blocker here        | External installation, not inside wrapper                                | NEEDS REVIEW: metadata MIT + CLI MIT text are strong evidence, but clarify scope across lib/CLI/vendor files and required notices                                |
| CodeBuddy toybox/zsh                    | Not shipped; no GPL functions established in the inspected zsh directory                   | Not inside wrapper                                                       | Toybox/zsh have actual permissive notices; preserve them. Verify the selected binary's provenance/compiled components; no finding that the entire project is GPL |
| Bundled ripgrep/PCRE2                   | Not shipped; no effect on the project's source license                                     | Not inside wrapper                                                       | NEEDS REVIEW: COPYING is only a dual-license selector; actual build/version/transitive notices must be covered, considering PCRE2's binary-chain exception       |
| Legacy Anthropic SDK/CLI                | References only; custom provider terms, not MIT                                            | External legacy route is not accepted as preview support                 | CONDITIONAL, not universally forbidden: assess actual unmodified CLI, own-user authentication and product/service terms; no credential proxy/resale              |
| Electron/Chromium                       | Not shipped                                                                                | External, not inside wrapper                                             | Preserve Electron license and relevant Chromium/compiled third-party notices; do not relabel the bundle MIT                                                      |
| extract-zip 1.0.5                       | Install-time dependency only; not shipped                                                  | Not inside wrapper                                                       | Its notice gap matters only if this npm/native installer component is actually copied; not an automatic gate for a runtime-only App                              |
| Other npm libraries                     | Lockfile references only; factual inventory retained                                       | Not inside wrapper                                                       | Include actual required licenses/NOTICE for the libraries copied into the artifact; dependency metadata alone is insufficient                                    |
| Names/logos                             | No official affiliation/endorsement claim; owner authority is not a vendor trademark grant | Same limitation                                                          | Branding and provider trademarks require separate, artifact-specific review                                                                                      |

## Confirmed evidence and corrected interpretations

### Project source and assets — permissive, retained obligations

`[user provided / local package]` Original LICENSE is MIT, copyright Leonard Hope;
its bytes are unchanged. UPSTREAM identifies the local derivation without claiming
the private commit is an upstream public commit. Authorized project icons and
sanitized synthetic fixtures remain at their reviewed hashes. No provider library
source, executable payload, credentials, business file or private development
history is added to this delivery. No copyleft-licensed code embedded in the public
snapshot was identified in this bounded follow-up; it is not an exhaustive
provenance audit of every line of earlier development.

### CodeBuddy — permissive indicators, whole-bundle scope not fully settled

`[user provided / local package]` SDK 0.3.243 declares MIT. No root LICENSE is
present; `cli/LICENSE.txt` supplies a Tencent Cloud MIT-form grant. Missing root
text does not itself prove infringement or prohibition. The text's precise scope
over all SDK/vendor files and the actual combined bundle needs confirmation if
redistributed. Rewriting Tencent's notice, applying this repository's MIT grant
to vendor files or upgrading dependencies would not resolve that scope question.

`[web search — verify]` [Official SDK documentation](https://www.codebuddy.ai/docs/cli/sdk)
expressly describes integrating agents into applications and custom interfaces.
The [public service-agreement sample](https://www.codebuddy.cn/document/term)
separately limits service/software transfers and distribution (sections 6.1,
6.3–6.4) and addresses third-party terms (3.5, 16.7). It names an international
contracting entity; it does not establish the agreement applicable to every
China/enterprise account. The SDK-specific hierarchy between the MIT evidence and
service terms is **unresolved**, not a finding that SDK integration is forbidden.
The [web attribution page](https://www.codebuddy.ai/document/open-source-attribution)
lists web libraries, not an exhaustive SDK 0.3.243 payload inventory.

Release posture: users install their own vendor dependencies and use their own
accounts; no shared entitlement, hosted service, credential intermediary or SDK
copy is supplied by this ZIP. Account-specific commercial/service clearance was
NOT_REVIEWED. [Prepared vendor questions](VENDOR-CLARIFICATION.md) have not been sent.

### Bundled native tools — permissive core, artifact-specific obligations

`[user provided / local package]` Toybox has its actual permissive LICENSE. Zsh's
LICENSE is permissive for the core and warns about some separately licensed GPL
functions. This package's zsh directory contains only LICENSE and bin/zsh, no
separate function files. That warning does not establish GPL content in this
binary or convert Draftspect to GPL. Exact binary build/provenance remains open.

The ripgrep vendor directory contains only a short COPYING selector plus
platform binaries/modules, not full MIT/Unlicense text. A **version-only** probe
of arm64-darwin reports ripgrep 14.1.1, revision 4649aa9700, with PCRE2 10.43.
It does not prove an unchanged upstream build or inventory the native Node module.

`[web search — verify]` [ripgrep 14.1.1's MIT notice](https://raw.githubusercontent.com/BurntSushi/ripgrep/14.1.1/LICENSE-MIT)
provides the upstream copyright/permission text; it cannot alone establish all
compiled vendor provenance. [PCRE2 10.43's license](https://raw.githubusercontent.com/PCRE2Project/pcre2/pcre2-10.43/LICENCE)
is BSD-form with an express exemption for downstream software that only includes
a binary library-like package using PCRE2 and does not independently use PCRE2.
Do not automatically impose an extra PCRE2 binary notice on every downstream
wrapper, or treat absence of a separate text as proven noncompliance without
assessing this exception. Ripgrep's own attribution and other component review
remain necessary if its code is copied. None is copied in this source ZIP.

### extract-zip — correct the reach of the missing-text issue

`[user provided / local package]` The locked `@electron-internal/extract-zip@1.0.5`
package and its gitHead-pinned Cargo/package/README declare BSD-2-Clause but lack
a root LICENSE. Electron's install.js imports its native extractor. No import
from this project's app/daemon/taskpane/scripts runtime was found. The thin
builder does not copy node_modules or any Electron npm installation helper.

`[web search — verify]` The [upstream relicensing PR](https://github.com/electron/extract-zip/pull/19),
merged September 10, 2026, explains the missing text and changes BSD metadata to
MIT with a new LICENSE; it describes a from-scratch Rust implementation. That
later change is useful provenance, not an explicit retroactive grant for this
pinned package. Do not silently substitute current main's MIT file or the old
max-mapper JavaScript project's BSD notice. If a future artifact ships this
installer/native module, resolve the pinned notice and compiled crates first;
if it excludes the component, that item is N/A for that artifact, not BLOCKED.

### Anthropic — custom terms, personal use and product integration distinguished

`[user provided / local package]` The pinned SDK/platform notices refer to
Anthropic agreements; they are not MIT/open-source grants. CodeBuddy remains the
accepted preview provider, and no provider SDK or CLI is distributed in this ZIP.

`[web search — verify]` The [October 7 help update](https://support.claude.com/en/articles/15036540-use-the-claude-agent-sdk-with-your-claude-plan)
states that SDK and third-party app usage can still use subscription limits.
The [legal/compliance page](https://code.claude.com/docs/en/legal-and-compliance)
separately permits conditional product integration of the unmodified CLI, requires
users' own billing/authentication, recommends API/cloud authentication for
developers and prohibits third-party login/credential intermediation. Accordingly,
neither “all personal SDK subscription use is banned” nor “any third-party App
may relay subscribers' credentials” is supported. No provider approval of this
specific application is inferred. Recheck current terms before enabling or
advertising a distributed Claude route.

## Directly completed versus remaining work

Completed: preserved source license/attribution; explicit source/vendor distinction;
scope-based release checklist; corrected installer-only and provider-auth claims;
new exact-version evidence; thin-wrapper project-notice copying/hashing and tests.
No vendor package was edited, replaced, downloaded or removed for this follow-up.

Remaining source-publication steps are operational: destination/security channel,
exact staged tree/identity review and explicit push approval. They are not the
unresolved vendor-bundle licensing questions. Another-machine installation,
Windows/Intel support, signing/notarization and provider E2E are separate acceptance
gates, not copyright prerequisites for publishing this narrowly labelled source
preview. Do not broaden source acceptance into those claims.

Remaining vendor-containing App work: decide the exact shipped payloads, obtain
SDK scope clarification where needed, construct an artifact-specific notice/SBOM
set (including selected compiled tools) and assess actual service/auth/branding
terms. A full Chromium/vendor binary audit was NOT_RUN. No unreviewed component
is deemed permissive, and no claim of complete binary clearance is made.

Jurisdiction/account enforceability was not decided. This memo uses actual notices
and dated primary provider guidance; the project owner retains outbound decisions.
See [evidence hashes](LICENSE-EVIDENCE.json), [dependency inventory](dependency-inventory.json),
[tests](TEST-SUMMARY.md) and [checklist](CHECKLIST.md).
