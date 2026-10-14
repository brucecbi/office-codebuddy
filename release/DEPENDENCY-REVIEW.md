# Dependency and outbound-source review

Date: 2026-10-10. Scope: source-only candidate; no vendor libraries, SDK/CLI
executables, node_modules, installed App or binary installer redistributed.
This is evidence-based release preparation, not a blanket legal clearance.

Preview.3 separates [the three delivery scopes](LICENSE-DECISION.md). No unresolved
copyright-license blocker is identified for the exact source snapshot. The narrow
thin-wrapper notice fix does not clear an App containing external vendor products.

## Repository and assets

- **Upstream source:** actual LICENSE inspected; MIT, copyright Leonard Hope.
  Original text retained without alteration. Source attribution is in UPSTREAM.md.
- **Local additions/assets:** project owner explicitly confirms personal ownership/
  authority to open-source and no additional company OSS policy. Repository MIT
  terms are retained; public posting remains separately unapproved in this task.
- **Synthetic Office fixtures:** initialized for tests, not business documents.
  Word author metadata sanitized; only `docProps/core.xml` changed. No vendor
  Office binaries or runtime libraries are included.
- **Icons:** retained project assets; owner's publication authority confirmed.
  Microsoft/Tencent/Anthropic/Apple affiliation or brand permission is not claimed.

## Referenced dependencies

[dependency-inventory.json](dependency-inventory.json) lists all **122** non-root
lockfile entries with resolved versions, integrity and actual license-file hashes
where installed. The preview.2 follow-up rechecked a fresh locked install: **115**
packages are installed locally; **113** have a root
license/notice text file. The remaining seven entries are optional platform packages
not installed on this Mac. No platform packages were fetched to fill the gap.

Metadata families: MIT 98, ISC 9, BSD-2-Clause 2, BSD-3-Clause 2, Apache-2.0 2,
and vendor-specific/see-license declarations 9. These are metadata counts, not
113/122 independent legal approvals or confirmation that all packages are OSI-licensed.

| Component                                                      | Actual-file evidence / action                                                                                                                                                                                                                                                           |
| -------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@tencent-ai/agent-sdk@0.3.243`                                | Package metadata says MIT; no top-level LICENSE shipped in the inspected package. Its `cli/LICENSE.txt` contains an MIT-form grant by Tencent Cloud. Do not assume that a CLI-scoped text conclusively covers every SDK file; clarify before bundling.                                  |
| CodeBuddy bundled toybox                                       | Actual vendor LICENSE is a permissive grant; preserve text if redistributed. Not included in source delivery.                                                                                                                                                                           |
| CodeBuddy bundled ripgrep / PCRE2                              | COPYING selects MIT or Unlicense but is not their full text. Arm64-Darwin version-only probe reports ripgrep 14.1.1 with PCRE2 10.43. Native provenance/notice review, including PCRE2's binary-chain exception, is needed only if the relevant payload is copied.                      |
| CodeBuddy bundled zsh                                          | Exact fresh vendor directory contains only LICENSE and bin/zsh, no separate shell-function files. The license's warning does not establish that GPL functions are distributed here. Compiled contents/provenance of the binary remain unreviewed; SDK metadata is not binary clearance. |
| `@anthropic-ai/claude-agent-sdk@0.2.141` and platform payloads | Actual LICENSE refers to Anthropic legal agreements; do not classify these as MIT/open source. Retained as an install-time dependency for the legacy adapter, not shipped in the source archive. Review applicable authentication and distribution terms independently.                 |
| `electron@42.11.3`                                             | Actual root MIT-form license inspected. Its downloaded distribution also has Chromium third-party notices; this review does not clear every component of a distributable Electron App.                                                                                                  |
| `ws@8.21.3`, `zod@4.6.1`, `prettier@3.8.3`                     | Actual license files inventoried; preserve their notices if distributing their code. Current artifact references them but does not include them.                                                                                                                                        |
| `@electron-internal/extract-zip@1.0.5`                         | Electron's installer imports it; this project's runtime does not. The pinned BSD metadata lacks text; a later upstream PR adds MIT. Review if actually bundled, not as a blanket source/App gate. No current-main notice is silently substituted.                                       |
| Seven non-arm64-Darwin Claude platform payloads                | Not installed; actual-text review NOT_RUN. Not shipped. No cross-platform binary clearance claimed.                                                                                                                                                                                     |

## Release decision

The original source MIT notice and ownership confirmation support preparing this
source derivative. Referencing an install-time dependency is distinct from copying
its code into this source archive. Do not relicense provider packages or describe
the entire installed dependency tree as MIT/open source.

**Vendor-containing App redistribution is NOT CLEARED.** SDK scope/native payload
and Chromium obligations depend on the files actually shipped. The local thin
wrapper ships no such files and now retains its own project notices. Users obtain
dependencies and model access
under their own applicable agreements. No shared credentials or hosted subscription
proxy is supplied.

See [fresh-install license follow-up](DEPENDENCY-FOLLOWUP.md) for the pinned-version
evidence and narrower conclusions. A missing root license file is not by itself a
finding of infringement, nor a reason to relicense or indiscriminately remove the
dependency without understanding the exact planned artifact.

Personal-project ownership authority is recorded and MIT remains selected. Public
destination/security channel, exact staged tree and push approval remain operational
owner gates. This preparation does not authorize a push or clear vendor services.

## Sources and limits

- `[user provided / local package]` LICENSE and exact installed license files;
  their relative paths and hashes are in the inventory. Full private review copies
  are retained outside the public candidate.
- `[web search — verify]` [Anthropic legal and compliance](https://code.claude.com/docs/en/legal-and-compliance), checked 2026-10-10: use is subject to provider agreements; SDK authentication and product distribution have separate conditions. This replaces inherited claims that all subscription integrations are automatically sanctioned.
- `[web search — verify]` [Original Draftspect repository](https://github.com/LeonardHope/Draftspect-Add-Ins-for-Word-and-Excel-Powered-by-Claude-Code).
- No jurisdiction-specific enforceability opinion, complete vendor-binary audit,
  vulnerability scan, or company-policy clearance is claimed. Missing text is
  unresolved, never presumed permissive.

The detailed preview.3 decision also reflects Anthropic's October 7 subscription
update and its separate legal integration/authentication restrictions; the legacy
SDK is neither relabelled open source nor categorically banned for personal use.
