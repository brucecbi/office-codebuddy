# Fresh-install license follow-up — 2026-10-10

Scope: 0.2.0-preview.2 source-only artifact. This is a factual inventory and release
gate assessment, not a jurisdiction-specific legal opinion or a vendor-binary
clearance. Actual installed notices were read after a fresh locked installation;
hashes appear in dependency-inventory.json. Provider terms may change.

This section preserves the preview.2 fresh-install evidence. The current
[preview.3 scope decision](LICENSE-DECISION.md) supersedes broad release-gate
interpretations: a missing notice is relevant to the component actually copied,
not automatically to a source reference or thin wrapper without vendor payloads.

## Confirmed and unresolved

- The root project MIT notice and original attribution remain unchanged. The
  owner confirms personal-project public rights and no extra company approval.
  This does not establish permission to apply MIT to third-party SDKs or brands.
- CodeBuddy SDK 0.3.243: npm/package metadata says MIT. No top-level LICENSE is
  included. `cli/LICENSE.txt` is an MIT-form Tencent Cloud grant. The scope over
  all SDK files is still unclear; obtain provider clarification before copying
  these files into a distributable App. No such code is copied into this ZIP.
- The exact CodeBuddy zsh vendor directory contains only `LICENSE` and `bin/zsh`.
  No separate shell functions are shipped in that directory. The license warns
  that some functions are GPL-licensed; that warning alone does not prove this
  artifact contains them or make all project code GPL. The compiled binary's
  source/provenance and obligations have not been exhaustively audited.
- `@electron-internal/extract-zip@1.0.5`: npm metadata identifies gitHead
  `b83e459fd04c53b0a1c8438a6792df8f64be47fc`. Its pinned package.json/README say
  BSD-2-Clause, but the installed package and pinned tree lack a root LICENSE.
  Current main's later MIT text is not exact-version evidence. Do not silently
  substitute it. This is an Electron installation-only module in this project;
  its notice gap is a gate only for an artifact that copies this component.
  Later upstream PR #19 confirms the change, not an explicit retroactive grant.
- Anthropic agent SDK 0.2.141 and the installed darwin-arm64 package contain
  vendor-specific legal-agreement notices, not MIT licenses. The provider's
  current legal page permits some product integrations with an unmodified CLI
  under specified commercial terms; it is incorrect to call all redistribution
  forbidden, or all SDK subscription authentication permitted. Distinguish
  product distribution, each user's own authentication and SDK API/cloud-key
  routes. No third-party credential intermediary is supplied by this project.
- Electron's actual downloaded distribution has its MIT notice and Chromium
  third-party notices. Their presence/hashes are confirmed, but a full future
  App's compiled component obligations are not thereby cleared. Seven optional
  non-current-platform Claude payloads were not installed or text-reviewed.

## Practical release boundary

The planned ZIP ships project source, original project assets, sanitized synthetic
fixtures and package references, not node_modules, SDK/CLI/runtime executables,
accounts, certificates or a prebuilt App. Source-only publication and bundling
vendor executables are different release decisions. Retain this qualification,
require users to obtain dependencies under their own applicable agreements, and
do not claim the entire installed tree is open source or MIT.

Before a vendor-containing binary release, settle the scope/notices of the SDK
and components actually copied, inventory compiled tools, preserve required notices
and assess the actual distribution/authentication arrangement. Excluded installer
helpers are N/A for that bundle. The current thin wrapper copies the project MIT
license and origin/notice files without copying external runtimes. This follow-up
does not authorize publication, hosted operation or a self-contained App release.

## Primary references

- [CodeBuddy SDK npm version](https://www.npmjs.com/package/@tencent-ai/agent-sdk/v/0.3.243) and the installed version's `cli/LICENSE.txt`.
- [extract-zip pinned source](https://github.com/electron/extract-zip/tree/b83e459fd04c53b0a1c8438a6792df8f64be47fc) and [pinned package metadata](https://raw.githubusercontent.com/electron/extract-zip/b83e459fd04c53b0a1c8438a6792df8f64be47fc/package.json).
- [Current extract-zip license](https://github.com/electron/extract-zip/blob/main/LICENSE), for comparison only, not the pinned version's grant.
- [Anthropic legal and compliance](https://code.claude.com/docs/en/legal-and-compliance), checked 2026-10-10.
- Actual fresh Electron `dist/LICENSE` and `dist/LICENSES.chromium.html`; no vendor notice text is replaced with a metadata label.
