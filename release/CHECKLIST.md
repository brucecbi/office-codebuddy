# Release checklist — 0.2.0-preview.4

**Status: local source candidate; NOT PUBLISHED. No remote or CI configured.**

Read [license scope decision](LICENSE-DECISION.md). Source publication, provider
use and a vendor-containing App are separate decisions.

October 12 is a separate test-isolation candidate from the October 11 native-repair
ZIP. The owner authorized a new source ZIP after inherited-state/concurrent-port
checks and the required offline/format/release gates. Older ZIPs remain unchanged.
Prior same-Mac Word/Excel and serial CLI results are historical; this follow-up
does not rerun them or accept another-machine installation. The dated license
scope memo remains prior evidence, without a new vendor-payload clearance claim.

The local Git preparation is a documentation/package-metadata derivative of that
ZIP, with its own manifest/tree identity. The vendor-contact privacy question is
resolved. The owner selected Office CodeBuddy, brucecbi/office-codebuddy and GitHub
Private Vulnerability Reporting. The repository/reporting feature remain pending
creation and enablement. On October 14 the owner approved the complete public
commit identity using their existing email; GitHub no-reply is optional.

## Source-only license and privacy preparation

- [x] Original repository/history, previous candidates and installed plugin preserved.
- [x] Source-only allowlist; private development documents and screenshots excluded.
- [x] Upstream MIT license and origin retained.
- [x] Owner confirms personal-project public rights and no extra company OSS policy.
- [x] Selected repository license remains MIT; no vendor package is relicensed.
- [x] Authorized icons and synthetic fixtures are allowlisted; no vendor/runtime
      payload, App, account, credential, node_modules or private Git history shipped.
- [x] Exact-version evidence and source-versus-bundle decisions recorded.
- [x] Source-facing notice/authentication/install-only claims corrected.
- [x] Vendor-email privacy question independently closed: the pinned SDK README
      and public service agreement publish the contact. See PRIVACY-REVIEW.md.
- [x] Raw privacy reports and original development history excluded from this tree;
      upstream attribution remains intact and review blind spots are disclosed.
- [x] No unresolved copyright-license blocker identified for this exact source scope;
      this is bounded review, not general legal clearance or publication approval.

## Current October 12 engineering preparation

- [x] Test-only session state and dynamic-port isolation repairs; inherited state
      and two concurrent test workers checked, including intentional port conflict.
- [x] Full offline npm suite: 526/526, zero failures/skips; current formatting checked.
- [x] Final source allowlist reviewed, with unchanged production implementation
      and dependency versions. Archive integrity and protected-file closure are
      separate gates in the accompanying receipt; do not hand off without them.

## Preserved engineering evidence (October 10–11; not new preview.4 native runs)

- [x] Independent preview state; explicit account environment; safe extension defaults.
- [x] Installation refuses conflicting registrations; no automatic old-session import.
- [x] macOS/Windows/source/App support statements separated.
- [x] Dependency text evidence and unresolved redistribution questions recorded.
- [x] Fresh `npm ci` and explicit checksum-verified Electron download on the same Mac,
      with task-local caches and installation-script review; Node 22.23.1 tested.
- [x] Historical preview.2 same-Mac Electron and model-free Word/Excel native checks;
      overlays, manual menu activation and untested full-provider routes disclosed.
- [x] Final candidate offline tests, content/hash review and source artifact verification — see TEST-SUMMARY.md.
- [x] Historical October 10 preview.3 packaging fix retains project notices in thin-wrapper Resources
      with hashes; no application runtime/security/recovery behavior changed.
- [x] Current offline and notice-build checks are recorded separately from preview.2
      fresh/native evidence. This follow-up reuses those pinned dependencies.
- [x] Review-repair code and permission/handshake limits documented; current exact
      tests and manifest evidence recorded separately in TEST-SUMMARY.md.
- [x] Owner authorized repackaging only after native acceptance; build a new named
      source archive from the reviewed allowlist, preserve the old package and
      bind hashes/test evidence in a new separate receipt. See TEST-SUMMARY.md.
- [x] Historical October 11 native Word formatting/Excel preview and authorized
      synthetic serial CLI checks for the runtime repairs; native and simulated
      checks remain distinguished. No installed-App update or push.

## Before a public push (not authorized by this task)

- [ ] Approve exact candidate hashes and the source-only scope. Ownership confirmation
      and selected MIT license are already recorded; no vendor grant is inferred.
- [x] Select GitHub owner/name and private security-reporting channel: repository
      brucecbi/office-codebuddy, with GitHub Private Vulnerability Reporting.
      README/package/security metadata use this selected, unpublished destination.
- [ ] Establish the repository and enable/verify its private reporting form when
      separately authorized; do not describe the planned channel as already active.
- [x] Initialize a new Git repository from **only** this reviewed source candidate.
      Do not import private Git history, parent evidence directory or node_modules.
- [x] Re-run release check and inspect the exact staged tree after documentation
      closure; keep sensitive-data evidence bound to that tree, not only the old ZIP.
- [x] Confirm public author and committer display name: Bruce Bi.
- [x] Receive and verify the complete owner-approved public email for both author
      and committer in local Git configuration; no unapproved identity is inherited.
      The owner explicitly chose their existing email instead of GitHub no-reply.
- [ ] Create and inspect the first commit only after identity confirmation. A staged
      tree has no commit author metadata and does not complete this gate by itself.
- [x] Retain Developer Preview/NOT_RUN labels and the untested clean-machine boundary.
- [ ] Explicitly authorize public push. No automatic release/npm publish/CI enablement.

Another-machine validation is required before promising supported/reproducible
installation; it is not a copyright prerequisite for a candid source-review
preview. Preview.2 used existing trusted TLS and QA GUID/port overlays;
clean-account/provider/TLS onboarding remains NOT_RUN.

## Before a vendor-containing App/DMG (not cleared)

- [ ] Define and hash the exact final payload; distinguish runtime, installation-only
      helpers and external user-supplied dependencies. Do not copy all node_modules.
- [ ] If SDK/CLI is copied, settle CodeBuddy's whole-package grant/required notices
      and applicable service terms; retain actual vendor notices without rewriting.
- [ ] If native tools are copied, cover actual ripgrep/native-module/toybox/zsh
      provenance and obligations. Assess PCRE2's downstream binary exception;
      a generic zsh GPL warning is not a finding that Draftspect is GPL.
- [ ] If Anthropic CLI/SDK is copied/enabled, review conditional distribution,
      unmodified binary/authentication, own-user billing and branding; no proxy/resale.
- [ ] If Electron/Chromium or other libraries are copied, preserve the exact runtime
      license/NOTICE set and audit compiled third-party components.
- [ ] If pinned extract-zip is copied, resolve its exact-version text/compiled notices.
      If absent from the final artifact, mark N/A; it does not block a runtime-only
      App or source publication merely because it is an install dependency.
- [ ] Self-contained resources, no maintainer paths, own-machine onboarding and TLS.
- [ ] Final native Word/Excel/SDK acceptance on a separate clean Mac/account.
- [ ] Developer ID signing, notarization and downloaded-artifact installation checks.
- [ ] Safe upgrade/uninstall, release owner, support policy and redacted diagnostics.

The current local thin wrapper includes only its executable/icon/config/notices,
uses external paths and is not offered as a portable App. Its project-license
notice gap is repaired; complete App binary clearance is not claimed. Windows and
Intel remain unaccepted; no cross-machine or provider entitlement claim is made.
