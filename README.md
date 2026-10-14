# Office CodeBuddy — Developer Preview

Local Office add-in bridge for **Microsoft Word and Excel**, using your own
CodeBuddy account. Version **0.2.0-preview.4**. macOS-first, source-only preview;
not a signed installer or a promise of cross-platform compatibility.

Planned repository: [brucecbi/office-codebuddy](https://github.com/brucecbi/office-codebuddy).
This independent community project is based on Draftspect by Leonard Hope.
The preview retains Draftspect's internal package, launcher and add-in identifiers.

**Candidate status:** prepared locally; not published. Read the
[release checklist](release/CHECKLIST.md) before any public push.

## What it does

- Reads and analyses the active document through Office.js.
- Provides targeted Word text/formatting tools, tracked-change and comment tools,
  guarded edits, confirmation cards and explicit uncertain-result warnings.
- Provides Excel range/formula tools; coverage differs from Word.
- Keeps process detail collapsed and exposes connection/failure status.
- Supports serial session archiving and CLI resume. Office and CLI do not share a
  live document connection. Re-read the document after returning to Office.

No claim of feature parity with Microsoft's built-in agent or lossless handling
of every document. Use new synthetic files and backups for evaluation.

## Support boundary

| Platform / route                                   | Evidence                                                                                                                                                                                                                           |
| -------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| macOS arm64, local Word/Excel and desktop shortcut | Native tests passed on the preceding development baseline on one Mac                                                                                                                                                               |
| This sanitized preview candidate                   | October 12 test-isolation follow-up; production implementation files retain the October 11 native-repair baseline. Current offline results and historical native scope are separated in [candidate tests](release/TEST-SUMMARY.md) |
| Another Mac / Intel Mac                            | Not validated                                                                                                                                                                                                                      |
| Windows                                            | Experimental source paths only; native acceptance not run; preview installer disabled                                                                                                                                              |
| Browser Office / Linux                             | Not supported by this preview                                                                                                                                                                                                      |
| Distributable `.app` / `.dmg`                      | Not provided; the optional thin launcher is built for its own machine                                                                                                                                                              |

## Get started

The repository destination above has been selected; this candidate is still
unpublished. Use this source directory or a reviewed source archive until the
repository is available. Installation and support boundaries below remain the
same for this documentation revision.

1. Read [installation](docs/INSTALL.md) and [privacy/security](docs/SECURITY-AND-PRIVACY.md).
2. Prepare Node **22.12.0 or newer**, current desktop Office, and your own properly
   licensed/authenticated CodeBuddy CLI. The tested toolchain is documented in
   [compatibility](docs/COMPATIBILITY.md); minimum dependency engines are not a
   tested compatibility range.
3. Run `npm ci`, then `node node_modules/electron/install.js` in the candidate
   directory after reviewing dependency notices. These steps download dependencies
   and Electron; this is not an offline app. The explicit Electron step is required
   before building the thin launcher with this pinned version.
4. Explicitly select the account environment, prepare your own loopback TLS, then
   start the local service. Register the preview from its tray menu only on a
   machine without a conflicting Draftspect installation.

The default **office-only** profile does not forward personal MCP servers or load
user settings/hooks. It exposes no generic filesystem/shell tools to the model.
Extended capabilities require explicit opt-ins; see the security guide. CodeBuddy
may still use its own account files, runtime policies and network services.

Document content used in an AI request goes to the selected provider. Office.js
also loads from Microsoft's CDN. Local bridge/launcher does **not** mean local
model inference, free service, offline processing or provider endorsement.

## Development

```sh
npm test
npm run format:check
npm run check:release
```

Optional native scripts are separate and must be reviewed before execution. The
legacy Claude adapter remains for development compatibility, is not the supported
preview route, and is subject to separate provider terms and authentication.

## License and origin

Repository code retains the upstream [MIT license](LICENSE). Original Draftspect
by Leonard Hope; this is a modified derivative, not an official upstream,
Microsoft, Tencent or Anthropic release. See [UPSTREAM.md](UPSTREAM.md) and
[dependency review](release/DEPENDENCY-REVIEW.md). SDK/CLI/runtime dependencies
retain their own terms and are not relicensed by this repository.

The [license decision](release/LICENSE-DECISION.md) separates source publication,
local thin-wrapper builds and vendor-containing App distribution. No unresolved
copyright-license blocker was identified for this exact source-only scope; this
is not publication approval, a provider-service entitlement or binary clearance.
The dated preview.3 memo is retained as prior evidence; this test-only revision
does not expand the redistributed payload or claim a new legal review.
The source ZIP includes project icons and synthetic Office fixtures, but no
vendor code/binaries, npm dependencies or private Git history.

Contributions: [CONTRIBUTING.md](CONTRIBUTING.md). Security: [SECURITY.md](SECURITY.md).
GitHub Private Vulnerability Reporting is the selected security channel and will
be enabled and verified when the repository is established.
