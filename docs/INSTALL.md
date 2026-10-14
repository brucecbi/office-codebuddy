# Installing the source preview on macOS

This is a developer procedure, not a one-click installer. Evaluate with synthetic
documents on a separate Mac or user account. Do not replace a working installation
as part of trying this preview. macOS arm64 is the target; Windows installer is
deliberately disabled. See [compatibility](COMPATIBILITY.md).

## 1. Dependencies and account

Prepare Node >=22.12.0 and npm, desktop Word/Excel, and your own CodeBuddy CLI from
the provider's official distribution. Review the provider's account/usage terms.
Neither Office nor a model subscription is supplied. Do not paste tokens into the
repository, issue reports, launcher configuration or shell history.

From this source directory:

```sh
npm ci
node node_modules/electron/install.js
npm test
```

`npm ci` uses the lockfile and may run package installation scripts. The pinned
Electron 42.11.3 package downloads its binary lazily on first Electron use, not
during `npm ci`. The explicit second command invokes its bundled installer and
verifies the download against the package's checksums. Run it before building the
Swift launcher, which requires the binary to exist and does not download it.
The source archive excludes node_modules and runtime executables;
it includes the project's icons and synthetic Office fixture files.
It still references the legacy Claude SDK dependency for retained adapter tests;
that SDK is not an MIT-licensed part of this project. See the dependency review.

Installation/download and publication are separate scopes. This repository's MIT
license does not grant a model account, service access or rights to redistribute
provider packages. Users obtain dependencies under their own applicable terms.
CodeBuddy's SDK documentation explicitly contemplates app integration; the exact
package's whole-bundle license scope still needs clarification before shipping it
inside an App. The public service-agreement sample is not a substitute for the
agreement for your account/region. See [license decision](../release/LICENSE-DECISION.md).

For an installation audit, `npm ci --ignore-scripts` may be used first to inspect
the exact packages, followed by a normal `npm ci` and the explicit Electron
installer. Use a private task-local npm/Electron cache if testing isolation. Do
not copy dependencies from an existing installation and call that a fresh install.
See [Electron's installation instructions](https://www.electronjs.org/docs/latest/tutorial/installation).

Choose the environment matching your account; no environment is assumed:

```sh
# Example only: use external only if it matches your own account.
export CODEBUDDY_INTERNET_ENVIRONMENT=external
```

The pinned SDK declares `external`, `internal`, `ioa`, `cloudhosted`. They are
account environments, not interchangeable models. Official documentation describes
`internal` for the China edition and `ioa` for the enterprise edition; check your
account rather than copying the example blindly.
[CodeBuddy environment reference](https://www.codebuddy.ai/docs/cli/env-vars).
Sign in through CodeBuddy's own normal flow in the selected environment.

CodeBuddy executable discovery: explicit absolute `CODEBUDDY_CODE_PATH`, then PATH.
The desktop preflight requires an installed executable and does not silently
substitute a missing CLI. The daemon-only SDK route may otherwise use its bundled
CLI; prefer the desktop route and an explicit CLI installation.

## 2. Local TLS, without changing existing trust

The preview uses `~/.draftspect-preview/` by default. It never imports the original
installation's transport or sessions. Keep this directory private and outside
synced/shared folders. Do not copy another person's keys or account files.

Create a new per-user certificate/configuration explicitly:

```sh
node scripts/setup-preview-tls.mjs --create
node scripts/local-transport-status.mjs --require-https
```

The setup script requires OpenSSL with `req -addext`, refuses existing transport
configuration, creates a 90-day self-signed **non-CA** server leaf for
`127.0.0.1` only, and stores its key with mode 0600. It does not install trust,
register Office, run a model, or start listeners. It prints the certificate path
and SHA-256 fingerprint, never private key contents.

Inspect the certificate before trusting it. If authorized by your local policy,
use Keychain Access or the macOS `security add-trusted-cert` flow to trust only
this server certificate for SSL at `127.0.0.1` in your **login** keychain. This is
a separate explicit system change requiring local approval. Do not use a broad
root CA, disable TLS verification, or change system-wide trust. Native Office
loading must be checked after trust; the status script reports `NOT_CHECKED` for
native trust. Renew before expiry using a new pair and a reviewed config switch;
the setup script intentionally does not overwrite an existing configuration.

## 3. Start, then install deliberately

```sh
npm start
```

Use the tray menu to show status. Alternatively build the optional thin launcher
to a **new** path, then open it yourself:

```sh
npm run build:macos -- --output /absolute/existing-folder/Draftspect-Preview.app
```

The native Swift launcher needs Apple's local build tools and captures the
explicit account environment plus local executable/project paths. It is valid
only on the machine where built. Do not upload it as a portable release App.
The build preserves project LICENSE, UPSTREAM.md and THIRD-PARTY-NOTICES.md in its
Resources directory, with hashes in build-info.json. It contains no external SDK,
Electron, Node or CLI payload; the configured paths refer to separate installations.
There is no auto-update, login item, certificate installer or model auto-prompt.

The legacy Claude route is not accepted as a supported preview provider. As checked
on 2026-10-10, Anthropic's October 7 help update permits SDK/subscription usage;
its legal page separately restricts product authentication, credential
intermediation and conditional distribution of the unmodified CLI. Do not infer
that all personal subscription use is forbidden or that an App may freely proxy
subscribers' credentials. See the dated primary references in the license decision.

From the tray's explicit install command, register the preview. Installation
refuses conflicting Draftspect manifests instead of overwriting them. The preview
retains the Office GUIDs used by its verified bootstrap fixtures and shares ports
47833/47834 with the development line: **side-by-side use is not supported**.
Port conflicts fail; no other process is stopped. Do not clear Office caches or
manually remove unrelated manifests to force installation.

In a new synthetic Word document / Excel workbook, use Home → Add-ins → Developer
Add-ins → Draftspect. On the prior baseline, ordinary menu activation worked after
an Office cold restart without initialization. If necessary, the explicit
initialization action opens a fresh synthetic copy; it never binds a business file.
The candidate's isolated native QA results are in [test results](../release/TEST-SUMMARY.md).
They do not replace acceptance on your machine. If an already-running Office host
does not show a newly added registration, save your work and fully quit/reopen that
host once before diagnosing a missing manifest. Do not repeatedly install, clear
caches, or remove another plugin. The QA follow-up needed this for Excel's new ID.

## 4. First functional check

Enter a short invented paragraph or two cells. Wait for a connected sidebar and
an account-specific model catalogue; choose a model explicitly. Only then, if
authorized to send the synthetic content to the provider, request a read-only
summary. This request may be chargeable. Try a small edit, decline the approval
first, then approve a fresh proposal and independently verify its result.

## Stop, recovery, and removal

Close hides the status window; **Quit Draftspect** exits its owned service. Use
Stop/Quit before switching installations. Reconnection never resends an old
model request or document write. `RESULT_UNKNOWN` requires explicit live reread
and verification, not a second Apply.

The tray's uninstall action removes only exact matching preview manifest files;
modified/conflicting files are retained for manual review. Never bulk-delete the
Office wef folder. CLI accounts, original plugins, TLS trust and archived sessions
are not deleted. See [recovery](RECOVERY.md).
