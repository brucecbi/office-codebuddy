# Compatibility evidence

Preparation date: 2026-10-10; independent-review repair/native follow-up: 2026-10-11.
Target: macOS source developer preview.

| Component                        | Recorded version / limit                                              |
| -------------------------------- | --------------------------------------------------------------------- |
| Local machine                    | macOS 26.6.2 (25G83), arm64                                           |
| Word / Excel installed           | 16.113.3 / 16.113.3                                                   |
| Node used for preview.2 checks   | 22.23.1 (preview.1 used 25.8.0)                                       |
| npm                              | 10.9.8 (preview.1 used 11.11.0)                                       |
| Electron locked                  | 42.11.3                                                               |
| CodeBuddy SDK locked             | 0.3.243                                                               |
| CLI used for native continuation | 2.162.0; existing authenticated account, not clean-account onboarding |
| Legacy Claude SDK locked         | 0.2.141; not the supported preview route                              |
| ws / zod / Prettier              | 8.21.3 / 4.6.1 / 3.8.3                                                |
| Node dependency minimum          | >=22.12.0; minimum is not native acceptance evidence                  |
| OpenSSL command in repair tests  | System `/usr/bin/openssl`: LibreSSL 3.3.6; no OpenSSL 3 in test PATH  |

OpenSSL 3.x is not a test prerequisite. Earlier preview receipts did not separately
accept the factory LibreSSL route; the repair revision now tests it on this Mac
with Node 22.23.1 plus system-only command bins. The transport-configuration and
desktop-preflight suites finish in seconds, with stdin explicitly closed,
20-second per-command limits, `req -new` for existing keys and key mode 0600.
The full offline suite uses the same PATH. This does not accept every LibreSSL/
OpenSSL version, another machine or native TLS onboarding. Other combinations
remain NOT_RUN; see the dated [test results](../release/TEST-SUMMARY.md).

The October 11 native follow-up accepts the repaired Word formatting path,
Excel bounded old-value/formula previews and stale refusal, plus an explicitly
authorized synthetic Word → CLI → Word continuation on this Mac. Node 22.23.1
was used for the full suite and actual provider processes; the focused failing/
passing reference regression used Node 25.8.0. Locked dependencies were reused,
not installed afresh. Test-only ports/GUIDs, synthetic bindings and a private
process-exit observer isolate the run. The actual tool/approval/handoff code is
the repaired candidate; private instrumentation is excluded from the archive.
Existing trusted TLS and an existing account were used. Fresh TLS/account
onboarding, another Mac and the updated launcher's native GUI were not rerun.
Preserved preview.2 native results below remain historical.

The preview.2 follow-up used freshly installed dependencies, a newly downloaded
Electron binary, isolated state/ports and unique synthetic Office QA registrations
on this same Mac. Its Electron main process, secured status window, real daemon,
and selected native Office scenarios are recorded in [test results](../release/TEST-SUMMARY.md).
The Office QA harness bypasses model generation and overlays only test addresses,
GUIDs and synthetic binding. It is not a complete provider-backed installation
test, acceptance on another Mac, or proof of supported side-by-side installation.
Preview.2's account/model generation was NOT_RUN; the later October 11 follow-up
tested only the separately authorized synthetic serial path above. The exact
Node minimum 22.12.0 and native sleep/wake or soak testing remain NOT_RUN.
Baseline native passes are not silently transferred to the candidate.

Windows platform paths remain in the source for future work. The preview's public
installer refuses Windows to avoid suggesting tested support. Electron itself
supports multiple platforms; that does not establish Office/runtime support for
this project. [Electron platform documentation](https://www.electronjs.org/docs/latest/).

Intel Mac, another account/machine, Office web, Linux, a lowest supported Office
version, native sleep/wake stress, long-running soak, and signed/notarized binary
distribution have not been accepted. Office API availability varies by build and
feature. `.doc` compatibility is not separately accepted; use synthetic `.docx`
files for the preview.
