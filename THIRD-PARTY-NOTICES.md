# Third-party notices

This source candidate derives from **Draftspect by Leonard Hope**. Its original
MIT copyright and permission notice are reproduced verbatim in [LICENSE](LICENSE).
See [UPSTREAM.md](UPSTREAM.md) for the derivation and modification scope.

The candidate does not contain installed npm dependencies, SDK/CLI vendor payloads,
Electron/Chromium binaries, Office binaries or bundled credentials. Dependencies
listed in package.json/package-lock.json are obtained separately on installation
and retain their own copyrights and terms. Their inventory and unresolved license
questions are in [release/DEPENDENCY-REVIEW.md](release/DEPENDENCY-REVIEW.md).

The project's MIT grant covers this derivative's project code and authorized
assets, not separately installed provider software, accounts or trademarks.
The owner has confirmed authority to open-source the additions/assets as a
personal project. Original upstream attribution remains unchanged.

The optional local thin Swift wrapper copies this file, LICENSE and UPSTREAM.md
verbatim into `Contents/Resources/` and records their hashes in build-info.json.
It references external project/Node/Electron/CLI paths; it does not contain those
products. Its build is machine-specific, not a portable vendor-containing App.
See [the scope decision](release/LICENSE-DECISION.md). This source notice makes
no claim that vendor dependencies' required notices are already complete.

This notice is not a substitute for the full third-party notices that a future
binary bundle would require. Do not ship an App/DMG using this source-only notice
as proof of complete runtime licensing compliance.
