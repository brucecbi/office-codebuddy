# Contributing

Office CodeBuddy's selected repository is
[brucecbi/office-codebuddy](https://github.com/brucecbi/office-codebuddy).
Use its issues for non-sensitive discussion after publication. Follow
[SECURITY.md](SECURITY.md) for private vulnerability reporting.

Please discuss substantial changes before implementing them. Work on a feature
branch; keep patches scoped, preserve upstream attribution, and confirm you have
the right to contribute under this repository's MIT license.

Use synthetic fixtures only. Run `npm test`, `npm run format:check` and
`npm run check:release`. Update the changelog, support matrix and test summary;
distinguish mock, real SDK argument serialization, native Electron, native Office
and model-generation tests. Never mark unrun layers passed.

Do not add credentials, private paths, accounts, raw session data, business
screenshots, vendor binaries or copied third-party source without review. Do not
replace write approvals or uncertain-result checks with prompt-only assurances.

There is no configured remote or CI in this candidate. Publication and CI setup
require separate owner approval; tests can be run locally in the meantime.
