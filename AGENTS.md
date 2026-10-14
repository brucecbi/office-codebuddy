# Draftspect Developer Preview

This is a source-only preview, defaulting to CodeBuddy on macOS. Read README.md,
docs/SECURITY-AND-PRIVACY.md and docs/TESTING.md before changing runtime behavior.
Keep LICENSE and UPSTREAM.md. Do not restore private development paths or reports.

- Preserve document write approvals, live preconditions, serial execution and
  RESULT_UNKNOWN. Never replay model requests or Office writes on reconnect.
- Default to Office-only tools; user settings/hooks and extra MCP servers require
  separate explicit opt-ins. Do not assume an account environment or model ID.
- Keep preview state isolated. Do not migrate another installation's history,
  replace its manifests, modify trust, kill unknown processes, or start real models
  as a test. Use temporary synthetic fixtures.
- Tests: `npm test`, `npm run format:check`, `npm run check:release`.
  Native tests are opt-in and do not become native Office evidence automatically.
- Use a feature branch for subsequent work. Publishing, remote changes, CI setup,
  global configuration and dependency upgrades require explicit authorization.
- Update CHANGELOG.md and release/TEST-SUMMARY.md for meaningful changes. Generate
  release/approved-files.json only after final review; the manifest excludes itself.
