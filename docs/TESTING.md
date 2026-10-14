# Testing without disturbing another installation

## Offline suite

```sh
npm test
npm run format:check
npm run check:release
```

The Node suite uses temporary synthetic state, injected Office/SDK behavior, some
real SDK argument serialization and temporary loopback listeners. It does not
send model requests or register Office. A runner that forbids loopback binding
may report `EPERM`; obtain normal runner permission, never disable application
security settings or silently skip the tests. The tests may create/remove their
own temporary files and start/stop their own fixture processes.

Certificate fixtures close subprocess stdin and have 20-second command limits.
They support the tested factory LibreSSL 3.3.6 route; OpenSSL 3 is not mandatory.
After `npm ci`, SDK-dependent tests must run, not skip. Without the CodeBuddy SDK,
the nine SDK/version-adapter tests in `codebuddy-close`/`preview-policy` explicitly
skip; SDK-independent registry/policy tests still run. Installed SDK entry/internal-API drift fails.
This missing-dependency check is not a substitute for the complete installed suite.

The release checker verifies the reviewed relative-path/hash allowlist and a
limited local sensitive-pattern scan. It never uploads source. It excludes a local
node_modules directory and Git metadata from source checks; the **archive** must
be built from the allowlist, not by zipping the whole development directory.
Only the manifest itself is excluded from its hash list to avoid self-reference.

## Session and port isolation regressions

Session tests override HOME, USERPROFILE and DRAFTSPECT_STATE_DIR before each
fresh session-store import, then restore all three even when a callback fails.
An inherited caller state directory must stay untouched. The regression covers
both a previously set and an absent state variable; external acceptance also
runs the file twice against the same synthetic inherited directory.

```sh
node --test tests/sessions.test.mjs tests/graceful-shutdown.test.mjs
```

Shutdown fixtures hold two distinct kernel-assigned loopback ports during
allocation, release them immediately before spawning, and retry at most three
times only on a real address-in-use startup failure. The intentionally conflicting
second child uses the first child's exact pair and must fail without retries.
The parallel-fixture test requires two daemons to remain ready simultaneously on
four distinct bound ports. Separate concurrent test-worker runs cover isolation
between invocations. Cleanup stops only owned children and awaits their exit
before deleting their temporary fixture. Public plugin ports/GUIDs are unchanged.

## Native layers

`npm run build:macos` compiles the thin Swift launcher and runs its config checker.
That is not native Electron/Office runtime acceptance. Scripts named `*-e2e.mjs`
are opt-in; review each one's process, window, state and model scope first.
Some use real provider control connections, and their inclusion is not permission
to run them. The default npm suite excludes them.

Native Office tests require a new synthetic document, isolated state and explicit
installation/trust coordination. Do not displace an active plugin, alter business
documents or clear Office caches to obtain a pass. Real-model tests require separate
data/account authorization. Record PASS, FAIL, BLOCKED and NOT_RUN individually.

## Regression obligations

- Configuration: unknown environment/mode, selected provider dependencies, default
  no-MCP/no-hook policy, explicit extension opt-ins, safe CLI resume arguments.
- Installation: inspect both hosts before writes; refuse legacy or duplicate-ID
  registrations, changed files and symlinks; preserve unrelated registrations.
- Writes: deny stale/ambiguous plans, honor cancellation, execute once, retain
  RESULT_UNKNOWN through disconnect, require live verification before further edits.
- Review repairs: show bounded Excel old values/formulas; retain complete range
  preconditions; reject Word format-only changes both at approval and inner write
  gates, including changes during preparation. Inner Word checks must use the
  already-loaded paragraph collection: refreshing it can invalidate captured
  native proxies. Fresh OOXML checks remain mandatory. Queue eviction must not discard
  persisted write receipts.
- CLI handoff: unchanged history returns normally; external writes and inode
  replacement refuse first return. New explicit confirmation accepts only an
  unchanged challenge snapshot; later changes refuse again. Never replay edits.
- File writes: token/context symlinks and linked immediate token-state directories
  must be rejected without changing their targets. Builder snapshot links must fail.
- Lifecycle: distinguish creation/readiness, renderer/daemon/main faults; bounded
  recovery, Stop/Quit races, actual owned-child exit, no replay.
- Packaging: fixture metadata and part diff, source allowlist, license/provenance,
  no private state/history/binaries, clean install with the pinned lockfile.

For CLI acceptance, do not equate CodeBuddy 2.162's stream init `tools` inventory
with enabled capabilities. The installed checker applies `--tools ""` and the
capability collector excludes disabled built-ins. Verify the version's actual
filter, use a nonempty positive control, and check for executed tool events.
The October 11 follow-up combined those offline installed-code checks with real
synthetic continuation and no CLI tool calls; it did not intercept a model-wire
request. See the vendor's [CLI flags](https://www.codebuddy.ai/docs/cli/cli-reference).

See [release/TEST-SUMMARY.md](../release/TEST-SUMMARY.md) for this candidate's
actual results. Historical baseline evidence never fills an unrun candidate test.
