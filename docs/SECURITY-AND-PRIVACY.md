# Security and privacy boundaries

This is a single-user local development preview, not a hardened multi-user
service or an OS sandbox. Document instructions are untrusted content.

## Default access

`DRAFTSPECT_AGENT_ACCESS=office-only` is the default. SDK sessions receive
`tools: []`, empty allowed builtins/hooks/agents, `settingSources: []` and strict
MCP configuration. Only the in-process `office` MCP server is passed. The
permission callback denies non-Office tool names. Word mutation tools retain
their own approval and live-precondition checks; not every Excel operation has
identical Word approval UX. Review results before use.

The daemon returns before reading user MCP configuration in this mode. Context
file paths and general shell/filesystem work are unavailable to the model;
perform broader context work only after deliberate opt-in. CLI authentication,
provider-managed policy and SDK internals remain outside an OS-level isolation
claim. No promise of zero local reads or zero provider telemetry is made.

## Explicit extended mode

Only for trusted local work, after reviewing your own settings:

```sh
export DRAFTSPECT_AGENT_ACCESS=extended
# Independently opt in to user settings/hooks:
export DRAFTSPECT_CODEBUDDY_SETTING_SOURCES=user
# Independently opt in to forwarding ~/.claude.json MCP servers:
export DRAFTSPECT_FORWARD_USER_MCP=1
```

Extended mode allows general agent capabilities subject to the retained permission
guard. Each source is off unless explicitly enabled; setting an old opt-in alone
does not bypass office-only mode. Extended native/provider execution was not
accepted in this preparation run. The thin launcher intentionally does not carry
these extension flags: use a reviewed terminal launch if extended mode is needed.

## Where data goes

- Office.js reads the open document when a tool is called. Its bridge is loopback
  only, with Host/Origin checks and a per-process bridge token.
- Model requests send prompts and relevant tool/document output to the selected
  provider under the user's own account. Office.js loads from Microsoft's CDN.
- SDK history, application logs and user-requested archives can contain plaintext
  document/tool content. Protect them as you protect the document itself.
- New lifecycle diagnostic fields are bounded metadata; legacy daemon logs and
  macOS crash reports are **not** guaranteed to be free of sensitive information.
- The application does not automatically upload its crash diagnostics. This does
  not describe all behavior of independent SDK/CLI services.

## Safe recovery

Disconnect means unavailable, not confirmed crash. Main-process watchdog warnings
are separate from renderer/daemon faults. Recovery budgets are bounded. Manual
Stop/Quit cancels pending recovery. Neither reconnect nor recovery replays model
requests or Office writes. Unknown writes stay unknown until live readback.

CLI resume is serial: release the add-in session, wait for its actual child exit,
then use the generated command. Default preview commands disable generic tools,
user settings and external MCP. Exit CLI before returning to Office. Old tool
results are archived snapshots, not current document state.

Return checks the transcript's release fingerprint (file identity, size/times and
hash). Changes require a second explicit confirmation of a still-stable new
fingerprint; continuing writes refuse again. This detects observed changes, not
external CLI ownership/closure, and is not a cross-process lock. Never run Office
and CLI concurrently. See the [recovery guide](RECOVERY.md).

## Reporting

Share versions, error codes, a minimal synthetic reproduction and a redacted
diagnostic summary. Never upload real documents, status tokens, session JSONL,
provider credentials, raw crash dumps or your state directory to public issues.
See [SECURITY.md](../SECURITY.md).
