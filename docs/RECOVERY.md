# Recovery guide

1. Save work in Office. A service failure does not authorize discarding unsaved edits.
2. Read the sidebar/status error. Connection unavailable, model timeout, daemon
   exit, renderer failure and main-process unresponsiveness are separate states.
3. For an uncertain edit, read the affected range in Office and compare with the
   proposal before making a new request. Never re-Apply automatically.
4. Use Stop/Quit, then reopen your own launcher once. If it reports a port conflict,
   identify the owner; do not kill every Node/Electron process.
5. For TLS expiry or account environment errors, correct only this preview's own
   configuration. Trust renewal is explicit. Never disable TLS or OS protections.
6. Preserve a local diagnostic copy with version, stage, exit code/signal and time.
   Redact paths, tokens and content before sharing. `status.json` may contain a
   token; do not attach it raw.

The thin launcher supervises its owned Electron child and can report failures
before application JavaScript runs. It does not automatically restart the main
process. Renderer/daemon recovery is bounded and respects Stop/Quit. Main heartbeat
warnings indicate possible unresponsiveness, not a proven crash. The original
native registration SIGABRT trigger remains unconfirmed.

Revert by normally quitting the preview and returning to an independently retained
working installation. Do not copy preview state into a legacy installation. This
source preparation did not update the installed app or the desktop alias.

## Returning after serial CLI continuation

Release/archive from Setup and wait for the add-in's owned child to exit before
using the generated CLI command. Exit the CLI yourself before clicking return.
If `TRANSCRIPT_CHANGED` appears, the session remains released and composing stays
blocked: the CLI may have legitimately continued the history, or a writer may
still be active. Inspect the changed history, then explicitly click the second
confirmation button. If the file changes again, return is refused again with a
new confirmation. Missing/unreadable/unstable files remain errors, not success.

The fingerprint includes file identity, size/times and a content hash; it is a
stable read observation, **not** proof of external CLI closure or an exclusive
lock. Do not resume the CLI concurrently with Office, including after return.
Successful return reloads historical chat only; it never sends model requests or
applies archived tool calls. Reread the live document before further analysis or
editing. Archived text, paragraph IDs and approvals are not current authority.

## Native Word formatting errors

`RESULT_UNKNOWN: GeneralException` can occur when a host paragraph reference
becomes invalid. The native follow-up fixed one confirmed path: an inner
formatting check reloaded the collection after target references were captured.
The check now reuses that collection while rereading live target OOXML. This
does not diagnose every `GeneralException`. Save your document, inspect the
affected text and formatting, and compare them with the proposal before making
a new request. Do not automatically repeat a failed formatting write.
