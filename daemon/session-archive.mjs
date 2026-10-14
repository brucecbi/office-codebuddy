import { mkdir, readFile, writeFile, lstat, readdir } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { join, dirname, resolve } from "node:path";
import { stateDirectory } from "./state-paths.mjs";
import { agentAccess, codebuddyEnvironment } from "./preview-policy.mjs";
import { locateSessionFile } from "./transcript.mjs";
import {
  acceptanceProfile,
  acceptanceCliArguments,
  assertAcceptanceWorkspace,
} from "./acceptance-profile.mjs";

export const shellQuote = (value) => "'" + String(value).replaceAll("'", "'\\''") + "'";
const unchanged = (a, b) =>
  a.size === b.size && a.mtimeMs === b.mtimeMs && a.ctimeMs === b.ctimeMs && a.ino === b.ino;
async function readStableFile(path) {
  const before = await lstat(path);
  if (!before.isFile() || before.isSymbolicLink())
    throw new Error("Archive refuses non-regular files: " + path);
  if (before.size > 128 * 1024 * 1024)
    throw new Error("Archive file exceeds the 128 MiB safety limit: " + path);
  const bytes = await readFile(path);
  const after = await lstat(path);
  if (!unchanged(before, after))
    throw new Error("Session files are still changing. Close their writer and retry archiving.");
  return { bytes, before };
}
async function inventory(path, relative = "", found = []) {
  const meta = await lstat(path);
  if (meta.isSymbolicLink() || (!meta.isFile() && !meta.isDirectory()))
    throw new Error("Archive refuses links or special files: " + path);
  if (!relative && !meta.isDirectory())
    throw new Error("Session attachment root is not a directory: " + path);
  found.push({ path, relative, meta });
  if (
    found.length > 5000 ||
    found.reduce((n, x) => n + (x.meta.isFile() ? x.meta.size : 0), 0) > 256 * 1024 * 1024
  )
    throw new Error("Session attachments exceed the 5000-entry / 256 MiB archive safety limit");
  if (meta.isDirectory())
    for (const name of (await readdir(path)).sort())
      await inventory(join(path, name), join(relative, name), found);
  return found;
}
export function resumeCommands({ provider, session_id, cwd }, { env = process.env } = {}) {
  if (provider !== "codebuddy") return null;
  if (!/^[a-zA-Z0-9][a-zA-Z0-9_:-]*$/.test(session_id || "")) throw new Error("Invalid session ID");
  assertAcceptanceWorkspace(cwd);
  const reminder =
    "This is a serial continuation of a Word add-in session. Prior Word tool results are historical snapshots. The add-in Office connection and approvals are not available here. Use this context for discussion/research; do not modify the previously open Word file via filesystem tools. Return to the add-in and reread the live document for Word edits.";
  const config = env.CODEBUDDY_CONFIG_DIR?.trim();
  let command = config
    ? `CODEBUDDY_CONFIG_DIR=${shellQuote(resolve(config))} codebuddy`
    : "codebuddy";
  if (env.CODEBUDDY_INTERNET_ENVIRONMENT)
    command = `CODEBUDDY_INTERNET_ENVIRONMENT=${shellQuote(codebuddyEnvironment(env))} ` + command;
  const previewArgs =
    agentAccess(env) === "office-only"
      ? [
          "--strict-mcp-config",
          "--mcp-config",
          '{"mcpServers":{}}',
          "--tools",
          "",
          "--setting-sources",
          "none",
        ]
      : [];
  const scoped =
    " " + (acceptanceProfile ? acceptanceCliArguments() : previewArgs).map(shellQuote).join(" ");
  const base = `cd ${shellQuote(cwd)} && ${command} --resume ${shellQuote(session_id)}${scoped}`;
  return {
    resume: base + " --append-system-prompt " + shellQuote(reminder),
    reuse_as_new_session: base + " --fork-session --append-system-prompt " + shellQuote(reminder),
  };
}

// Called only after the user explicitly releases the pane's session. It copies
// original transcript bytes; it does not rewrite the CLI's history or settings.
export async function archiveSession(
  info,
  { root = join(stateDirectory(), "archives"), locate = locateSessionFile } = {},
) {
  const { session_id, provider, cwd, host } = info;
  const source = await locate(session_id, { provider });
  if (!source)
    throw new Error(
      "No persisted transcript found. No archive was claimed; finish one conversation first.",
    );
  const { bytes, before } = await readStableFile(source);
  const artifactRoot = join(dirname(source), session_id);
  let artifacts;
  try {
    artifacts = await inventory(artifactRoot);
  } catch (error) {
    if (error.code !== "ENOENT" || error.path !== artifactRoot) throw error;
    artifacts = [];
  }
  const created = new Date().toISOString();
  const folder = join(root, created.replaceAll(":", "-") + "-" + randomUUID());
  await mkdir(folder, { recursive: true, mode: 0o700 });
  const manifest = {
    schema: 2,
    session_id,
    provider,
    cwd,
    host,
    created_at: created,
    transcript_sha256: createHash("sha256").update(bytes).digest("hex"),
    transcript_bytes: bytes.length,
    source_transcript: source,
    source_session_files: artifacts.length ? artifactRoot : null,
    session_files: [],
    commands: resumeCommands(info),
    scope:
      "Point-in-time conversation archive; not a Word document backup or a live context mirror.",
    privacy:
      "Contains local plaintext conversation/tool content. Protect it like the source document.",
    restore:
      "CLI resume uses its original history store. This archive remains evidence if that store is cleaned; automatic import into the CLI store is not performed.",
    attachment_scope:
      "Copies regular files under the session-ID sibling directory, including tool outputs/subagent logs when present. External workspace files, Word documents and files elsewhere are not copied; references to them may later expire.",
  };
  await writeFile(join(folder, "session.jsonl"), bytes, { flag: "wx", mode: 0o600 });
  for (const item of artifacts) {
    const dest = join(folder, "session-files", item.relative);
    if (item.meta.isDirectory()) await mkdir(dest, { mode: 0o700 });
    else {
      const file = await readStableFile(item.path);
      if (!unchanged(item.meta, file.before))
        throw new Error("Session attachments changed during archive preparation");
      await writeFile(dest, file.bytes, { flag: "wx", mode: 0o600 });
      manifest.session_files.push({
        path: join("session-files", item.relative),
        bytes: file.bytes.length,
        sha256: createHash("sha256").update(file.bytes).digest("hex"),
      });
    }
  }
  await writeFile(
    join(folder, "README.md"),
    `# Word session archive\n\nSession: ${session_id}\n\n${manifest.scope}\n\n${manifest.privacy}\n\n${manifest.attachment_scope}\n\n${manifest.restore}\n\n` +
      (manifest.commands
        ? `## Continue the same session (after closing the add-in session)\n\n\`\`\`sh\n${manifest.commands.resume}\n\`\`\`\n\n## Reuse context in a separate session\n\n\`\`\`sh\n${manifest.commands.reuse_as_new_session}\n\`\`\`\n\n`
        : "") +
      "Return to Word only after exiting the CLI conversation. Reopen the add-in, select Return to Word, and reread the current document. Saved tool outputs do not track later manual edits.\n",
    { flag: "wx", mode: 0o600 },
  );
  // Manifest is written last: its presence marks a fully written package.
  if (!unchanged(before, await lstat(source)))
    throw new Error("Transcript changed during archive creation; package is incomplete");
  for (const item of artifacts)
    if (!unchanged(item.meta, await lstat(item.path)))
      throw new Error("Session attachments changed during archive creation; package is incomplete");
  await writeFile(join(folder, "manifest.json"), JSON.stringify(manifest, null, 2) + "\n", {
    flag: "wx",
    mode: 0o600,
  });
  return { ...manifest, archive_path: folder };
}
