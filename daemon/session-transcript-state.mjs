import { constants } from "node:fs";
import { open, lstat } from "node:fs/promises";
import { createHash } from "node:crypto";
import { locateSessionFile } from "./transcript.mjs";

const metadata = (stat) => ({
  dev: stat.dev,
  ino: stat.ino,
  size: stat.size,
  mtime_ms: stat.mtimeMs,
  ctime_ms: stat.ctimeMs,
});
const sameMetadata = (a, b) => Object.keys(a).every((key) => a[key] === b[key]);

// A stable read observation, not a lock on or proof of closure of an external CLI.
// No transcript content, credentials or complete environment are logged.
export async function captureTranscriptFingerprint(info, { locate = locateSessionFile } = {}) {
  const path = await locate(info.session_id, { provider: info.provider });
  if (!path)
    throw new Error(
      "TRANSCRIPT_UNAVAILABLE: close the CLI and retry release; no transcript found.",
    );
  const file = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK);
  try {
    const before = await file.stat();
    if (!before.isFile() || before.size > 128 * 1024 * 1024)
      throw new Error(
        "TRANSCRIPT_UNAVAILABLE: expected a regular transcript within the 128 MiB limit.",
      );
    const bytes = await file.readFile();
    const after = await file.stat();
    const current = await lstat(path);
    if (
      !current.isFile() ||
      current.isSymbolicLink() ||
      !sameMetadata(metadata(before), metadata(after)) ||
      !sameMetadata(metadata(after), metadata(current))
    )
      throw new Error(
        "TRANSCRIPT_UNSTABLE: the session file is still changing; close its writer and retry.",
      );
    return { path, ...metadata(after), sha256: createHash("sha256").update(bytes).digest("hex") };
  } finally {
    await file.close();
  }
}

export function sameTranscriptFingerprint(a, b) {
  return (
    !!a &&
    !!b &&
    ["path", "dev", "ino", "size", "mtime_ms", "ctime_ms", "sha256"].every(
      (key) => a[key] !== undefined && a[key] === b[key],
    )
  );
}
