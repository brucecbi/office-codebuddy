import { constants } from "node:fs";
import { open, lstat } from "node:fs/promises";

// Check the opened descriptor before truncating: an attacker-supplied symlink
// must not redirect or truncate another file. This guards the final component.
export async function writeRegularFile(path, contents, { mode = 0o666 } = {}) {
  try {
    const existing = await lstat(path);
    if (!existing.isFile() || existing.isSymbolicLink())
      throw new Error("Refusing symlink or non-regular write target");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  const file = await open(
    path,
    constants.O_WRONLY | constants.O_CREAT | constants.O_NOFOLLOW | constants.O_NONBLOCK,
    mode,
  );
  try {
    if (!(await file.stat()).isFile()) throw new Error("Refusing non-regular write target");
    if (mode === 0o600) await file.chmod(mode);
    await file.truncate(0);
    await file.writeFile(contents, "utf8");
  } finally {
    await file.close();
  }
}
