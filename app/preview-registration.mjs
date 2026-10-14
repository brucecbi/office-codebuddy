// Preview-only registration. Inspect BOTH hosts before writing; never overwrite
// legacy manifests, follow destination symlinks, or clean up someone else's files.
import { readFile, readdir, lstat, mkdir, open, unlink } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";
import { renderManifestTransport } from "../daemon/local-transport.mjs";

export const previewManifestName = (host) => `draftspect-preview-${host.toLowerCase()}.xml`;
const hostDir = (home, host) =>
  join(home, "Library", "Containers", `com.microsoft.${host}`, "Data", "Documents", "wef");
const manifestId = (xml) => xml.match(/<Id>\s*([^<]+)\s*<\/Id>/i)?.[1]?.toLowerCase();
const missing = (error) => {
  if (error.code === "ENOENT") return [];
  throw error;
};

export async function planPreviewRegistration({ home = homedir(), projectRoot, transport }) {
  const plan = [];
  for (const host of ["Word", "Excel"]) {
    const text = renderManifestTransport(
      await readFile(join(projectRoot, "manifests", `${host.toLowerCase()}.xml`), "utf8"),
      transport,
    );
    const id = manifestId(text);
    if (!id) throw new Error("PREVIEW_MANIFEST_INVALID");
    const directory = hostDir(home, host);
    const directoryInfo = await lstat(directory).catch((error) => {
      if (error.code === "ENOENT") return null;
      throw error;
    });
    if (directoryInfo && !directoryInfo.isDirectory())
      throw new Error(
        "PREVIEW_REGISTRATION_CONFLICT: registration directory is not a real directory.",
      );
    const file = previewManifestName(host);
    for (const name of await readdir(directory).catch(missing)) {
      if (!name.endsWith(".xml")) continue;
      const path = join(directory, name);
      const info = await lstat(path);
      if (name === file && !info.isFile())
        throw new Error("PREVIEW_REGISTRATION_CONFLICT: destination is not a regular file.");
      if (name === `${host.toLowerCase()}.xml`)
        throw new Error(
          "PREVIEW_REGISTRATION_CONFLICT: an existing installation must be managed separately; no files changed.",
        );
      if (!info.isFile()) continue;
      const existing = await readFile(path, "utf8");
      if ((name === file && existing !== text) || (name !== file && manifestId(existing) === id))
        throw new Error(
          "PREVIEW_REGISTRATION_CONFLICT: an existing registration would conflict; no files changed.",
        );
    }
    plan.push({ host, path: join(directory, file), directory, text });
  }
  return plan;
}

export async function installPreviewRegistration(options) {
  const plan = await planPreviewRegistration(options);
  const created = [];
  try {
    for (const item of plan) {
      await mkdir(item.directory, { recursive: true });
      let handle;
      try {
        handle = await open(item.path, "wx", 0o600);
      } catch (error) {
        if (
          error.code === "EEXIST" &&
          (await lstat(item.path)).isFile() &&
          (await readFile(item.path, "utf8")) === item.text
        )
          continue;
        throw error;
      }
      const identity = await handle.stat();
      created.push({ path: item.path, ino: identity.ino });
      try {
        await handle.writeFile(item.text);
      } finally {
        await handle.close();
      }
    }
  } catch (error) {
    for (const item of created) {
      const current = await lstat(item.path).catch(() => null);
      if (current?.isFile() && current.ino === item.ino) await unlink(item.path);
    }
    throw error;
  }
  return { installed: plan.map(({ host, path }) => ({ host, path })) };
}

export async function uninstallPreviewRegistration(options) {
  // Exact-byte match required before removing any file. No legacy cleanup.
  const plan = await planPreviewRegistration(options);
  const removed = [];
  for (const item of plan) {
    try {
      if (!(await lstat(item.path)).isFile() || (await readFile(item.path, "utf8")) !== item.text)
        throw new Error("PREVIEW_REGISTRATION_CONFLICT: modified registration retained.");
      await unlink(item.path);
      removed.push({ host: item.host, path: item.path });
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
    }
  }
  return { removed };
}
