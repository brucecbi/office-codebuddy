// User-triggered fallback when the ordinary Word gallery does not activate.
// Open a verified synthetic document, never bind user files or run automatically.
import { createHash } from "node:crypto";
import { readFile, mkdtemp, writeFile } from "node:fs/promises";
import { homedir, tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const PROJECT_ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");

// Updating the manifest identity requires a newly validated bound asset and
// hash. Fail closed on version drift instead of opening an old binding.
export const WORD_BOOTSTRAP = Object.freeze({
  asset: "app/assets/word-initialize-synthetic.docx",
  id: "27ee7761-18ca-4fae-b482-4171d5afa6cd",
  version: "1.0.0.1",
  sha256: "461917ce2481078671223b9f98bc268228176d00e9c9a3357bc643f7c0e90599",
});

function manifestIdentity(xml) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) {
    throw new Error("WORD_BOOTSTRAP_MANIFEST_INVALID: unsupported manifest XML.");
  }
  const clean = xml.replace(/<!--[\s\S]*?-->/g, "");
  const ids = [...clean.matchAll(/<Id>\s*([^<]+?)\s*<\/Id>/g)];
  const versions = [...clean.matchAll(/<Version>\s*([^<]+?)\s*<\/Version>/g)];
  if (ids.length !== 1 || versions.length !== 1) {
    throw new Error("WORD_BOOTSTRAP_MANIFEST_INVALID: expected one Id and Version.");
  }
  return { id: ids[0][1].toLowerCase(), version: versions[0][1] };
}

export function verifyWordBootstrap({ manifestXml, installedManifestXml, assetBytes }) {
  for (const [label, xml] of [
    ["Project", manifestXml],
    ["Installed Word", installedManifestXml],
  ]) {
    const identity = manifestIdentity(xml);
    if (identity.id !== WORD_BOOTSTRAP.id || identity.version !== WORD_BOOTSTRAP.version) {
      throw new Error(
        `WORD_BOOTSTRAP_MANIFEST_MISMATCH: ${label} manifest and the synthetic bootstrap ` +
          "must have the same GUID and version. Install matching project assets first.",
      );
    }
  }
  const digest = createHash("sha256").update(assetBytes).digest("hex");
  if (digest !== WORD_BOOTSTRAP.sha256) {
    throw new Error(
      "WORD_BOOTSTRAP_ASSET_MISMATCH: the synthetic bootstrap was changed; restore the verified asset.",
    );
  }
  return { id: WORD_BOOTSTRAP.id, version: WORD_BOOTSTRAP.version };
}

export async function openWordBootstrap({
  openPath,
  projectRoot = PROJECT_ROOT,
  installedManifestPath = join(
    homedir(),
    "Library",
    "Containers",
    "com.microsoft.Word",
    "Data",
    "Documents",
    "wef",
    "draftspect-preview-word.xml",
  ),
  tempRoot = tmpdir(),
  platform = process.platform,
} = {}) {
  if (platform !== "darwin") {
    throw new Error("This initialization path is only verified for Word on macOS.");
  }
  if (typeof openPath !== "function") throw new TypeError("A user-triggered openPath is required.");
  let installedManifestXml;
  try {
    installedManifestXml = await readFile(installedManifestPath, "utf8");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    throw new Error(
      "WORD_BOOTSTRAP_NOT_INSTALLED: install the Word add-in first from the tray menu.",
    );
  }
  const [manifestXml, assetBytes] = await Promise.all([
    readFile(join(projectRoot, "manifests", "word.xml"), "utf8"),
    readFile(join(projectRoot, WORD_BOOTSTRAP.asset)),
  ]);
  verifyWordBootstrap({ manifestXml, installedManifestXml, assetBytes });

  // Word may save its synthetic setup or add-in state. Give it a private fresh
  // copy so repeated initialization cannot modify the verified packaged asset.
  const dir = await mkdtemp(join(tempRoot, "draftspect-word-initialize-"));
  const path = join(dir, "Draftspect-SYNTHETIC-Initialize-Word.docx");
  await writeFile(path, assetBytes, { flag: "wx", mode: 0o600 });
  const error = await openPath(path);
  if (error) throw new Error(`WORD_BOOTSTRAP_OPEN_FAILED: ${error}`);
  // An OS-open request is not native-ready evidence and is not a permanent
  // registration. The tray instructions ask the user to confirm Word's pane.
  return { path, requiresEachWordProcess: true };
}
