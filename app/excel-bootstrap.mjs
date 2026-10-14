// User-triggered, marker-only fallback for Mac hosts whose ordinary gallery
// does not activate the add-in. Never bind the user's workbook or clear caches.
import { createHash } from "node:crypto";
import { readFile, mkdtemp, writeFile } from "node:fs/promises";
import { homedir, tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { loadLocalTransport } from "../daemon/local-transport.mjs";
import { inspectManifestRegistration } from "./addin-registration.mjs";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
export const EXCEL_BOOTSTRAP = Object.freeze({
  asset: "app/assets/excel-initialize-synthetic.xlsx",
  id: "4b5861d6-18ee-4d4a-b9ee-d43d7da80ff4",
  version: "1.0.0.0",
  sha256: "c455fda12cb5110f46bdb0fe03ecfd127a99c364436b4756c921e3bf81ddefc0",
});

export function verifyExcelBootstrap({ manifestXml, installedManifestXml, assetBytes, transport }) {
  const registration = inspectManifestRegistration({
    projectXml: manifestXml,
    installedXml: installedManifestXml,
    transport,
  });
  if (registration.state === "invalid")
    throw new Error("EXCEL_BOOTSTRAP_MANIFEST_INVALID: unsupported manifest XML.");
  const clean = manifestXml.replace(/<!--[\s\S]*?-->/g, "");
  if (
    registration.state !== "matched" ||
    !clean.includes(`<Id>${EXCEL_BOOTSTRAP.id}</Id>`) ||
    !clean.includes(`<Version>${EXCEL_BOOTSTRAP.version}</Version>`)
  )
    throw new Error(
      "EXCEL_BOOTSTRAP_MANIFEST_MISMATCH: install matching Excel manifest, URLs and command resources first.",
    );
  if (createHash("sha256").update(assetBytes).digest("hex") !== EXCEL_BOOTSTRAP.sha256)
    throw new Error("EXCEL_BOOTSTRAP_ASSET_MISMATCH: restore the verified synthetic workbook.");
  return { id: EXCEL_BOOTSTRAP.id, version: EXCEL_BOOTSTRAP.version };
}

export async function openExcelBootstrap({
  openPath,
  projectRoot = ROOT,
  installedManifestPath = join(
    homedir(),
    "Library",
    "Containers",
    "com.microsoft.Excel",
    "Data",
    "Documents",
    "wef",
    "draftspect-preview-excel.xml",
  ),
  tempRoot = tmpdir(),
  platform = process.platform,
  transport,
} = {}) {
  if (platform !== "darwin")
    throw new Error("This initialization path is only verified for Excel on macOS.");
  if (typeof openPath !== "function") throw new TypeError("A user-triggered openPath is required.");
  let installedManifestXml;
  try {
    installedManifestXml = await readFile(installedManifestPath, "utf8");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    throw new Error("EXCEL_BOOTSTRAP_NOT_INSTALLED: install the Excel add-in first.");
  }
  const [manifestXml, assetBytes, settings] = await Promise.all([
    readFile(join(projectRoot, "manifests", "excel.xml"), "utf8"),
    readFile(join(projectRoot, EXCEL_BOOTSTRAP.asset)),
    transport || loadLocalTransport(),
  ]);
  verifyExcelBootstrap({ manifestXml, installedManifestXml, assetBytes, transport: settings });
  const dir = await mkdtemp(join(tempRoot, "draftspect-excel-initialize-"));
  const path = join(dir, "Draftspect-SYNTHETIC-Initialize-Excel.xlsx");
  await writeFile(path, assetBytes, { flag: "wx", mode: 0o600 });
  const error = await openPath(path);
  if (error) throw new Error(`EXCEL_BOOTSTRAP_OPEN_FAILED: ${error}`);
  // OS-open success is not a pane handshake. The user must confirm the pane
  // in this workbook, then use Home -> Show Draftspect in the target workbook.
  return { path, requiresEachExcelProcess: true };
}
