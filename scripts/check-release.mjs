// Offline release integrity check. No uploads, dependency install or model calls.
import { readFile, readdir, lstat } from "node:fs/promises";
import { createHash } from "node:crypto";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const manifestPath = "release/approved-files.json";
const sha = (bytes) => createHash("sha256").update(bytes).digest("hex");
export function scanText(text) {
  const problems = [];
  if (/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----\s*\n[A-Za-z0-9+/=]{20,}/.test(text))
    problems.push("private-key-material");
  if (
    /\b(?:gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|sk-ant-[A-Za-z0-9_-]{30,})\b/.test(
      text,
    )
  )
    problems.push("credential-pattern");
  if (/\/Users\/(?!you(?:\/|\b)|x(?:\/|\b)|test(?:\/|\b))[A-Za-z0-9_.-]+\//.test(text))
    problems.push("personal-home-path");
  return problems;
}
export async function checkRelease(directory = root) {
  const manifest = JSON.parse(await readFile(join(directory, manifestPath), "utf8"));
  const expected = new Map(manifest.files.map((f) => [f.path, f.sha256]));
  const found = [];
  const errors = [];
  async function visit(relative = "") {
    for (const name of await readdir(join(directory, relative))) {
      if (!relative && ["node_modules", ".git"].includes(name)) continue;
      const path = relative ? `${relative}/${name}` : name;
      if (path === manifestPath) continue;
      const meta = await lstat(join(directory, path));
      if (meta.isSymbolicLink()) {
        errors.push({ path, issue: "symlink" });
        continue;
      }
      if (meta.isDirectory()) {
        await visit(path);
        continue;
      }
      if (!meta.isFile()) {
        errors.push({ path, issue: "special-file" });
        continue;
      }
      found.push(path);
      const bytes = await readFile(join(directory, path));
      if (!expected.has(path)) errors.push({ path, issue: "not-allowlisted" });
      else if (expected.get(path) !== sha(bytes)) errors.push({ path, issue: "hash-mismatch" });
      if (!bytes.includes(0))
        for (const issue of scanText(bytes.toString("utf8"))) errors.push({ path, issue });
      if (
        /(?:^|\/)(?:\.env[^/]*|sessions\.json|transport\.json)$|\.(?:pem|key|p12|pfx|log)$|\.app\//i.test(
          path,
        )
      )
        errors.push({ path, issue: "private-or-binary-artifact" });
    }
  }
  await visit();
  for (const path of expected.keys())
    if (!found.includes(path)) errors.push({ path, issue: "missing-file" });
  return {
    status: errors.length ? "FAIL" : "PASS",
    files: found.length,
    errors,
    limits:
      "Pattern and hash checks only; native behavior, all secrets and legal clearance are not established.",
  };
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const result = await checkRelease();
    console.log(JSON.stringify(result, null, 2));
    process.exitCode = result.errors.length ? 1 : 0;
  } catch (error) {
    console.error("Release manifest unavailable or invalid; regenerate only after source review.");
    process.exitCode = 1;
  }
}
