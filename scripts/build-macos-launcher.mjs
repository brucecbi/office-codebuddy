#!/usr/bin/env node
// A local thin wrapper: no bundled credentials, downloads, login items or service startup.
import { constants } from "node:fs";
import {
  access,
  chmod,
  copyFile,
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  rename,
  writeFile,
} from "node:fs/promises";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { dirname, extname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { codebuddyEnvironment } from "../daemon/preview-policy.mjs";

export const ALLOWED_ENVIRONMENT = Object.freeze([
  "DRAFTSPECT_STATE_DIR",
  "CODEBUDDY_CONFIG_DIR",
  "CODEBUDDY_INTERNET_ENVIRONMENT",
  "DRAFTSPECT_TLS_KEY",
  "DRAFTSPECT_TLS_CERT",
]);
const SOURCE_ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const NO_CONTROL = /^[^\x00-\x1f\x7f]+$/;

function absolutePath(value, label) {
  if (typeof value !== "string" || !isAbsolute(value) || !NO_CONTROL.test(value)) {
    throw new Error(`${label} must be an absolute path without control characters`);
  }
  return resolve(value);
}

export function validateEnvironment(environment = {}) {
  if (!environment || typeof environment !== "object" || Array.isArray(environment)) {
    throw new Error("environment must be a plain key/value object");
  }
  const clean = {};
  for (const [key, value] of Object.entries(environment)) {
    if (!ALLOWED_ENVIRONMENT.includes(key)) throw new Error("Unsupported launcher environment key");
    if (typeof value !== "string" || !NO_CONTROL.test(value))
      throw new Error(`Invalid launcher environment value for ${key}`);
    if (key === "CODEBUDDY_INTERNET_ENVIRONMENT") {
      if (!/^[A-Za-z0-9_-]{1,32}$/.test(value))
        throw new Error("Invalid CodeBuddy environment name");
      clean[key] = value;
    } else {
      clean[key] = absolutePath(value, key);
    }
  }
  if (Boolean(clean.DRAFTSPECT_TLS_KEY) !== Boolean(clean.DRAFTSPECT_TLS_CERT)) {
    throw new Error("Set both TLS path overrides or neither");
  }
  return clean;
}

async function requireFile(value, label, executable = false) {
  const path = absolutePath(value, label);
  try {
    const canonical = await realpath(path);
    if (!(await lstat(canonical)).isFile()) throw new Error();
    await access(canonical, executable ? constants.X_OK | constants.R_OK : constants.R_OK);
    return canonical;
  } catch {
    throw new Error(`${label} must point to an accessible ${executable ? "executable " : ""}file`);
  }
}

export async function createLauncherConfiguration({ projectRoot, cliPath, nodePath, environment }) {
  const root = await realpath(absolutePath(projectRoot, "projectRoot")).catch(() => {
    throw new Error("projectRoot is unavailable");
  });
  await requireFile(join(root, "app", "main.mjs"), "Draftspect entry");
  await requireFile(
    join(root, "node_modules", "electron", "dist", "Electron.app", "Contents", "MacOS", "Electron"),
    "Electron",
    true,
  );
  let packageInfo;
  try {
    packageInfo = JSON.parse(await readFile(join(root, "package.json"), "utf8"));
  } catch {
    throw new Error("Cannot read project package.json");
  }
  if (packageInfo.name !== "draftspect-office-addins" || packageInfo.main !== "app/main.mjs")
    throw new Error("projectRoot is not the expected Draftspect project");
  const cleanEnvironment = validateEnvironment(environment);
  codebuddyEnvironment(cleanEnvironment);
  for (const key of ["DRAFTSPECT_TLS_KEY", "DRAFTSPECT_TLS_CERT"]) {
    if (cleanEnvironment[key])
      cleanEnvironment[key] = await requireFile(cleanEnvironment[key], key);
  }
  return {
    schemaVersion: 1,
    projectRoot: root,
    cliPath: await requireFile(cliPath, "CodeBuddy CLI", true),
    nodePath: await requireFile(nodePath, "Node.js", true),
    environment: cleanEnvironment,
  };
}

export async function validateOutputPath(outputPath) {
  const output = absolutePath(outputPath, "outputPath");
  if (extname(output) !== ".app") throw new Error("outputPath must end with .app");
  try {
    await lstat(output);
    throw new Error("Refusing to overwrite an existing output path or symlink");
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  const parent = await realpath(dirname(output)).catch(() => {
    throw new Error("Output parent directory must already exist");
  });
  return join(parent, output.slice(dirname(output).length + 1));
}

function run(command, args, label, options = {}) {
  const result = spawnSync(command, args, { encoding: "utf8", timeout: 120_000, ...options });
  if (result.error || result.status !== 0) {
    // Compiler output contains source diagnostics, not application environment/config values.
    const detail =
      label === "Swift compilation" ? `\n${String(result.stderr || "").slice(0, 6000)}` : "";
    throw new Error(`${label} failed${detail}`);
  }
  return result.stdout.trim();
}

export async function sourceHashes(root) {
  const paths = [
    "package.json",
    "package-lock.json",
    "scripts/build-macos-launcher.mjs",
    "LICENSE",
    "UPSTREAM.md",
    "THIRD-PARTY-NOTICES.md",
  ];
  async function addDirectory(directory) {
    const directoryInfo = await lstat(join(root, directory));
    if (!directoryInfo.isDirectory() || directoryInfo.isSymbolicLink()) {
      throw new Error("Source snapshot roots must be real project directories");
    }
    for (const entry of await readdir(join(root, directory), { withFileTypes: true })) {
      const relative = join(directory, entry.name);
      if (entry.isSymbolicLink()) throw new Error("Source snapshot refuses symlink entries");
      if (entry.isDirectory()) await addDirectory(relative);
      else if (entry.isFile()) paths.push(relative);
    }
  }
  await addDirectory("app");
  await addDirectory("daemon");
  await addDirectory("taskpane");
  await addDirectory("manifests");
  const hashes = {};
  for (const relative of [...new Set(paths)].sort()) {
    const info = await lstat(join(root, relative));
    if (!info.isFile() || info.isSymbolicLink()) {
      throw new Error("Source snapshot entries must be regular project files");
    }
    hashes[relative] = createHash("sha256")
      .update(await readFile(join(root, relative)))
      .digest("hex");
  }
  return hashes;
}

// Only project notices belong in this thin wrapper. External runtimes are not copied.
export async function copyProjectNotices(projectRoot, resources) {
  const names = ["LICENSE", "UPSTREAM.md", "THIRD-PARTY-NOTICES.md"];
  const files = [];
  // Validate every source before copying any notice; refuse symlink indirection.
  for (const name of names) {
    const path = join(projectRoot, name);
    const info = await lstat(path);
    if (!info.isFile() || info.isSymbolicLink())
      throw new Error(`${name} must be a regular project notice file`);
    const bytes = await readFile(path);
    if (!bytes.length) throw new Error(`${name} must not be empty`);
    files.push({ name, bytes });
  }
  const notices = [];
  for (const { name, bytes } of files) {
    await writeFile(join(resources, name), bytes, { flag: "wx", mode: 0o600 });
    notices.push({
      path: `Contents/Resources/${name}`,
      sha256: createHash("sha256").update(bytes).digest("hex"),
    });
  }
  return notices;
}

export async function buildLauncher({
  projectRoot = SOURCE_ROOT,
  outputPath,
  cliPath,
  nodePath = process.execPath,
  environment = {},
}) {
  const output = await validateOutputPath(outputPath);
  const configuration = await createLauncherConfiguration({
    projectRoot,
    cliPath,
    nodePath,
    environment,
  });
  if (process.platform !== "darwin")
    throw new Error("The local launcher can only be built on macOS");
  const iconPath = await requireFile(
    join(configuration.projectRoot, "app", "assets", "CodeWord.icns"),
    "Application icon",
  );
  const swiftc = run("/usr/bin/xcrun", ["--find", "swiftc"], "Swift compiler discovery");
  const sdkPath = run("/usr/bin/xcrun", ["--show-sdk-path"], "macOS SDK discovery");
  const swiftVersion = run(swiftc, ["--version"], "Swift compiler version").split("\n")[0];
  const hashes = await sourceHashes(configuration.projectRoot);
  const packageInfo = JSON.parse(
    await readFile(join(configuration.projectRoot, "package.json"), "utf8"),
  );
  if (!/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(packageInfo.version || ""))
    throw new Error("Invalid project version for launcher metadata");
  // Apple bundle versions require numeric components. Keep the full SemVer
  // (including preview suffix) in project metadata, not in these numeric keys.
  const bundleVersion = packageInfo.version.split("-")[0];
  const staging = await mkdtemp(join(dirname(output), ".draftspect-build-"));
  const stagedApp = join(staging, "Draftspect.app");
  const contents = join(stagedApp, "Contents");
  await mkdir(join(contents, "MacOS"), { recursive: true, mode: 0o700 });
  await mkdir(join(contents, "Resources"), { mode: 0o700 });
  const executable = join(contents, "MacOS", "DraftspectLauncher");
  try {
    const projectNotices = await copyProjectNotices(
      configuration.projectRoot,
      join(contents, "Resources"),
    );
    await writeFile(
      join(contents, "Resources", "launcher-config.json"),
      `${JSON.stringify(configuration, null, 2)}\n`,
      { mode: 0o600 },
    );
    await copyFile(iconPath, join(contents, "Resources", "CodeWord.icns"));
    await writeFile(
      join(contents, "Info.plist"),
      `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
<key>CFBundleExecutable</key><string>DraftspectLauncher</string>
<key>CFBundleIdentifier</key><string>local.draftspect.preview.launcher</string>
<key>CFBundleName</key><string>Draftspect Preview</string>
<key>CFBundleDisplayName</key><string>Draftspect Preview</string>
<key>CFBundleIconFile</key><string>CodeWord.icns</string>
<key>CFBundlePackageType</key><string>APPL</string>
<key>CFBundleShortVersionString</key><string>${bundleVersion}</string>
<key>CFBundleVersion</key><string>${bundleVersion}</string>
<key>DraftspectProjectVersion</key><string>${packageInfo.version}</string>
<key>LSMinimumSystemVersion</key><string>13.0</string>
<key>LSUIElement</key><true/>
<key>NSHighResolutionCapable</key><true/>
</dict></plist>\n`,
      { mode: 0o600 },
    );
    run(
      swiftc,
      [
        "-parse-as-library",
        "-O",
        "-sdk",
        sdkPath,
        "-target",
        `${process.arch === "arm64" ? "arm64" : "x86_64"}-apple-macosx13.0`,
        "-framework",
        "AppKit",
        "-module-cache-path",
        join(staging, "module-cache"),
        join(configuration.projectRoot, "app", "macos-launcher.swift"),
        join(configuration.projectRoot, "app", "macos-launch-monitor.swift"),
        "-o",
        executable,
      ],
      "Swift compilation",
    );
    await chmod(executable, 0o700);
    const binaryHash = createHash("sha256")
      .update(await readFile(executable))
      .digest("hex");
    const electronInfo = JSON.parse(
      await readFile(
        join(configuration.projectRoot, "node_modules", "electron", "package.json"),
        "utf8",
      ),
    );
    const metadata = {
      schemaVersion: 1,
      builtAt: new Date().toISOString(),
      platform: process.platform,
      architecture: process.arch,
      projectVersion: packageInfo.version,
      electronVersion: electronInfo.version,
      builderNodeVersion: process.version,
      swiftVersion,
      sourceHashes: hashes,
      projectNotices,
      unsignedBinarySha256: binaryHash,
      scope:
        "Build-time source snapshot only; thin launcher follows the live external project after updates. No immutable installation, runtime or Word acceptance implied.",
    };
    await writeFile(
      join(contents, "Resources", "build-info.json"),
      `${JSON.stringify(metadata, null, 2)}\n`,
      { mode: 0o600 },
    );
    run("/usr/bin/plutil", ["-lint", join(contents, "Info.plist")], "Bundle plist validation");
    run("/usr/bin/codesign", ["--force", "--sign", "-", stagedApp], "Local ad-hoc signing");
    run("/usr/bin/codesign", ["--verify", "--strict", stagedApp], "Local signature verification");
    run(executable, ["--check"], "Native configuration check");
    // Exclusive mkdir closes the preflight-to-publication overwrite race.
    // A failed publication is retained for diagnosis, never recursively deleted.
    await mkdir(output, { mode: 0o700 });
    await rename(contents, join(output, "Contents"));
    return { outputPath: output, buildInfo: metadata, stagingPath: staging };
  } catch (error) {
    throw new Error(`${error.message}\nFresh build artifacts retained at: ${staging}`);
  }
}

export async function findCommand(name, searchPath = process.env.PATH || "") {
  for (const directory of searchPath.split(":")) {
    if (!isAbsolute(directory)) continue;
    const candidate = join(directory, name);
    try {
      return await requireFile(candidate, "command", true);
    } catch {
      // No shell or platform-specific /usr/bin/test dependency for PATH lookup.
    }
  }
  throw new Error(`Pass an explicit --${name === "codebuddy" ? "cli" : "node"} path`);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const options = { projectRoot: SOURCE_ROOT, cliPath: undefined, nodePath: process.execPath };
    const flags = {
      "--output": "outputPath",
      "--project-root": "projectRoot",
      "--cli": "cliPath",
      "--node": "nodePath",
    };
    for (let i = 2; i < process.argv.length; i += 2) {
      const key = flags[process.argv[i]];
      if (!key || !process.argv[i + 1])
        throw new Error(
          "Usage: node scripts/build-macos-launcher.mjs --output ABS.app [--project-root ABS] [--cli ABS] [--node ABS]",
        );
      options[key] = process.argv[i + 1];
    }
    options.cliPath ||= await findCommand("codebuddy");
    // Only documented nonsecret overrides are copied; no auth/settings file is read.
    options.environment = Object.fromEntries(
      ALLOWED_ENVIRONMENT.filter((key) => process.env[key]).map((key) => [key, process.env[key]]),
    );
    const result = await buildLauncher(options);
    console.log(
      JSON.stringify(
        {
          outputPath: result.outputPath,
          stagingPath: result.stagingPath,
          scope: result.buildInfo.scope,
        },
        null,
        2,
      ),
    );
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
