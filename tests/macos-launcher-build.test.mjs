import test from "node:test";
import assert from "node:assert/strict";
import {
  mkdtemp,
  mkdir,
  writeFile,
  symlink,
  readFile,
  realpath,
  copyFile,
  stat,
} from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { tmpdir, userInfo } from "node:os";
import { join } from "node:path";
import {
  createLauncherConfiguration,
  validateEnvironment,
  validateOutputPath,
  buildLauncher,
  findCommand,
  copyProjectNotices,
  sourceHashes,
} from "../scripts/build-macos-launcher.mjs";

async function fixture() {
  const root = await mkdtemp(join(tmpdir(), "draftspect-launcher-test-"));
  const projectRoot = join(root, "Tester's 中文 Project");
  await mkdir(join(projectRoot, "app"), { recursive: true });
  await mkdir(join(projectRoot, "node_modules/electron/dist/Electron.app/Contents/MacOS"), {
    recursive: true,
  });
  await writeFile(
    join(projectRoot, "package.json"),
    JSON.stringify({
      name: "draftspect-office-addins",
      version: "0.2.0-preview.1",
      main: "app/main.mjs",
    }),
  );
  await writeFile(join(projectRoot, "app/main.mjs"), "// synthetic fixture\n");
  await writeFile(
    join(projectRoot, "node_modules/electron/dist/Electron.app/Contents/MacOS/Electron"),
    "synthetic executable",
    { mode: 0o700 },
  );
  const cliPath = join(root, "CodeBuddy's 中文 CLI");
  await writeFile(cliPath, "synthetic executable", { mode: 0o700 });
  return {
    root,
    projectRoot,
    cliPath,
    nodePath: process.execPath,
    environment: { CODEBUDDY_INTERNET_ENVIRONMENT: "external" },
  };
}

test("default CLI discovery uses filesystem checks without a platform-specific test executable", async () => {
  const paths = await fixture();
  await symlink(paths.cliPath, join(paths.root, "codebuddy"));
  assert.equal(
    await findCommand("codebuddy", `relative:${paths.root}`),
    await realpath(paths.cliPath),
  );
  await assert.rejects(findCommand("missing", `relative:${paths.root}`), /explicit/);
});

test("launcher configuration preserves spaces, Unicode and apostrophes without shell escaping", async () => {
  const paths = await fixture();
  const config = await createLauncherConfiguration({
    ...paths,
    environment: {
      CODEBUDDY_INTERNET_ENVIRONMENT: "ioa",
      CODEBUDDY_CONFIG_DIR: join(paths.root, "user's 中文 config"),
    },
  });
  assert.deepEqual(JSON.parse(JSON.stringify(config)), config);
  assert.match(config.projectRoot, /Tester's 中文 Project$/);
  assert.match(config.cliPath, /CodeBuddy's 中文 CLI$/);
  assert.equal(config.schemaVersion, 1);
  assert.deepEqual(Object.keys(config).sort(), [
    "cliPath",
    "environment",
    "nodePath",
    "projectRoot",
    "schemaVersion",
  ]);
});

test("launcher rejects secret or unknown env without reflecting their values", () => {
  for (const key of [
    "ANTHROPIC_API_KEY",
    "CODEBUDDY_API_KEY",
    "NODE_OPTIONS",
    "ELECTRON_RUN_AS_NODE",
    "PATH",
    "HOME",
    "DRAFTSPECT_RUNTIME",
    "UNRECOGNIZED",
  ]) {
    assert.throws(
      () => validateEnvironment({ [key]: "secret-must-not-appear" }),
      (error) => !error.message.includes("secret-must-not-appear"),
    );
  }
  assert.throws(
    () =>
      validateEnvironment({ CODEBUDDY_INTERNET_ENVIRONMENT: "https://user:secret@example.com" }),
    /Invalid/,
  );
  assert.throws(() => validateEnvironment({ DRAFTSPECT_STATE_DIR: "relative" }), /absolute/);
  assert.throws(
    () => validateEnvironment({ DRAFTSPECT_STATE_DIR: "/tmp/test\nSECRET" }),
    /Invalid/,
  );
  assert.throws(() => validateEnvironment({ DRAFTSPECT_TLS_KEY: "/tmp/key" }), /both/);
});

test("launcher requires real accessible dependencies and expected project identity", async () => {
  const paths = await fixture();
  await assert.rejects(
    createLauncherConfiguration({ ...paths, cliPath: join(paths.root, "missing") }),
    /CodeBuddy CLI/,
  );
  await assert.rejects(createLauncherConfiguration({ ...paths, nodePath: paths.root }), /Node.js/);
  await assert.rejects(
    createLauncherConfiguration({ ...paths, projectRoot: "relative" }),
    /absolute/,
  );
  await writeFile(
    join(paths.projectRoot, "package.json"),
    JSON.stringify({ name: "other-app", main: "app/main.mjs" }),
  );
  await assert.rejects(createLauncherConfiguration(paths), /expected Draftspect/);
});

test("launcher canonicalizes explicit CLI symlink", async () => {
  const paths = await fixture();
  const link = join(paths.root, "cli-link");
  await symlink(paths.cliPath, link);
  assert.equal(
    (await createLauncherConfiguration({ ...paths, cliPath: link })).cliPath,
    await realpath(paths.cliPath),
  );
});

test("build refuses existing output, dangling symlink, relative output and missing parent", async () => {
  const paths = await fixture();
  const existing = join(paths.root, "Existing.app");
  await mkdir(existing);
  await writeFile(join(existing, "keep"), "user content");
  await assert.rejects(buildLauncher({ ...paths, outputPath: existing }), /overwrite/);
  assert.equal(await readFile(join(existing, "keep"), "utf8"), "user content");
  const dangling = join(paths.root, "Dangling.app");
  await symlink(join(paths.root, "missing-target"), dangling);
  await assert.rejects(validateOutputPath(dangling), /overwrite/);
  await assert.rejects(validateOutputPath("relative.app"), /absolute/);
  await assert.rejects(validateOutputPath(join(paths.root, "not-an-app")), /\.app/);
  await assert.rejects(validateOutputPath(join(paths.root, "missing", "Draftspect.app")), /parent/);
  assert.match(await validateOutputPath(join(paths.root, "Draftspect.app")), /Draftspect\.app$/);
});

test("native launcher passes argument arrays and a fresh allowlisted environment", async () => {
  const source = await readFile(new URL("../app/macos-launcher.swift", import.meta.url), "utf8");
  assert.match(source, /child\.arguments = \[config\.projectRoot\]/);
  assert.match(source, /child\.environment = try launchEnvironment\(config\)/);
  assert.match(source, /var environment = config\.environment/);
  assert.match(source, /CODEBUDDY_CODE_PATH/);
  assert.match(source, /DRAFTSPECT_NODE_PATH/);
  assert.doesNotMatch(
    source,
    /ProcessInfo\.processInfo\.environment|\/bin\/(?:ba|z)?sh|osascript|do shell script/,
  );
  assert.match(source, /checkOnly/);
  assert.match(source, /O_NOFOLLOW/);
  assert.match(source, /getpwuid\(getuid\(\)\)/);
});

test("thin launcher copies project attribution verbatim and records exact notice hashes", async () => {
  const paths = await fixture();
  const resources = join(paths.root, "Resources");
  await mkdir(resources);
  const names = ["LICENSE", "UPSTREAM.md", "THIRD-PARTY-NOTICES.md"];
  for (const name of names)
    await copyFile(new URL(`../${name}`, import.meta.url), join(paths.projectRoot, name));
  const notices = await copyProjectNotices(paths.projectRoot, resources);
  assert.deepEqual(
    notices.map((n) => n.path),
    names.map((n) => `Contents/Resources/${n}`),
  );
  for (const [index, name] of names.entries()) {
    const original = await readFile(join(paths.projectRoot, name));
    assert.deepEqual(await readFile(join(resources, name)), original);
    assert.equal(notices[index].sha256, createHash("sha256").update(original).digest("hex"));
  }
  await assert.rejects(stat(join(resources, "node_modules")), { code: "ENOENT" });
});

test("launcher source hashing rejects symlink files and directories instead of omitting them", async () => {
  for (const isDirectory of [false, true]) {
    const paths = await fixture();
    const target = isDirectory ? paths.root : paths.cliPath;
    await symlink(target, join(paths.projectRoot, "app", "unreviewed-link"));
    await assert.rejects(sourceHashes(paths.projectRoot), /symlink entries/);
  }
});

test("thin launcher notice copying rejects missing, empty and symlinked notices", async () => {
  for (const kind of ["missing", "empty", "symlink"]) {
    const paths = await fixture();
    const resources = join(paths.root, "Resources");
    await mkdir(resources);
    await copyFile(new URL("../LICENSE", import.meta.url), join(paths.projectRoot, "LICENSE"));
    await copyFile(
      new URL("../THIRD-PARTY-NOTICES.md", import.meta.url),
      join(paths.projectRoot, "THIRD-PARTY-NOTICES.md"),
    );
    if (kind === "empty") await writeFile(join(paths.projectRoot, "UPSTREAM.md"), "");
    if (kind === "symlink") await symlink(paths.cliPath, join(paths.projectRoot, "UPSTREAM.md"));
    await assert.rejects(copyProjectNotices(paths.projectRoot, resources));
    await assert.rejects(stat(join(resources, "LICENSE")), { code: "ENOENT" });
  }
});

// Explicit compiled-binary acceptance mode; ordinary unit tests never compile or launch apps.
// Set this only to a launcher built by scripts/build-macos-launcher.mjs in the same run.
if (process.env.DRAFTSPECT_TEST_LAUNCHER_BINARY) {
  test("compiled native launcher supervises an isolated synthetic child with exact environment", async () => {
    const paths = await fixture();
    const bundle = join(paths.root, "Synthetic Launcher.app");
    await mkdir(join(bundle, "Contents/MacOS"), { recursive: true });
    await mkdir(join(bundle, "Contents/Resources"));
    const executable = join(bundle, "Contents/MacOS/DraftspectLauncher");
    await copyFile(process.env.DRAFTSPECT_TEST_LAUNCHER_BINARY, executable);
    await writeFile(
      join(bundle, "Contents/Info.plist"),
      `<?xml version="1.0"?><plist version="1.0"><dict><key>CFBundleExecutable</key><string>DraftspectLauncher</string><key>CFBundleIdentifier</key><string>local.draftspect.synthetic.handoff</string><key>CFBundlePackageType</key><string>APPL</string></dict></plist>`,
    );
    const state = join(paths.root, "isolated-state");
    const config = await createLauncherConfiguration({
      ...paths,
      environment: { DRAFTSPECT_STATE_DIR: state, CODEBUDDY_INTERNET_ENVIRONMENT: "external" },
    });
    const configPath = join(bundle, "Contents/Resources/launcher-config.json");
    await writeFile(configPath, JSON.stringify(config));
    const check = spawnSync(executable, ["--check"], { encoding: "utf8", timeout: 5000 });
    assert.equal(check.status, 0, check.stderr);
    assert.equal(JSON.parse(check.stdout).ok, true);
    await assert.rejects(stat(state), { code: "ENOENT" });
    for (const account of [undefined, "not-an-environment"]) {
      const environment = { ...config.environment };
      if (account === undefined) delete environment.CODEBUDDY_INTERNET_ENVIRONMENT;
      else environment.CODEBUDDY_INTERNET_ENVIRONMENT = account;
      await writeFile(configPath, JSON.stringify({ ...config, environment }));
      const invalid = spawnSync(executable, ["--check"], { encoding: "utf8", timeout: 5000 });
      assert.equal(invalid.status, 1);
      assert.equal(JSON.parse(invalid.stderr).ok, false);
      await assert.rejects(stat(state), { code: "ENOENT" });
    }
    for (const field of ["projectRoot", "cliPath", "nodePath"]) {
      await writeFile(configPath, JSON.stringify({ ...config, [field]: "/tmp/invalid\tpath" }));
      const invalid = spawnSync(executable, ["--check"], { encoding: "utf8", timeout: 5000 });
      assert.equal(invalid.status, 1, field);
      assert.equal(JSON.parse(invalid.stderr).ok, false);
      await assert.rejects(stat(state), { code: "ENOENT" });
    }
    await writeFile(configPath, JSON.stringify(config));
    const syntheticElectron = join(
      paths.projectRoot,
      "node_modules/electron/dist/Electron.app/Contents/MacOS/Electron",
    );
    await writeFile(
      syntheticElectron,
      `#!${process.execPath}\nconst fs=require('node:fs'); const path=require('node:path'); const root=process.argv[2]; fs.writeFileSync(path.join(root,'started.json'),JSON.stringify({pid:process.pid,argv:process.argv.slice(2),cwd:process.cwd(),home:process.env.HOME,runtime:process.env.DRAFTSPECT_RUNTIME,nodePath:process.env.DRAFTSPECT_NODE_PATH,cliPath:process.env.CODEBUDDY_CODE_PATH,injectedNode:process.env.NODE_OPTIONS||null,injectedElectron:process.env.ELECTRON_RUN_AS_NODE||null,state:process.env.DRAFTSPECT_STATE_DIR})); const state={schemaVersion:1,token:process.env.DRAFTSPECT_LAUNCH_TOKEN,pid:process.pid,phase:'desktop-visible',desktopVisible:true,serviceReady:true}; const publish=()=>fs.writeFileSync(process.env.DRAFTSPECT_LAUNCH_STATUS_PATH,JSON.stringify(state)); publish(); console.log('synthetic-secret-must-not-be-logged'); setTimeout(()=>{fs.writeFileSync(path.join(root,'completed.json'),JSON.stringify({pid:process.pid,finished:true}));state.phase='quitting';state.serviceReady=false;publish();},1200);\n`,
      { mode: 0o700 },
    );
    const began = Date.now();
    const run = spawnSync(executable, [], {
      encoding: "utf8",
      // Cold process registration plus the production monitor's 2-second poll
      // can exceed 5 seconds. Keep a bounded test budget, not a production change.
      timeout: 15000,
      env: {
        ...process.env,
        NODE_OPTIONS: "--require /must-not-run.mjs",
        ELECTRON_RUN_AS_NODE: "1",
        DRAFTSPECT_STATE_DIR: "/must-not-use",
      },
    });
    assert.equal(run.status, 0, run.error?.message || run.stderr);
    assert.ok(Date.now() - began >= 1100, "launcher must monitor its child through quit");
    async function waitForFile(path, timeout = 5000) {
      const deadline = Date.now() + timeout;
      while (Date.now() < deadline) {
        try {
          return JSON.parse(await readFile(path, "utf8"));
        } catch (error) {
          if (error.code !== "ENOENT") throw error;
        }
        await new Promise((resolve) => setTimeout(resolve, 30));
      }
      throw new Error("Synthetic child output did not arrive");
    }
    const started = await waitForFile(join(paths.projectRoot, "started.json"));
    assert.throws(() => process.kill(started.pid, 0), { code: "ESRCH" });
    assert.deepEqual(started.argv, [config.projectRoot]);
    assert.equal(started.cwd, config.projectRoot);
    assert.equal(started.home, userInfo().homedir);
    assert.equal(started.runtime, "codebuddy");
    assert.equal(started.nodePath, config.nodePath);
    assert.equal(started.cliPath, config.cliPath);
    assert.equal(started.state, state);
    assert.equal(started.injectedNode, null);
    assert.equal(started.injectedElectron, null);
    const completed = await waitForFile(join(paths.projectRoot, "completed.json"));
    assert.equal(completed.finished, true);
    const log = await readFile(join(state, "launcher.log"), "utf8");
    const records = log.trim().split("\n").map(JSON.parse);
    assert.equal(records[0].projectVersion, "0.2.0-preview.1");
    assert.deepEqual(
      records.map((record) => record.event),
      ["launch-requested", "spawned", "desktop-visible", "manual-quit"],
    );
    assert.equal((await stat(join(state, "launcher.log"))).mode & 0o777, 0o600);
    assert.doesNotMatch(log, /must-not-run|must-not-use|synthetic-secret|DRAFTSPECT_LAUNCH_TOKEN/);
  });
}
