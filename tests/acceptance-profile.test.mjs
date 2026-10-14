import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir, homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { createRequire } from "node:module";
import vm from "node:vm";
import {
  loadAcceptanceProfile,
  acceptanceAllowsDocument,
  assertAcceptanceWorkspace,
  acceptanceSdkOptions,
  acceptanceCliArguments,
} from "../daemon/acceptance-profile.mjs";

const repo = fileURLToPath(new URL("../", import.meta.url));
let root, env, profile;
before(async () => {
  root = await mkdtemp(join(tmpdir(), "draftspect-acceptance-profile-test-"));
  const document = join(root, "Synthetic-Session-Handoff.docx");
  await writeFile(document, "Synthetic configuration fixture; not a native DOCX test");
  await mkdir(join(root, "state"));
  env = {
    DRAFTSPECT_ACCEPTANCE_DOCUMENT: document,
    DRAFTSPECT_ACCEPTANCE_MODEL: "synthetic-model-b",
    DRAFTSPECT_STATE_DIR: join(root, "state"),
    DRAFTSPECT_RUNTIME: "codebuddy",
  };
  profile = loadAcceptanceProfile(env);
});
after(async () => {
  if (root) await rm(root, { recursive: true });
});

test("acceptance defaults off, refuses partial scope, and protects production state", () => {
  assert.equal(loadAcceptanceProfile({}), null);
  for (const patch of [
    { DRAFTSPECT_ACCEPTANCE_DOCUMENT: undefined },
    { DRAFTSPECT_ACCEPTANCE_MODEL: undefined },
    { DRAFTSPECT_ACCEPTANCE_DOCUMENT: "relative.docx" },
    { DRAFTSPECT_ACCEPTANCE_MODEL: "model with spaces" },
    { DRAFTSPECT_STATE_DIR: undefined },
    { DRAFTSPECT_STATE_DIR: root },
    { DRAFTSPECT_STATE_DIR: homedir() },
    { DRAFTSPECT_RUNTIME: "claude" },
    { DRAFTSPECT_RUNTIME: undefined },
  ])
    assert.throws(() => loadAcceptanceProfile({ ...env, ...patch }), /Acceptance/);
});

test("acceptance binds exact Word document and canonical workspace, including file URLs", () => {
  assert.equal(acceptanceAllowsDocument("word", env.DRAFTSPECT_ACCEPTANCE_DOCUMENT, profile), true);
  assert.equal(
    acceptanceAllowsDocument(
      "word",
      pathToFileURL(env.DRAFTSPECT_ACCEPTANCE_DOCUMENT).href,
      profile,
    ),
    true,
  );
  for (const [host, document] of [
    ["excel", profile.document],
    ["word", null],
    ["word", join(root, "other.docx")],
    ["word", "https://example.invalid/document"],
  ])
    assert.equal(acceptanceAllowsDocument(host, document, profile), false);
  assertAcceptanceWorkspace(profile.workspace, profile);
  assert.throws(() => assertAcceptanceWorkspace(homedir(), profile), /fixed/);
});

test("acceptance SDK options force empty built-ins/settings/hooks and refuse non-Office MCP/model/cwd", async () => {
  const office = { type: "sdk" };
  const original = {
    cwd: profile.workspace,
    model: profile.model,
    mcpServers: { office },
    tools: ["Bash"],
    settingSources: ["user"],
    hooks: { unsafe: [] },
    allowedTools: ["Bash"],
    extraArgs: { tools: "Bash" },
    canUseTool: async () => ({ behavior: "allow" }),
  };
  assert.equal(acceptanceSdkOptions(original, null), original);
  const options = acceptanceSdkOptions(original, profile);
  assert.deepEqual(options.tools, []);
  assert.deepEqual(options.settingSources, []);
  assert.deepEqual(options.hooks, {});
  assert.deepEqual(options.allowedTools, []);
  assert.deepEqual(options.extraArgs, {});
  assert.equal(options.strictMcpConfig, true);
  assert.equal(options.maxTurns, 12);
  assert.deepEqual(Object.keys(options.mcpServers), ["office"]);
  assert.equal((await options.canUseTool("mcp__office__office_read_story", {})).behavior, "allow");
  for (const name of [
    "Bash",
    "Read",
    "mcp__other__office_read_story",
    "mcp__office__excel_read_range",
  ])
    assert.equal((await options.canUseTool(name, {})).behavior, "deny");
  for (const patch of [
    { model: "other" },
    { cwd: homedir() },
    { mcpServers: {} },
    { mcpServers: { office, external: {} } },
  ])
    assert.throws(() => acceptanceSdkOptions({ ...original, ...patch }, profile), /Acceptance/);
  // Installed SDK argument builder only: no start/connect/spawn or MCP instance.
  const require = createRequire(import.meta.url);
  const { ProcessTransport } = require(
    join(dirname(require.resolve("@tencent-ai/agent-sdk")), "transport/process-transport.js"),
  );
  const transport = new ProcessTransport({ ...options, mcpServers: {} });
  const args = transport.buildArgs();
  assert.equal(args[args.indexOf("--tools") + 1], "");
  assert.equal(args[args.indexOf("--setting-sources") + 1], "none");
  assert.equal(args[args.indexOf("--max-turns") + 1], "12");
  assert.ok(args.includes("--strict-mcp-config"));
  assert.equal(transport.process, null);
});

test("actual user MCP loader returns before any user config read under acceptance", async () => {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const loader = source.slice(
    source.indexOf("async function loadUserMcpServers()"),
    source.indexOf("// Hoisted ABOVE the top-level awaits"),
  );
  const context = vm.createContext({
    acceptanceProfile: profile,
    readFile() {
      throw new Error("User configuration was read");
    },
  });
  vm.runInContext(loader + "\nglobalThis.load=loadUserMcpServers;", context);
  assert.equal(Object.keys(await context.load()).length, 0);
  assert.match(source, /allowHello: acceptanceProfile \? acceptanceAllowsDocument : undefined/);
});

test("profile runtime selection and fixed test catalog do not spawn discovery; CLI commands preserve isolation", async () => {
  const script = `
    import runtime from './daemon/runtime/index.mjs';
    import { resumeCommands } from './daemon/session-archive.mjs';
    import { getSessionId } from './daemon/sessions.mjs';
    const models=await runtime.listModels();
    const id=await getSessionId('word',process.env.DRAFTSPECT_ACCEPTANCE_DOCUMENT,'codebuddy','new-doc');
    const commands=resumeCommands({provider:'codebuddy',session_id:'synthetic-session',cwd:${JSON.stringify(root)}});
    console.log(JSON.stringify({runtime:runtime.id,models,id,commands}));
  `;
  const { stdout, stderr } = await promisify(execFile)(
    process.execPath,
    ["--input-type=module", "-e", script],
    { cwd: repo, env: { PATH: process.env.PATH, ...env } },
  );
  assert.equal(stderr, "");
  assert.match(stdout, /provider discovery NOT_RUN/);
  const result = JSON.parse(stdout.trim().split("\n").at(-1));
  assert.equal(result.runtime, "codebuddy");
  assert.equal(result.id, null);
  assert.equal(result.models.length, 1);
  assert.equal(result.models[0].value, profile.model);
  for (const flag of [
    "--strict-mcp-config",
    "--tools",
    "--setting-sources",
    "synthetic-model-b",
    "--max-turns",
  ])
    assert.ok(result.commands.resume.includes(flag));
  assert.deepEqual(acceptanceCliArguments(null), []);
  assert.ok(acceptanceCliArguments(profile).includes(""));
});
