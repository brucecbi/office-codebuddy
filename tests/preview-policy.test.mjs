import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  agentAccess,
  codebuddyEnvironment,
  userSettingsSources,
  forwardUserMcp,
  previewSdkOptions,
} from "../daemon/preview-policy.mjs";
import { stateDirectory } from "../daemon/state-paths.mjs";
import { homedir } from "node:os";
import { join } from "node:path";
import { loadProcessTransport } from "./helpers/codebuddy-transport.mjs";
const ProcessTransport = await loadProcessTransport();

test("default preview never forwards user MCP/settings even with legacy opt-in variables", () => {
  const env = { DRAFTSPECT_CODEBUDDY_SETTING_SOURCES: "user", DRAFTSPECT_FORWARD_USER_MCP: "1" };
  assert.equal(agentAccess(env), "office-only");
  assert.deepEqual(userSettingsSources(env), []);
  assert.equal(forwardUserMcp(env), false);
  assert.equal(forwardUserMcp({ ...env, DRAFTSPECT_AGENT_ACCESS: "extended" }), true);
  assert.deepEqual(userSettingsSources({ DRAFTSPECT_AGENT_ACCESS: "extended" }), []);
  assert.throws(() => agentAccess({ DRAFTSPECT_AGENT_ACCESS: "typo" }), /PREVIEW_ACCESS_INVALID/);
});
test("account environment must be explicit, typed, and never echoed on error", () => {
  for (const value of [undefined, "", "https://secret.invalid/token", "EXTERNAL", {}])
    assert.throws(
      () => codebuddyEnvironment({ CODEBUDDY_INTERNET_ENVIRONMENT: value }),
      (e) => e.message.includes("PREVIEW_ENV_REQUIRED") && !e.message.includes("secret.invalid"),
    );
  for (const value of ["external", "internal", "ioa", "cloudhosted"])
    assert.equal(codebuddyEnvironment({ CODEBUDDY_INTERNET_ENVIRONMENT: value }), value);
});
test("SDK options eliminate ambient hooks, agents and external servers and fail closed on tool names", async () => {
  let called = 0;
  const options = previewSdkOptions(
    {
      mcpServers: { office: { type: "sdk" }, evil: {} },
      tools: ["Bash"],
      hooks: { SessionStart: [{}] },
      extraArgs: { dangerous: true },
      canUseTool: async () => {
        called++;
        return { behavior: "allow" };
      },
    },
    {},
  );
  assert.deepEqual(Object.keys(options.mcpServers), ["office"]);
  assert.deepEqual(options.tools, []);
  assert.deepEqual(options.hooks, {});
  assert.deepEqual(options.agents, {});
  assert.deepEqual(options.extraArgs, {});
  for (const name of [
    "Bash",
    "Read",
    "Write",
    "Skill",
    "mcp__other__office_find",
    "office_find",
    "mcp__office__office_find;Bash",
    "mcp__office__office_find\n",
  ])
    assert.equal((await options.canUseTool(name, {})).behavior, "deny");
  for (const name of ["mcp__office__office_find", "mcp__office__excel_read_range"])
    assert.equal((await options.canUseTool(name, {})).behavior, "allow");
  assert.equal(called, 2);
  assert.equal(
    (await previewSdkOptions({}, {}).canUseTool("mcp__office__office_find", {})).behavior,
    "deny",
  );
});
test(
  "pinned real SDK serializes no-settings, no-builtins and strict MCP switches",
  { skip: !ProcessTransport && "CodeBuddy SDK absent; run npm ci for pinned SDK checks" },
  () => {
    const transport = new ProcessTransport(previewSdkOptions({ mcpServers: {} }, {}));
    const args = transport.buildArgs();
    assert.ok(args.includes("--strict-mcp-config"));
    assert.equal(args[args.indexOf("--setting-sources") + 1], "none");
    assert.equal(args[args.indexOf("--tools") + 1], "");
  },
);
test("preview state is independent and legacy source auto-migration is disabled", async () => {
  assert.equal(stateDirectory({}), join(homedir(), ".draftspect-preview"));
  for (const p of [join(homedir(), ".draftspect-cb"), join(homedir(), ".draftspect-cb", "child")])
    assert.throws(() => stateDirectory({ DRAFTSPECT_STATE_DIR: p }));
  const source = await readFile(new URL("../daemon/sessions.mjs", import.meta.url), "utf8");
  assert.match(source, /legacyFile = null/);
  const daemon = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  assert.match(daemon, /if \(acceptanceProfile \|\| !forwardUserMcp\(\)\) return \{\};/);
});
