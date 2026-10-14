import test from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { forwardUserMcp } from "../daemon/preview-policy.mjs";
import { runSession as codebuddy } from "../daemon/runtime/runtime-codebuddy.mjs";
import { runSession as claude } from "../daemon/runtime/runtime-claude.mjs";

test("actual MCP loader does not read user configuration by default; explicit extended opt-in reads only the fixture", async () => {
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  const loader = source.slice(
    source.indexOf("async function loadUserMcpServers()"),
    source.indexOf("// Hoisted ABOVE the top-level awaits"),
  );
  for (const [env, expected] of [
    [{}, 0],
    [{ DRAFTSPECT_FORWARD_USER_MCP: "1" }, 0],
    [{ DRAFTSPECT_AGENT_ACCESS: "extended" }, 0],
    [{ DRAFTSPECT_AGENT_ACCESS: "extended", DRAFTSPECT_FORWARD_USER_MCP: "1" }, 1],
  ]) {
    let reads = 0;
    const ctx = vm.createContext({
      acceptanceProfile: null,
      forwardUserMcp: () => forwardUserMcp(env),
      join,
      homedir: () => "/synthetic",
      readFile: async (path) => {
        assert.equal(path, "/synthetic/.claude.json");
        reads++;
        return '{"mcpServers":{"synthetic":{}}}';
      },
    });
    vm.runInContext(loader + "\nglobalThis.load=loadUserMcpServers;", ctx);
    const result = await ctx.load();
    assert.equal(reads, expected);
    assert.equal(Object.keys(result).length, expected);
  }
});

const turn = {
  cwd: "/synthetic",
  mcpServers: { office: { synthetic: true }, untrusted: {} },
  canUseTool: async () => ({ behavior: "allow" }),
  turnSource: { next: async () => null },
  onMessage() {},
};
function assertPolicy(options) {
  assert.deepEqual(options.tools, []);
  assert.deepEqual(options.settingSources, []);
  assert.deepEqual(options.hooks, {});
  assert.equal(options.strictMcpConfig, true);
  assert.deepEqual(Object.keys(options.mcpServers), ["office"]);
}
test("both actual runtime entry points pass the preview policy to injected transports without provider execution", async () => {
  process.env.CODEBUDDY_INTERNET_ENVIRONMENT = "external";
  process.env.DRAFTSPECT_AGENT_ACCESS = "office-only";
  let cb = 0,
    cl = 0;
  const c = codebuddy(
    { ...turn, onEnd: (e) => assert.equal(e.reason, "ended") },
    {
      create: (options) => {
        assertPolicy(options);
        cb++;
        return { connect: async () => {}, close() {} };
      },
      observeClose: (s) => ({ close: () => s.close(), waitForExit: async () => {} }),
    },
  );
  await c.done;
  const a = claude(
    { ...turn, onEnd: (e) => assert.equal(e.reason, "ended") },
    {
      querySession: ({ options }) => {
        assertPolicy(options);
        cl++;
        return (async function* () {})();
      },
    },
  );
  await a.done;
  assert.equal(cb, 1);
  assert.equal(cl, 1);
});
test("invalid preview configuration terminates both runtimes before transport construction", async () => {
  process.env.DRAFTSPECT_AGENT_ACCESS = "typo";
  let constructed = 0;
  const errors = [];
  const onEnd = (e) => {
    assert.equal(e.reason, "error");
    errors.push(e.error.message);
  };
  const cb = codebuddy(
    { ...turn, onEnd },
    {
      create: () => {
        constructed++;
        throw new Error("must not run");
      },
    },
  );
  const cl = claude(
    { ...turn, onEnd },
    {
      querySession: () => {
        constructed++;
        throw new Error("must not run");
      },
    },
  );
  await Promise.all([cb.done, cl.done]);
  assert.equal(constructed, 0);
  assert.equal(errors.length, 2);
  assert.ok(errors.every((e) => e.includes("PREVIEW_ACCESS_INVALID")));
  delete process.env.DRAFTSPECT_AGENT_ACCESS;
});
