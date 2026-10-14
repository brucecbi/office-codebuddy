import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const repo = fileURLToPath(new URL("../", import.meta.url));
const run = (env) =>
  promisify(execFile)(
    process.execPath,
    [
      "--input-type=module",
      "-e",
      "import runtime from './daemon/runtime/index.mjs'; console.log(runtime.id);",
    ],
    { cwd: repo, env: { PATH: process.env.PATH, ...env } },
  );

test("an unset runtime selects CodeBuddy and an unknown value fails", async () => {
  const unset = await run({});
  assert.equal(unset.stdout.trim(), "codebuddy");
  const named = await run({ DRAFTSPECT_RUNTIME: "codebuddy" });
  assert.equal(named.stdout.trim(), "codebuddy");
  await assert.rejects(run({ DRAFTSPECT_RUNTIME: "no-such-runtime" }), (error) => {
    assert.match(error.stderr, /Unknown DRAFTSPECT_RUNTIME="no-such-runtime"/);
    return true;
  });
});

test("DRAFTSPECT_RUNTIME=claude still selects the Claude runtime", async () => {
  const selected = await run({ DRAFTSPECT_RUNTIME: "claude" });
  assert.equal(selected.stdout.trim(), "claude");
  assert.equal(selected.stderr, "");
});

test("CodeBuddy runtime plus Office registration load without any Anthropic import", async () => {
  const loader = `export async function resolve(specifier, context, next) { if (specifier.includes('@anthropic-ai/')) throw new Error('Unselected SDK imported'); return next(specifier, context); }`;
  const script = `import { register } from 'node:module'; register('data:text/javascript,'+encodeURIComponent(${JSON.stringify(loader)}),import.meta.url); const runtime=(await import('./daemon/runtime/index.mjs')).default; const {createOfficeBridgeMcp}=await import('./daemon/office-tools.mjs'); const server=createOfficeBridgeMcp({callTaskpaneTool:async()=>({})},'word','synthetic',runtime.mcp); console.log(JSON.stringify({id:runtime.id,tools:Object.keys(server.instance._registeredTools).length}));`;
  const result = await promisify(execFile)(
    process.execPath,
    ["--input-type=module", "-e", script],
    { cwd: repo, env: { PATH: process.env.PATH, DRAFTSPECT_RUNTIME: "codebuddy" } },
  );
  assert.deepEqual(JSON.parse(result.stdout.trim().split("\n").at(-1)), {
    id: "codebuddy",
    tools: 34,
  });
});
