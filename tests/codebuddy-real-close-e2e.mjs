// Opt-in real installed SDK/CLI control connection. Never calls send(), tools,
// user hooks or extra MCP servers; only synthetic cwd and model metadata.
import assert from "node:assert/strict";
import { mkdir, writeFile } from "node:fs/promises";
import { resolve, join } from "node:path";
import { unstable_v2_createSession } from "@tencent-ai/agent-sdk";
import { observeCodeBuddyClose } from "../daemon/runtime/codebuddy-process-close.mjs";
import { createModelDiscovery } from "../daemon/runtime/model-discovery.mjs";
import { runSession, resolveCliPath } from "../daemon/runtime/runtime-codebuddy.mjs";
import { codebuddyEnvironment } from "../daemon/preview-policy.mjs";

if (process.env.DRAFTSPECT_ALLOW_PROVIDER_CONTROL_TEST !== "1")
  throw new Error(
    "Explicit provider-control test authorization required; set DRAFTSPECT_ALLOW_PROVIDER_CONTROL_TEST=1 after reviewing the test.",
  );
const accountEnvironment = codebuddyEnvironment();
const out = resolve(process.argv[2]);
await mkdir(out, { recursive: true, mode: 0o700 });
const cwd = join(out, "synthetic-workspace");
await mkdir(cwd, { mode: 0o700 });
const children = [],
  checks = [];
let generationAttempts = 0,
  controls = 0,
  runtime;
const options = {
  cwd,
  environment: accountEnvironment,
  pathToCodebuddyCode: resolveCliPath(),
  settingSources: [],
  hooks: {},
  tools: [],
  mcpServers: {},
  strictMcpConfig: true,
  canUseTool: async () => ({ behavior: "deny", message: "Synthetic control-only test" }),
};
function create(extra) {
  const session = unstable_v2_createSession({ ...extra, ...options });
  session.send = () => {
    generationAttempts++;
    throw new Error("NO_MODEL_GENERATION_ALLOWED");
  };
  const connect = session.connect.bind(session);
  session.connect = async () => {
    await connect();
    controls++;
  };
  return session;
}
function observe(session) {
  return observeCodeBuddyClose(session, { onProcess: (child) => children.push(child) });
}
async function until(check, timeout = 45000) {
  const end = Date.now() + timeout;
  while (!check()) {
    if (Date.now() >= end) throw new Error("Bounded real SDK control timeout");
    await new Promise((r) => setTimeout(r, 50));
  }
}
const discovery = createModelDiscovery({ create, observeClose: observe, options: () => options });
let error;
try {
  let ending;
  runtime = runSession(
    {
      cwd,
      mcpServers: {},
      turnSource: { next: () => new Promise(() => {}) },
      onMessage: () => {
        throw new Error("No generated stream expected");
      },
      onEnd: (result) => {
        ending = result.reason;
      },
    },
    { create, observeClose: observe, connectTimeoutMs: 40000 },
  );
  await until(() => controls === 1 || ending);
  assert.equal(controls, 1, "real SDK control connection established");
  assert.equal(runtime.getCloseStatus().phase, "running");
  const pid = runtime.getCloseStatus().pid;
  runtime.close();
  await Promise.race([
    runtime.closed,
    new Promise((_, reject) => {
      const t = setTimeout(() => reject(new Error("Runtime close timeout")), 5000);
      t.unref();
    }),
  ]);
  assert.equal(runtime.getCloseStatus().phase, "exited");
  checks.push({
    name: "real-session-connect-close",
    status: "PASS",
    pid,
    exit: runtime.getCloseStatus().exit,
    ending,
  });
  const models = await discovery.list({ cwd, timeoutMs: 40000 });
  assert.ok(Array.isArray(models) && models.length, "real catalog returned metadata");
  await discovery.shutdown(5000);
  checks.push({ name: "real-catalog-discovery-close", status: "PASS", modelCount: models.length });
  assert.equal(generationAttempts, 0);
  assert.ok(children.every((c) => c.exitCode !== null || c.signalCode !== null));
} catch (e) {
  error = { name: e.name, message: String(e.message).slice(0, 300) };
} finally {
  runtime?.close();
  await discovery.shutdown(5000).catch(() => {});
  // Exact captured handles only; never find/kill unrelated processes.
  for (const child of children)
    if (child.exitCode === null && child.signalCode === null) child.kill("SIGTERM");
}
await writeFile(
  join(out, "result.json"),
  JSON.stringify(
    {
      status: error ? "FAIL" : "PASS",
      scope: "REAL_SDK_CLI_CONTROL_NO_GENERATION",
      controls,
      generationAttempts,
      checks,
      children: children.map((c) => ({ pid: c.pid, code: c.exitCode, signal: c.signalCode })),
      ...(error ? { error } : {}),
    },
    null,
    2,
  ),
  { flag: "wx", mode: 0o600 },
);
console.log(
  JSON.stringify({ status: error ? "FAIL" : "PASS", controls, generationAttempts, checks }),
);
if (error) process.exitCode = 1;
