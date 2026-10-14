// Opt-in actual daemon shutdown code + real control-only SDK child, no generation.
// The child is deliberately absent from sessions, but retained in the real close registry.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createServer } from "node:net";
import { createRequire } from "node:module";
import { resolveLocalTransport } from "../daemon/local-transport.mjs";
const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const out = resolve(process.argv[2]);
await mkdir(out, { recursive: true, mode: 0o700 });
const state = join(out, "state"),
  workspace = join(out, "synthetic-workspace");
await mkdir(state, { mode: 0o700 });
await mkdir(workspace, { mode: 0o700 });
async function port() {
  const s = createServer();
  await new Promise((r) => s.listen(0, "127.0.0.1", r));
  const p = s.address().port;
  await new Promise((r) => s.close(r));
  return p;
}
const httpPort = await port(),
  wsPort = await port(),
  tls = await resolveLocalTransport({});
const original = await readFile(join(root, "daemon/index.mjs"), "utf8");
const rewritten = original
  .replace(
    /from "(\.\.?\/[^"\n]+)"/g,
    (_, p) => "from " + JSON.stringify(pathToFileURL(resolve(root, "daemon", p)).href),
  )
  .replace(
    "const __dirname = dirname(fileURLToPath(import.meta.url));",
    "const __dirname = " + JSON.stringify(join(root, "daemon")) + ";",
  );
const injected = `
const {unstable_v2_createSession:qaCreate}=await import(${JSON.stringify(pathToFileURL(createRequire(import.meta.url).resolve("@tencent-ai/agent-sdk")).href)});
const {observeCodeBuddyClose:qaObserve}=await import(${JSON.stringify(pathToFileURL(join(root, "daemon/runtime/codebuddy-process-close.mjs")).href)});
const {writeFileSync:qaWrite}=await import('node:fs');
const qaHandle=runtime.runSession({cwd:${JSON.stringify(workspace)},mcpServers:{},turnSource:{next:()=>new Promise(()=>{})},onMessage:()=>{throw new Error('No generated stream expected')},onEnd:()=>{}},{
 create:(options)=>{
  const s=qaCreate({...options,tools:[],hooks:{},settingSources:[],mcpServers:{},canUseTool:async()=>({behavior:'deny',message:'No tools in control-only QA'})});
  s.send=()=>{throw new Error('NO_GENERATION_ALLOWED')};
  const connect=s.connect.bind(s);s.connect=async()=>{await connect();process.send({type:'qa_connected'});};return s;
 },
 observeClose:(s)=>qaObserve(s,{onProcess:(child)=>{process.send({type:'qa_child',pid:child.pid});child.once('exit',(code,signal)=>qaWrite(${JSON.stringify(join(out, "sdk-exit.json"))},JSON.stringify({pid:child.pid,code,signal,time:new Date().toISOString()}),{flag:'wx',mode:0o600}));}})
});
runtimeClosures.track('synthetic-removed-session',qaHandle.closed,qaHandle);
`;
const entry = join(out, "instrumented-daemon.mjs");
await writeFile(entry, rewritten + injected, { flag: "wx", mode: 0o600 });
const child = spawn(process.execPath, [entry], {
  cwd: workspace,
  stdio: ["ignore", "ignore", "pipe", "ipc"],
  env: {
    HOME: process.env.HOME,
    PATH: process.env.PATH,
    TMPDIR: process.env.TMPDIR,
    LANG: process.env.LANG,
    DRAFTSPECT_STATE_DIR: state,
    DRAFTSPECT_RUNTIME: "codebuddy",
    DRAFTSPECT_HTTP_PORT: String(httpPort),
    DRAFTSPECT_WS_PORT: String(wsPort),
    CODEBUDDY_CODE_PATH: process.env.CODEBUDDY_CODE_PATH,
    DRAFTSPECT_TLS_KEY: tls.keyPath,
    DRAFTSPECT_TLS_CERT: tls.certPath,
  },
});
let sdkPid,
  connected = false,
  errors = "",
  outcome;
child.stderr.on("data", (d) => (errors += d));
child.on("message", (m) => {
  if (m.type === "qa_child") sdkPid = m.pid;
  if (m.type === "qa_connected") connected = true;
});
const exited = new Promise((r) =>
  child.once("exit", (code, signal) => {
    outcome = { code, signal };
    r(outcome);
  }),
);
let failure;
try {
  const deadline = Date.now() + 50000;
  while (!connected && !outcome && Date.now() < deadline)
    await new Promise((r) => setTimeout(r, 50));
  assert.ok(
    connected && Number.isInteger(sdkPid),
    "Real control connection required: " + errors.slice(0, 200),
  );
  child.disconnect();
  const timeout = new Promise((_, rej) => {
    const t = setTimeout(() => rej(new Error("Daemon parent-loss timeout")), 5000);
    t.unref();
  });
  assert.deepEqual(await Promise.race([exited, timeout]), { code: 0, signal: null });
  const exit = JSON.parse(await readFile(join(out, "sdk-exit.json"), "utf8"));
  assert.equal(exit.pid, sdkPid);
  assert.throws(() => process.kill(sdkPid, 0), { code: "ESRCH" });
} catch (e) {
  failure = e.message;
  process.exitCode = 1;
} finally {
  if (!outcome) child.kill("SIGTERM");
  await exited;
}
await writeFile(
  join(out, "result.json"),
  JSON.stringify(
    {
      status: failure ? "FAIL" : "PASS",
      scope: "REAL_DAEMON_IPC_DISCONNECT_WITH_REAL_REMOVED_SDK_CHILD",
      daemonPid: child.pid,
      sdkPid,
      connected,
      outcome,
      modelRequests: 0,
      ...(failure ? { failure } : {}),
      stderrPresent: !!errors,
    },
    null,
    2,
  ),
  { flag: "wx", mode: 0o600 },
);
console.log(
  JSON.stringify({
    status: failure ? "FAIL" : "PASS",
    daemonPid: child.pid,
    sdkPid,
    outcome,
    failure,
  }),
);
