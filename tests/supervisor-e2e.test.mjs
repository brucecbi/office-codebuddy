// Offline E2E: actual tray supervisor functions + actual daemon children.
// No Electron UI, Word, user messages, provider calls or production state/ports.
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { mkdtemp, mkdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import net from "node:net";
import WebSocket from "ws";
import { trayHarness } from "./helpers/tray-supervisor.mjs";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function waitFor(check, detail) {
  const deadline = Date.now() + 10000;
  while (!check()) {
    if (Date.now() >= deadline) throw new Error(`Timeout: ${detail}`);
    await delay(20);
  }
}
async function ephemeralPorts() {
  const servers = [net.createServer(), net.createServer()];
  try {
    await Promise.all(servers.map((server) => once(server.listen(0, "127.0.0.1"), "listening")));
    return servers.map((server) => server.address().port);
  } finally {
    await Promise.all(servers.map((server) => new Promise((resolve) => server.close(resolve))));
  }
}

test(
  "offline E2E: one manual restart owns one real replacement, then Stop releases both ports",
  { timeout: 30000 },
  async (t) => {
    const root = await mkdtemp(join(tmpdir(), "draftspect-supervisor-e2e-"));
    const fixtureHome = join(root, "home"),
      state = join(root, "state"),
      workspace = join(root, "workspace");
    await Promise.all([fixtureHome, state, workspace].map((path) => mkdir(path)));
    const [httpPort, wsPort] = await ephemeralPorts();
    const repo = fileURLToPath(new URL("../", import.meta.url));
    const outputs = new Map();
    let ws, h;
    t.after(async () => {
      ws?.terminate();
      h?.stop();
      for (const child of h?.children ?? []) {
        if (child.exitCode === null && child.signalCode === null) {
          const closed = once(child, "close");
          child.kill("SIGKILL"); // only this test's captured child, on failed cleanup
          await closed;
        }
      }
      await rm(root, { recursive: true, force: true });
    });
    h = await trayHarness({
      findWorkspace: async () => workspace,
      spawnChild(executable, args, options) {
        const child = spawn(executable, args, options);
        outputs.set(child, "");
        const collect = (chunk) => outputs.set(child, outputs.get(child) + chunk.toString());
        child.stdout.on("data", collect);
        child.stderr.on("data", collect);
        return child;
      },
      overrides: {
        process: {
          execPath: process.execPath,
          env: {
            PATH: process.env.PATH,
            HOME: fixtureHome,
            DRAFTSPECT_RUNTIME: "claude", // static catalog; no eager model session
            DRAFTSPECT_STATE_DIR: state,
            DRAFTSPECT_HTTP_PORT: String(httpPort),
            DRAFTSPECT_WS_PORT: String(wsPort),
          },
        },
        PROJECT_ROOT: repo,
        DAEMON_ENTRY: join(repo, "daemon/index.mjs"),
        setTimeout,
        clearTimeout,
        setInterval,
        clearInterval,
      },
    });
    const ready = async (child) => {
      await waitFor(() => {
        if (child.exitCode !== null || child.signalCode !== null)
          throw new Error(`Synthetic daemon exited before ready: ${outputs.get(child)}`);
        return outputs.get(child).includes("Ready; waiting for a taskpane");
      }, "daemon ready");
    };
    await h.start();
    const first = h.owner();
    await ready(first);
    const firstClosed = once(first, "close");
    h.restart();
    const [code, signal] = await firstClosed;
    assert.equal(code, 0);
    assert.equal(signal, null, "handled SIGTERM is a clean exit, not a signal exit");
    await waitFor(() => h.children.length >= 2, "replacement creation");
    const replacement = h.children[1];
    await ready(replacement);
    const token = (await readFile(join(state, "bridge-token"), "utf8")).trim();
    ws = new WebSocket(`ws://127.0.0.1:${wsPort}`, { origin: `http://127.0.0.1:${httpPort}` });
    await once(ws, "open");
    const welcome = once(ws, "message");
    ws.send(
      JSON.stringify({
        type: "hello",
        token,
        host: "word",
        active_doc: join(workspace, "synthetic.docx"),
        pane_id: "synthetic-supervisor",
        protocol_version: 2,
      }),
    );
    assert.equal(JSON.parse(String((await welcome)[0])).type, "welcome");
    await delay(1250); // exceed the old duplicate-start backoff, using real time
    assert.equal(h.children.length, 2, "an extra daemon was launched after manual restart");
    assert.equal(h.owner(), replacement);
    assert.equal(h.status(), "running");
    assert.doesNotMatch(h.logs.join("\n"), /auto-restarting|EADDRINUSE/);
    const replacementClosed = once(replacement, "close");
    h.stop();
    assert.deepEqual(await replacementClosed, [0, null]);
    assert.equal(h.owner(), null);
    assert.equal(h.status(), "stopped");
    for (const port of [httpPort, wsPort]) {
      const server = net.createServer();
      try {
        await once(server.listen(port, "127.0.0.1"), "listening");
      } finally {
        await new Promise((resolve) => server.close(resolve));
      }
    }
  },
);
