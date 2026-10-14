// Tests for the graceful-shutdown / restart-reliability package:
//  1. createLocalServer.close() must release the port even when clients hold
//     keep-alive connections open (the 2026-09-20 EADDRINUSE restart loop).
//  2. Tray restart uses exponential backoff (app/restart-backoff.mjs).
//  3. Timestamped console install (daemon/log.mjs).
//  4. A second daemon on the same ports fails with a clear, actionable message.
//  5. E2E regression guard: SIGTERM to a real spawned daemon releases both
//     ports quickly enough for an immediate restart.

import { test } from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import { mkdtemp, mkdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import net from "node:net";
import { spawn } from "node:child_process";

import { createLocalServer } from "../daemon/local-transport.mjs";
import { restartDelayMs } from "../app/restart-backoff.mjs";
import { installTimestampedConsole } from "../daemon/log.mjs";

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

test(
  "local server close() releases the port despite an in-flight request",
  { timeout: 5000 },
  async (t) => {
    let handlerCalls = 0;
    let entered;
    const handlerEntered = new Promise((resolve) => {
      entered = resolve;
    });
    const transport = createLocalServer({
      port: 0,
      handler: () => {
        handlerCalls++;
        entered();
        /* never responds — simulates a stalled asset request during shutdown */
      },
    });
    const address = await transport.ready;
    const held = net.connect(address.port, "127.0.0.1");
    t.after(async () => {
      held.destroy();
      await transport.close();
    });
    await once(held, "connect");
    held.write(`GET /word/index.html HTTP/1.1\r\nHost: 127.0.0.1:${address.port}\r\n\r\n`);
    try {
      await handlerEntered;
      assert.equal(handlerCalls, 1, "the stalled handler must actually be entered");
      const closed = await Promise.race([
        transport.close().then(() => "closed"),
        delay(2000).then(() => "hang"),
      ]);
      assert.equal(closed, "closed", "close() must not wait for in-flight requests");
      const rebind = net.createServer();
      await Promise.race([
        once(rebind.listen(address.port, "127.0.0.1"), "listening"),
        delay(1500).then(() => {
          throw new Error("port was not released for rebind");
        }),
      ]);
      rebind.close();
    } finally {
      held.destroy();
      await transport.close();
    }
  },
);

test("restart backoff doubles with a cap", () => {
  assert.equal(restartDelayMs(1), 1000);
  assert.equal(restartDelayMs(2), 2000);
  assert.equal(restartDelayMs(3), 4000);
  assert.equal(restartDelayMs(4), 8000);
  assert.equal(restartDelayMs(9), 8000, "capped");
});

test("installTimestampedConsole prefixes ISO timestamps and is idempotent", () => {
  const calls = [];
  const fake = { log: (...args) => calls.push(args) };
  installTimestampedConsole(fake);
  const again = installTimestampedConsole(fake);
  fake.log("hello", 42);
  assert.equal(calls.length, 1, "no double-wrapping");
  const [stamp, text, num] = calls[0];
  assert.ok(!Number.isNaN(Date.parse(stamp)), "first arg is an ISO timestamp");
  assert.equal(text, "hello");
  assert.equal(num, 42);
  assert.equal(again, fake);
});

// ---------------------------------------------------------------------------
// Real-daemon E2E (offline): isolated HOME/state/ports, claude runtime so no
// CLI subprocess or provider contact happens; no user message is ever sent.
// ---------------------------------------------------------------------------

const DAEMON = new URL("../daemon/index.mjs", import.meta.url).pathname;

async function spawnDaemon(env, workspace) {
  const proc = spawn(process.execPath, [DAEMON, workspace], {
    env: {
      ...process.env,
      HOME: env.home,
      DRAFTSPECT_RUNTIME: "claude",
      DRAFTSPECT_STATE_DIR: env.state,
      DRAFTSPECT_HTTP_PORT: String(env.httpPort),
      DRAFTSPECT_WS_PORT: String(env.wsPort),
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  env.children.add(proc);
  let output = "";
  proc.stdout.on("data", (d) => (output += d));
  proc.stderr.on("data", (d) => (output += d));
  const waitFor = (needle, timeoutMs = 20000) =>
    new Promise((resolve, reject) => {
      const deadline = Date.now() + timeoutMs;
      const poll = () => {
        if (output.includes(needle)) return resolve(output);
        if (proc.exitCode !== null || proc.signalCode !== null)
          return reject(
            new Error(`daemon exited (${proc.exitCode}) before "${needle}":\n${output}`),
          );
        if (Date.now() > deadline)
          return reject(new Error(`timed out waiting for "${needle}":\n${output}`));
        setTimeout(poll, 50);
      };
      poll();
    });
  return { proc, waitFor, output: () => output };
}

async function stopOwnedDaemon(proc) {
  if (proc.exitCode !== null || proc.signalCode !== null) return;
  const exited = once(proc, "exit");
  const force = setTimeout(() => proc.kill("SIGKILL"), 3000);
  proc.kill("SIGTERM");
  try {
    await exited;
  } finally {
    clearTimeout(force);
  }
}

async function daemonFixture(t) {
  const root = await mkdtemp(join(tmpdir(), "draftspect-shutdown-"));
  const home = join(root, "home");
  const state = join(root, "state");
  const workspace = join(root, "workspace");
  await mkdir(home, { recursive: true });
  await mkdir(state, { recursive: true });
  await mkdir(workspace, { recursive: true });
  const children = new Set();
  t.after(async () => {
    await Promise.all([...children].map(stopOwnedDaemon));
    await rm(root, { recursive: true, force: true });
  });
  return { home, state, workspace, children };
}

async function allocatePortPair() {
  // Hold both kernel-assigned ports together, then release immediately before
  // spawning. Another process can win that handoff race; startup retries below
  // are bounded and only apply to an actual address-in-use failure.
  const reservations = [];
  try {
    for (let i = 0; i < 2; i++) {
      const server = net.createServer();
      reservations.push(server);
      await once(server.listen(0, "127.0.0.1"), "listening");
    }
    const [httpPort, wsPort] = reservations.map((server) => server.address().port);
    assert.notEqual(httpPort, wsPort);
    assert.ok(![httpPort, wsPort].some((port) => [47833, 47834].includes(port)));
    return { httpPort, wsPort };
  } finally {
    await Promise.all(
      reservations
        .filter((server) => server.listening)
        .map(
          (server) =>
            new Promise((resolve, reject) =>
              server.close((error) => (error ? reject(error) : resolve())),
            ),
        ),
    );
  }
}

async function startIsolatedDaemon(fixture) {
  for (let attempt = 1; attempt <= 3; attempt++) {
    const ports = await allocatePortPair();
    const daemon = await spawnDaemon({ ...fixture, ...ports }, fixture.workspace);
    try {
      await daemon.waitFor("Ready; waiting for a taskpane");
      return { ...daemon, ...ports };
    } catch (error) {
      await stopOwnedDaemon(daemon.proc);
      if (attempt === 3 || !/EADDRINUSE|is already in use/i.test(daemon.output())) throw error;
    }
  }
}

async function portFree(port) {
  const probe = net.createServer();
  try {
    await once(probe.listen(port, "127.0.0.1"), "listening");
    return true;
  } catch {
    return false;
  } finally {
    probe.close();
  }
}

test(
  "SIGTERM releases both ports fast enough for an immediate restart",
  { timeout: 60000 },
  async (t) => {
    const fixture = await daemonFixture(t);
    const daemon = await startIsolatedDaemon(fixture);
    const { httpPort, wsPort } = daemon;
    assert.equal(await portFree(wsPort), false, "WS port is bound while running");
    assert.equal(await portFree(httpPort), false, "HTTP port is bound while running");

    daemon.proc.kill("SIGTERM");
    await once(daemon.proc, "exit");

    const deadline = Date.now() + 5000;
    let released = false;
    while (Date.now() < deadline) {
      if ((await portFree(wsPort)) && (await portFree(httpPort))) {
        released = true;
        break;
      }
      await delay(100);
    }
    assert.ok(released, "both ports must be rebindable within 5s of SIGTERM");
  },
);

test(
  "a second daemon on the same ports exits with a clear EADDRINUSE message",
  { timeout: 60000 },
  async (t) => {
    const fixture = await daemonFixture(t);
    const first = await startIsolatedDaemon(fixture);
    const { httpPort, wsPort } = first;

    // Deliberate conflict: bypass the allocation/retry helper for this child.
    const second = await spawnDaemon({ ...fixture, httpPort, wsPort }, fixture.workspace);
    await once(second.proc, "exit");
    assert.notEqual(second.proc.exitCode, 0);
    assert.match(
      second.output(),
      /already in use|already running/i,
      "port conflicts must produce an actionable message, not only a stack trace",
    );
    assert.equal(first.proc.exitCode, null, "the conflicting child must not stop the first");
    assert.equal(first.proc.signalCode, null);
  },
);

test(
  "independent daemon fixtures stay ready concurrently on distinct ports",
  { timeout: 60000 },
  async (t) => {
    const fixtures = await Promise.all([daemonFixture(t), daemonFixture(t)]);
    const daemons = await Promise.all(fixtures.map(startIsolatedDaemon));
    const ports = daemons.flatMap(({ httpPort, wsPort }) => [httpPort, wsPort]);
    assert.equal(new Set(ports).size, 4, "parallel fixtures must own four distinct ports");
    for (const daemon of daemons) {
      assert.equal(daemon.proc.exitCode, null);
      assert.equal(daemon.proc.signalCode, null);
    }
    for (const port of ports) assert.equal(await portFree(port), false, "each ready port is bound");
    t.diagnostic(JSON.stringify({ scenario: "parallel-isolated-daemons", ports }));
  },
);
