// Synthetic installed-CLI history replay. Never sends a user/model prompt.
// macOS only: requires sandbox-exec and refuses to start the SDK until live
// network/filesystem isolation probes pass. --self-loopback permits only the
// CLI's explicitly pinned synthetic localhost port; all other network stays denied.
// Run: node scripts/session-replay-offline.mjs
// Codex's outer sandbox may require elevated execution to apply this stricter
// child sandbox. No HOME/CODEX_HOME override, credentials or real sessions.
import assert from "node:assert/strict";
import childProcess from "node:child_process";
import { createHash, randomUUID } from "node:crypto";
import { readFile, writeFile, mkdir, mkdtemp, realpath, stat } from "node:fs/promises";
import net from "node:net";
import { dirname, join, delimiter } from "node:path";
import { fileURLToPath } from "node:url";
import { syncBuiltinESMExports } from "node:module";

const self = fileURLToPath(import.meta.url);
const repo = dirname(dirname(self));
const sha = (bytes) => createHash("sha256").update(bytes).digest("hex");
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const json = (value) => process.stdout.write(JSON.stringify(value) + "\n");

async function probe() {
  const [readable, forbidden, port, selfPort] = process.argv.slice(3);
  assert.equal(await readFile(readable, "utf8"), "Synthetic isolation fixture");
  let fileDenial;
  try {
    await readFile(forbidden);
  } catch (error) {
    fileDenial = error.code;
  }
  assert.ok(["EPERM", "EACCES"].includes(fileDenial), "File isolation was not proved");
  const networkDenial = await new Promise((resolve, reject) => {
    const socket = net.connect({ host: "127.0.0.1", port: Number(port) });
    socket.once("connect", () => {
      socket.destroy();
      reject(new Error("Sandbox unexpectedly allowed a loopback connection"));
    });
    socket.once("error", (error) => resolve(error.code));
    socket.setTimeout(3000, () => {
      socket.destroy();
      reject(new Error("Network denial probe timed out"));
    });
  });
  assert.ok(["EPERM", "EACCES"].includes(networkDenial), "No OS network denial was proved");
  let externalDenial = null,
    allowedSelfConnect = false,
    otherPortBindDenial = null,
    actualSelfAddress = null;
  if (selfPort) {
    // TEST-NET-1 is reserved documentation space, never a business endpoint.
    externalDenial = await new Promise((resolve, reject) => {
      const socket = net.connect({ host: "192.0.2.1", port: 443 });
      socket.once("connect", () => {
        socket.destroy();
        reject(new Error("External connection unexpectedly allowed"));
      });
      socket.once("error", (error) => resolve(error.code));
      socket.setTimeout(2000, () => {
        socket.destroy();
        reject(new Error("External denial was not proved"));
      });
    });
    assert.ok(["EPERM", "EACCES"].includes(externalDenial), "No OS external denial was proved");
    const selfServer = net.createServer((socket) => socket.end());
    await new Promise((resolve, reject) => {
      selfServer.once("error", reject);
      selfServer.listen(Number(selfPort), "127.0.0.1", resolve);
    });
    try {
      actualSelfAddress = selfServer.address().address;
      assert.equal(
        actualSelfAddress,
        "127.0.0.1",
        "Probe did not bind the explicit IPv4 loopback address",
      );
      await new Promise((resolve, reject) => {
        const socket = net.connect({ host: "127.0.0.1", port: Number(selfPort) });
        socket.once("connect", () => {
          allowedSelfConnect = true;
          socket.destroy();
          resolve();
        });
        socket.once("error", reject);
      });
    } finally {
      await new Promise((resolve) => selfServer.close(resolve));
    }
    otherPortBindDenial = await new Promise((resolve, reject) => {
      const server = net.createServer();
      server.once("error", (error) => resolve(error.code));
      server.listen(Number(port), "127.0.0.1", () =>
        server.close(() => reject(new Error("Another port binding unexpectedly allowed"))),
      );
    });
    assert.ok(
      ["EPERM", "EACCES"].includes(otherPortBindDenial),
      "Other-port bind denial was not proved",
    );
  }
  json({
    phase: "isolation",
    ok: true,
    networkDenial,
    fileDenial,
    externalDenial,
    allowedSelfConnect,
    otherPortBindDenial,
    actualSelfAddress,
  });
}

async function replay() {
  const [cli, fixtureCwd, fixtureConfig, sessionId, transcript, selfPort] = process.argv.slice(3);
  assert.equal(process.env.CODEBUDDY_CONFIG_DIR, fixtureConfig);
  const before = await readFile(transcript);
  const children = [];
  let childStderr = "";
  const originalSpawn = childProcess.spawn;
  // Observe only SDK-owned children so close evidence is an OS exit event,
  // not the SDK's synchronous close() return. No real daemon is involved.
  childProcess.spawn = function (command, args, options) {
    assert.equal(command, process.execPath, "Unexpected SDK executable");
    assert.ok(args.includes(join(dirname(dirname(cli)), "dist", "codebuddy-headless.js")));
    const child = originalSpawn.call(this, command, args, options);
    const entry = { pid: child.pid, exited: false };
    json({ phase: "cli_child_started", pid: child.pid });
    child.stderr?.on("data", (chunk) => {
      if (childStderr.length < 65536) childStderr += chunk;
    });
    child.once("exit", (code, signal) => Object.assign(entry, { exited: true, code, signal }));
    children.push(entry);
    return child;
  };
  syncBuiltinESMExports();
  const { unstable_v2_resumeSession } = await import("@tencent-ai/agent-sdk");
  const session = unstable_v2_resumeSession(sessionId, {
    cwd: fixtureCwd,
    environment: "external",
    pathToCodebuddyCode: cli,
    executable: process.execPath,
    persistSession: false,
    settingSources: [],
    strictMcpConfig: true,
    mcpServers: {},
    tools: [],
    ...(selfPort ? { extraArgs: { port: selfPort, host: "127.0.0.1" } } : {}),
    canUseTool: async () => ({ behavior: "deny", message: "Read-only synthetic replay" }),
  });
  // A guard makes an accidental later refactor to send() fail before dispatch.
  session.send = () => {
    throw new Error("Forbidden: this test never sends a prompt");
  };
  const events = [];
  let failure;
  try {
    for await (const event of session.stream()) {
      events.push(event);
      if (events.length > 30) throw new Error("Unexpected replay event count");
    }
  } catch (error) {
    failure = {
      name: error.name,
      code: error.code ?? null,
      authentication_required: /authentication|unauthorized|login required/i.test(error.message),
      timeout: /timeout|timed out/i.test(error.message),
      permission_denied: /EPERM|EACCES|permission denied|operation not permitted/i.test(
        error.message,
      ),
      module_resolution: /MODULE_NOT_FOUND|cannot find module/i.test(error.message),
    };
  } finally {
    session.close();
  }
  for (let n = 0; n < 50 && children.some((c) => !c.exited); n++) await delay(100);
  const text = events
    .flatMap((event) => event.message?.content ?? [])
    .map((block) => (typeof block === "string" ? block : (block.text ?? "")))
    .join("\n");
  const ids = [...new Set(events.map((event) => event.session_id).filter(Boolean))];
  const result = events.find((event) => event.type === "result");
  const evidence = {
    phase: "replay",
    ok:
      !failure &&
      ids.length === 1 &&
      ids[0] === sessionId &&
      text.includes("BRONZE-TULIP-42") &&
      text.includes("Synthetic archive summary") &&
      result?.is_error !== true &&
      children.length > 0 &&
      children.every((c) => c.exited),
    session_id: sessionId,
    observed_session_ids: ids,
    event_types: events.map((event) => event.type),
    synthetic_user_seen: text.includes("BRONZE-TULIP-42"),
    synthetic_assistant_seen: text.includes("Synthetic archive summary"),
    tool_call_seen: events.some((event) =>
      event.message?.content?.some?.((b) => b.type === "tool_use"),
    ),
    tool_result_seen: events.some((event) =>
      event.message?.content?.some?.((b) => b.type === "tool_result"),
    ),
    result_subtype: result?.subtype ?? null,
    sends: 0,
    failure: failure ?? null,
    child_stderr_diagnostic: {
      bytes: Buffer.byteLength(childStderr),
      permission_denied: /EPERM|EACCES|permission denied|operation not permitted/i.test(
        childStderr,
      ),
      module_resolution: /MODULE_NOT_FOUND|cannot find module/i.test(childStderr),
      sandbox_execute_denied: /spawn.*(?:EPERM|EACCES)/i.test(childStderr),
      error_first_lines: childStderr
        .split("\n")
        .filter(Boolean)
        .slice(0, 3)
        .map((line) =>
          line
            .replace(/\u001b\[[0-9;]*m/g, "")
            .replace(/https?:\/\/\S+/g, "[URL]")
            .replace(/\/(?:[^\s'"`<>:]+\/?)+/g, "[path]")
            .slice(0, 240),
        ),
    },
    children,
    source_sha256_before: sha(before),
    source_sha256_after: sha(await readFile(transcript)),
  };
  evidence.source_unchanged = evidence.source_sha256_before === evidence.source_sha256_after;
  evidence.ok &&= evidence.source_unchanged;
  json(evidence);
  process.exitCode = evidence.ok ? 0 : 1;
}

async function findCli() {
  for (const dir of (process.env.PATH || "").split(delimiter).filter(Boolean)) {
    try {
      const path = await realpath(join(dir, "codebuddy"));
      const pkg = JSON.parse(await readFile(join(dirname(dirname(path)), "package.json")));
      if (pkg.name === "@tencent-ai/codebuddy-code") return { path, version: pkg.version };
    } catch {}
  }
  throw new Error("Installed CodeBuddy CLI not found on PATH");
}

async function main() {
  const serial = process.argv.includes("--serial-handoff");
  if (serial)
    assert.ok(process.argv.includes("--self-loopback"), "Serial handoff requires --self-loopback");
  assert.equal(process.platform, "darwin", "This test requires the macOS OS sandbox");
  await stat("/usr/bin/sandbox-exec");
  const cli = await findCli();
  const sdkVersion = JSON.parse(
    await readFile(join(repo, "node_modules/@tencent-ai/agent-sdk/package.json")),
  ).version;
  const root = await mkdtemp("/private/tmp/draftspect-replay-");
  const allowed = join(root, "isolated");
  const fixtureCwd = join(allowed, "workspace");
  const fixtureConfig = join(allowed, "config");
  await mkdir(fixtureCwd, { recursive: true, mode: 0o700 });
  await mkdir(fixtureConfig, { mode: 0o700 });
  const readable = join(allowed, "probe-readable.txt");
  const forbidden = join(root, "probe-forbidden.txt");
  await writeFile(readable, "Synthetic isolation fixture", { mode: 0o600 });
  await writeFile(forbidden, "Synthetic canary outside child read allowlist", { mode: 0o600 });
  let selfPort = null;
  if (process.argv.includes("--self-loopback")) {
    const reservation = net.createServer();
    await new Promise((resolve, reject) => {
      reservation.once("error", reject);
      reservation.listen(0, "127.0.0.1", resolve);
    });
    selfPort = reservation.address().port;
    await new Promise((resolve) => reservation.close(resolve));
  }
  const q = (path) => JSON.stringify(path);
  const profile = [
    "(version 1)",
    "(allow default)",
    `(deny file-read-data (require-all (regex #"^/") (require-not (require-any ${["/System", "/usr", "/bin", "/dev", "/private/etc", allowed, join(repo, "node_modules"), dirname(dirname(cli.path)), ...(serial ? [join(repo, "daemon")] : [])].map((p) => `(subpath ${q(p)})`).join(" ")} (literal "/") (literal ${q(self)}) ${serial ? `(literal ${q(join(repo, "scripts", "serial-handoff-check.mjs"))})` : ""} (literal ${q(process.execPath)})))))`,
    `(deny file-write* (require-not (require-any (subpath ${q(allowed)}) (literal "/dev/null"))))`,
    `(deny process-exec (require-not (literal ${q(process.execPath)})))`,
    '(deny mach-lookup (global-name "com.apple.securityd") (global-name "com.apple.securityd.xpc") (global-name "com.apple.security.agent"))',
    "(deny network*)",
    ...(selfPort
      ? [
          `(allow network-bind (local tcp4 "localhost:${selfPort}"))`,
          // Seatbelt evaluates inbound permission at listen time, before a
          // remote peer exists. Actual IPv4 loopback binding is independently
          // verified below; do not claim this rejects wildcard bind syscalls.
          `(allow network-inbound (local tcp4 "localhost:${selfPort}"))`,
          `(allow network-outbound (remote tcp4 "localhost:${selfPort}"))`,
        ]
      : []),
  ].join("\n");
  const profileFile = join(root, "sandbox.sb");
  await writeFile(profileFile, profile, { mode: 0o600 });
  const env = {
    PATH: `${dirname(process.execPath)}:/usr/bin:/bin`,
    TMPDIR: allowed,
    LANG: "en_US.UTF-8",
    CODEBUDDY_CONFIG_DIR: fixtureConfig,
    ...(serial
      ? {
          CODEBUDDY_CODE_PATH: cli.path,
          CODEBUDDY_INTERNET_ENVIRONMENT: "external",
          DRAFTSPECT_CODEBUDDY_SETTING_SOURCES: "none",
          DRAFTSPECT_STATE_DIR: join(allowed, "draftspect-state"),
        }
      : {}),
    DISABLE_TELEMETRY: "1",
    DISABLE_GALILEO: "1",
    DISABLE_AUTOUPDATER: "1",
    CODEBUDDY_DISABLE_AUTO_MEMORY: "1",
    ...(selfPort ? { SERVER__PORT: String(selfPort), SERVER__HOST: "127.0.0.1" } : {}),
  };
  const run = (args, timeoutMs) =>
    new Promise((resolve) => {
      const proc = childProcess.spawn(
        "/usr/bin/sandbox-exec",
        ["-f", profileFile, process.execPath, self, ...args],
        { cwd: fixtureCwd, env, stdio: ["ignore", "pipe", "pipe"], detached: true },
      );
      let stdout = "",
        stderr = "",
        timedOut = false;
      let exited = false,
        monitorBusy = false,
        monitorWork = Promise.resolve();
      const listenerCheck = { verified: false, observed: [], error: null };
      const killGroup = () => {
        try {
          process.kill(-proc.pid, "SIGKILL");
        } catch {}
      };
      const monitor =
        selfPort && ["--replay", "--serial"].includes(args[0])
          ? setInterval(() => {
              if (monitorBusy || exited) return;
              const checkpoint = stdout
                .split("\n")
                .map((line) => {
                  try {
                    return JSON.parse(line);
                  } catch {
                    return null;
                  }
                })
                .filter((item) => item?.phase === "cli_child_started")
                .at(-1);
              if (!checkpoint?.pid) return;
              monitorBusy = true;
              monitorWork = new Promise((done) =>
                childProcess.execFile(
                  "/usr/sbin/lsof",
                  ["-nP", "-a", "-p", String(checkpoint.pid), "-iTCP", "-sTCP:LISTEN", "-Fn"],
                  { timeout: 2000 },
                  (error, output) => {
                    const addresses = output
                      .split("\n")
                      .filter((line) => line.startsWith("n"))
                      .map((line) => line.slice(1));
                    for (const address of addresses)
                      if (!listenerCheck.observed.includes(address))
                        listenerCheck.observed.push(address);
                    if (
                      addresses.some((address) => address !== `127.0.0.1:${selfPort}`) ||
                      (error && error.code !== 1)
                    ) {
                      listenerCheck.error =
                        "Unexpected CLI listener or failed scoped listener inspection";
                      if (!exited) killGroup();
                    } else if (addresses.length) listenerCheck.verified = true;
                    monitorBusy = false;
                    done();
                  },
                ),
              );
            }, 150)
          : null;
      proc.stdout.on("data", (chunk) => (stdout += chunk));
      proc.stderr.on("data", (chunk) => (stderr += chunk));
      const timer = setTimeout(() => {
        timedOut = true;
        killGroup();
      }, timeoutMs);
      proc.once("exit", async (code, signal) => {
        exited = true;
        clearTimeout(timer);
        if (monitor) clearInterval(monitor);
        await monitorWork;
        const lines = stdout.trim().split("\n");
        let evidence;
        const checkpoints = [];
        for (const line of lines) {
          try {
            const item = JSON.parse(line);
            if (item.phase === "cli_child_started") checkpoints.push(item);
          } catch {}
        }
        try {
          evidence = JSON.parse(lines.at(-1));
        } catch {}
        if (evidence?.phase === "cli_child_started") evidence = null;
        let groupExists = true;
        for (let n = 0; n < 10 && groupExists; n++) {
          try {
            process.kill(-proc.pid, 0);
          } catch (error) {
            if (error.code === "ESRCH") groupExists = false;
          }
          if (groupExists) await delay(100);
        }
        resolve({
          code,
          signal,
          timedOut,
          process_group: proc.pid,
          process_group_absent: !groupExists,
          checkpoints,
          listener_check: listenerCheck,
          evidence: evidence ?? null,
          stderr_bytes: Buffer.byteLength(stderr),
          stderr_sha256: sha(stderr),
          stderr_diagnostic: stderr
            .split("\n")
            .filter(Boolean)
            .slice(0, 3)
            .map((line) => line.replace(/https?:\/\/\S+/g, "[URL]").slice(0, 240)),
          sandbox_apply_denied: stderr.includes("sandbox_apply: Operation not permitted"),
        });
      });
    });
  const report = {
    scope: serial
      ? "Synthetic production serial-handoff component integration; no model, native pane or TUI acceptance"
      : "Installed CLI headless exact-ID history replay only; no model or TUI acceptance",
    cli_version: cli.version,
    sdk_version: sdkVersion,
    report_directory: root,
    isolation_mode: selfPort ? "one_synthetic_loopback_port_only" : "deny_all_network",
    allowed_self_port: selfPort,
    isolation: null,
    replay: null,
  };
  const server = net.createServer((socket) => socket.end());
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  try {
    const port = server.address().port;
    await new Promise((resolve, reject) => {
      const socket = net.connect({ host: "127.0.0.1", port });
      socket.once("connect", () => {
        socket.destroy();
        resolve();
      });
      socket.once("error", reject);
    });
    report.isolation = await run(
      ["--probe", readable, forbidden, String(port), ...(selfPort ? [String(selfPort)] : [])],
      10000,
    );
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
  if (report.isolation.code === 0 && report.isolation.evidence?.ok) {
    const sessionId = randomUUID();
    const projectName = fixtureCwd
      .replace(/[/\\:]/g, "-")
      .replace(/^-+|-+$/g, "")
      .replace(/-+/g, "-");
    const storage = join(fixtureConfig, "projects", projectName);
    await mkdir(storage, { recursive: true, mode: 0o700 });
    const ids = Array.from({ length: 4 }, () => randomUUID());
    const callId = randomUUID();
    const history = [
      {
        type: "message",
        role: "user",
        content: [
          { type: "input_text", text: "Synthetic archive keyword BRONZE-TULIP-42. Read only." },
        ],
      },
      {
        type: "function_call",
        callId,
        name: "mcp__office__office_read_story",
        arguments: '{"target":{"story":"body"}}',
        status: "completed",
      },
      {
        type: "function_call_result",
        callId,
        name: "mcp__office__office_read_story",
        output: { type: "text", text: "Synthetic document evidence BRONZE-TULIP-42" },
        status: "completed",
      },
      {
        type: "message",
        role: "assistant",
        content: [{ type: "output_text", text: "Synthetic archive summary: BRONZE-TULIP-42." }],
        status: "completed",
      },
    ].map((item, i) => ({
      ...item,
      id: ids[i],
      parentId: ids[i - 1] ?? null,
      sessionId,
      timestamp: Date.now() + i,
      cwd: fixtureCwd,
    }));
    const transcript = join(storage, `${sessionId}.jsonl`);
    await writeFile(transcript, history.map((item) => JSON.stringify(item)).join("\n") + "\n", {
      mode: 0o600,
    });
    report.source_sha256_before = sha(await readFile(transcript));
    report.replay = await run(
      [
        serial ? "--serial" : "--replay",
        cli.path,
        fixtureCwd,
        fixtureConfig,
        sessionId,
        transcript,
        ...(selfPort ? [String(selfPort)] : []),
      ],
      selfPort ? 70000 : 45000,
    );
    report.source_sha256_after = sha(await readFile(transcript));
    report.source_unchanged = report.source_sha256_before === report.source_sha256_after;
  }
  report.ok =
    report.isolation.evidence?.ok === true &&
    report.replay?.evidence?.ok === true &&
    (!selfPort || (report.replay.listener_check.verified && !report.replay.listener_check.error)) &&
    report.replay.process_group_absent === true &&
    report.source_unchanged === true;
  await writeFile(join(root, "report.json"), JSON.stringify(report, null, 2) + "\n", {
    mode: 0o600,
  });
  json(report);
  process.exitCode = report.ok ? 0 : 1;
}

try {
  if (process.argv[2] === "--probe") await probe();
  else if (process.argv[2] === "--replay") await replay();
  else if (process.argv[2] === "--serial") {
    const { checkSerialHandoff } = await import("./serial-handoff-check.mjs");
    json(await checkSerialHandoff(process.argv.slice(3)));
  } else await main();
} catch (error) {
  json({
    ok: false,
    failure: error.name,
    code: error.code ?? null,
    assertion: error.code === "ERR_ASSERTION" ? error.message : null,
    diagnostic: String(error.message)
      .replace(/https?:\/\/\S+/g, "[URL]")
      .slice(0, 300),
  });
  process.exitCode = 1;
}
