import { after, before, test } from "node:test";
import assert from "node:assert/strict";
import {
  mkdtemp,
  mkdir,
  writeFile,
  chmod,
  rm,
  symlink,
  readFile,
  realpath,
} from "node:fs/promises";
import { createServer } from "node:net";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { desktopPreflight } from "../app/desktop-preflight.mjs";
import { runOpenSsl } from "./helpers/openssl.mjs";

const assets = [
  "package.json",
  "app/main.mjs",
  "app/addin-registration.mjs",
  "app/excel-bootstrap.mjs",
  "app/macos-excel-open.mjs",
  "app/assets/excel-initialize-synthetic.xlsx",
  "app/desktop-status-window.mjs",
  "app/desktop-status-preload.cjs",
  "app/desktop-status.html",
  "app/desktop-status.js",
  "app/desktop-status.css",
  "daemon/index.mjs",
  "taskpane/word/index.html",
  "taskpane/excel/index.html",
  "taskpane/shared/taskpane.js",
  "taskpane/shared/tools-word.js",
  "taskpane/shared/tools-excel.js",
  "taskpane/shared/paths.js",
  "taskpane/shared/operation-queue.js",
  "taskpane/shared/word-edit-plan.js",
  "taskpane/shared/word-review.js",
  "taskpane/shared/word-ooxml-state.js",
  "taskpane/shared/word-write-safety.js",
  "taskpane/shared/styles.css",
  "manifests/word.xml",
  "manifests/excel.xml",
];
const dependencies = ["@tencent-ai/agent-sdk", "@anthropic-ai/claude-agent-sdk", "ws", "zod"];
let root, keyPath, certPath;

before(async () => {
  root = await realpath(await mkdtemp(join(tmpdir(), "draftspect-desktop-preflight-")));
  keyPath = join(root, "synthetic-key.pem");
  certPath = join(root, "synthetic-cert.pem");
  await runOpenSsl([
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-days",
    "30",
    "-subj",
    "/CN=Draftspect synthetic desktop test",
    "-addext",
    "subjectAltName=IP:127.0.0.1",
    "-addext",
    "basicConstraints=critical,CA:FALSE",
    "-addext",
    "extendedKeyUsage=serverAuth",
    "-keyout",
    keyPath,
    "-out",
    certPath,
  ]);
  await chmod(keyPath, 0o600);
});
after(async () => {
  if (root) await rm(root, { recursive: true, force: true });
});

async function fixture({ omitAsset, omitDependency, omitCli = false, omitNode = false } = {}) {
  const projectRoot = await mkdtemp(join(root, "project-"));
  for (const asset of assets) {
    if (asset === omitAsset) continue;
    const destination = join(projectRoot, asset);
    await mkdir(dirname(destination), { recursive: true });
    await writeFile(
      destination,
      asset === "package.json" ? '{"type":"module"}' : "synthetic fixture\n",
    );
  }
  for (const dependency of dependencies) {
    if (dependency === omitDependency) continue;
    const directory = join(projectRoot, "node_modules", dependency);
    await mkdir(directory, { recursive: true });
    await writeFile(
      join(directory, "package.json"),
      JSON.stringify({ name: dependency, main: "index.js" }),
    );
    await writeFile(
      join(directory, "index.js"),
      'throw new Error("Provider execution is forbidden in desktop preflight");\n',
    );
  }
  const bin = join(projectRoot, "bin");
  await mkdir(bin);
  for (const [name, omitted] of [
    ["codebuddy", omitCli],
    ["node", omitNode],
  ]) {
    if (omitted) continue;
    await writeFile(join(bin, name), "#!/bin/sh\nexit 89\n", { mode: 0o700 });
  }
  const state = join(projectRoot, "state");
  await mkdir(state);
  await writeFile(
    join(state, "transport.json"),
    JSON.stringify({ version: 1, protocol: "https", keyPath, certPath }),
  );
  const env = {
    PATH: bin,
    DRAFTSPECT_STATE_DIR: state,
    CODEBUDDY_INTERNET_ENVIRONMENT: "external",
  };
  return { projectRoot, env, checkPorts: false };
}

function hasCode(code) {
  return (error) => {
    assert.equal(error.code, code);
    assert.ok(!error.message.includes("synthetic-secret"));
    assert.equal(error.cause, undefined);
    return true;
  };
}

async function listeningServer() {
  const server = createServer((socket) => socket.destroy());
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  return server;
}
const closeServer = (server) =>
  new Promise((resolve, reject) => server.close((error) => (error ? reject(error) : resolve())));
async function freePortPair() {
  const first = await listeningServer();
  const second = await listeningServer();
  const ports = [first.address().port, second.address().port];
  await Promise.all([closeServer(first), closeServer(second)]);
  return ports;
}
async function assertPortReleased(port) {
  const server = createServer();
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });
  await closeServer(server);
}

test("valid saved TLS returns public metadata, leaves configuration unchanged, and executes no provider", async () => {
  const input = await fixture();
  const saved = join(input.env.DRAFTSPECT_STATE_DIR, "transport.json");
  const before = await readFile(saved);
  const originalEnv = { ...input.env };
  const result = await desktopPreflight(input);
  assert.equal(result.protocol, "https/wss");
  assert.equal(result.certificate.native_trust, "NOT_CHECKED");
  assert.equal(result.certificate.renewal_due, false);
  assert.deepEqual(result.warnings, []);
  assert.equal(result.cliPath, join(input.projectRoot, "bin/codebuddy"));
  assert.equal(result.nodePath, join(input.projectRoot, "bin/node"));
  for (const forbidden of [
    "PRIVATE KEY",
    "BEGIN CERTIFICATE",
    "synthetic-key.pem",
    "synthetic-cert.pem",
    '"tls"',
  ]) {
    assert.ok(!JSON.stringify(result).includes(forbidden));
  }
  assert.deepEqual(await readFile(saved), before);
  assert.deepEqual(input.env, originalEnv);
});

test("invalid project root fails with a sanitized controlled error", async () => {
  for (const projectRoot of [undefined, "relative", "synthetic-secret\0", {}]) {
    await assert.rejects(desktopPreflight({ projectRoot }), hasCode("DESKTOP_PROJECT_INVALID"));
  }
});

test("each required startup asset is checked before launch", async () => {
  for (const omitAsset of assets) {
    await assert.rejects(
      desktopPreflight(await fixture({ omitAsset })),
      hasCode("DESKTOP_ASSET_MISSING"),
    );
  }
});

test("provider and shared dependency entries must resolve without importing them", async () => {
  for (const omitDependency of dependencies) {
    const input = await fixture({ omitDependency });
    if (omitDependency === "@anthropic-ai/claude-agent-sdk") {
      await assert.doesNotReject(desktopPreflight(input));
      input.env.DRAFTSPECT_RUNTIME = "claude";
    }
    await assert.rejects(desktopPreflight(input), hasCode("DESKTOP_DEPENDENCY_MISSING"));
  }
});

test("missing CLI fails closed instead of selecting the SDK's bundled CLI", async () => {
  await assert.rejects(
    desktopPreflight(await fixture({ omitCli: true })),
    hasCode("DESKTOP_CLI_NOT_FOUND"),
  );
});

test("missing node fails with a dedicated actionable error", async () => {
  await assert.rejects(
    desktopPreflight(await fixture({ omitNode: true })),
    hasCode("DESKTOP_NODE_NOT_FOUND"),
  );
});

test("explicit overrides resolve symlinks and never fall back after an invalid override", async () => {
  const input = await fixture();
  const cliAlias = join(input.projectRoot, "codebuddy-link");
  const nodeAlias = join(input.projectRoot, "node-link");
  await symlink(join(input.projectRoot, "bin/codebuddy"), cliAlias);
  await symlink(join(input.projectRoot, "bin/node"), nodeAlias);
  const result = await desktopPreflight({
    ...input,
    env: { ...input.env, CODEBUDDY_CODE_PATH: cliAlias, DRAFTSPECT_NODE_PATH: nodeAlias },
  });
  assert.equal(result.cliPath, join(input.projectRoot, "bin/codebuddy"));
  assert.equal(result.nodePath, join(input.projectRoot, "bin/node"));
  for (const [override, code] of [
    ["CODEBUDDY_CODE_PATH", "DESKTOP_CLI_INVALID"],
    ["DRAFTSPECT_NODE_PATH", "DESKTOP_NODE_INVALID"],
  ]) {
    for (const value of [
      "",
      "relative",
      "/synthetic-secret/missing",
      null,
      "/synthetic-secret\0",
      input.projectRoot,
    ]) {
      await assert.rejects(
        desktopPreflight({ ...input, env: { ...input.env, [override]: value } }),
        hasCode(code),
      );
    }
  }
});

test("non-executable CLI and node files are rejected", async () => {
  for (const [name, override, code] of [
    ["codebuddy", "CODEBUDDY_CODE_PATH", "DESKTOP_CLI_INVALID"],
    ["node", "DRAFTSPECT_NODE_PATH", "DESKTOP_NODE_INVALID"],
  ]) {
    const input = await fixture();
    const file = join(input.projectRoot, "bin", name);
    await chmod(file, 0o600);
    await assert.rejects(
      desktopPreflight({ ...input, env: { ...input.env, [override]: file } }),
      hasCode(code),
    );
  }
});

test("relative PATH entries cannot select executable files from the current directory", async () => {
  const input = await fixture();
  await assert.rejects(
    desktopPreflight({ ...input, env: { ...input.env, PATH: ":.:bin" } }),
    hasCode("DESKTOP_CLI_NOT_FOUND"),
  );
});

test("absent saved HTTPS configuration cannot downgrade the desktop launcher to HTTP", async () => {
  const input = await fixture();
  input.env.DRAFTSPECT_STATE_DIR = join(input.projectRoot, "empty-state");
  await assert.rejects(desktopPreflight(input), hasCode("DESKTOP_HTTPS_REQUIRED"));
});

test("TLS failures never expose malformed configuration or private file paths", async () => {
  const input = await fixture();
  await writeFile(
    join(input.env.DRAFTSPECT_STATE_DIR, "transport.json"),
    "synthetic-secret malformed JSON",
  );
  await assert.rejects(desktopPreflight(input), hasCode("DESKTOP_TLS_INVALID"));
  await assert.rejects(
    desktopPreflight({
      ...input,
      env: {
        ...input.env,
        DRAFTSPECT_TLS_KEY: "/synthetic-secret/missing-key",
        DRAFTSPECT_TLS_CERT: "/synthetic-secret/missing-cert",
      },
    }),
    hasCode("DESKTOP_TLS_INVALID"),
  );
});

test("port inputs reject coercion, out-of-range and duplicate values", async () => {
  const input = await fixture();
  for (const key of ["DRAFTSPECT_HTTP_PORT", "DRAFTSPECT_WS_PORT"]) {
    for (const value of [
      "",
      "0",
      "65536",
      "-1",
      "1.1",
      "1e3",
      "0xCAFE",
      " 1234",
      "1234\n",
      "synthetic-secret",
      1234,
      null,
    ]) {
      await assert.rejects(
        desktopPreflight({ ...input, env: { ...input.env, [key]: value } }),
        hasCode("DESKTOP_PORT_INVALID"),
      );
    }
  }
  await assert.rejects(
    desktopPreflight({ ...input, env: { ...input.env, DRAFTSPECT_HTTP_PORT: "47833" } }),
    hasCode("DESKTOP_PORT_INVALID"),
  );
});

test("real loopback success closes both temporary listeners", async () => {
  const input = await fixture();
  const [httpPort, wsPort] = await freePortPair();
  const result = await desktopPreflight({
    ...input,
    checkPorts: true,
    env: {
      ...input.env,
      DRAFTSPECT_HTTP_PORT: String(httpPort),
      DRAFTSPECT_WS_PORT: String(wsPort),
    },
  });
  assert.equal(result.protocol, "https/wss");
  await assertPortReleased(httpPort);
  await assertPortReleased(wsPort);
});

test("occupied loopback ports are refused without connecting to or stopping their owners", async () => {
  for (const occupiedKey of ["DRAFTSPECT_HTTP_PORT", "DRAFTSPECT_WS_PORT"]) {
    const input = await fixture();
    const occupied = await listeningServer();
    let connections = 0;
    occupied.on("connection", () => connections++);
    try {
      const [httpPort, wsPort] = await freePortPair();
      const port = occupied.address().port;
      await assert.rejects(
        desktopPreflight({
          ...input,
          checkPorts: true,
          env: {
            ...input.env,
            DRAFTSPECT_HTTP_PORT: String(httpPort),
            DRAFTSPECT_WS_PORT: String(wsPort),
            [occupiedKey]: String(port),
          },
        }),
        hasCode("DESKTOP_PORT_IN_USE"),
      );
      assert.equal(occupied.listening, true);
      assert.equal(occupied.address().port, port);
      assert.equal(connections, 0);
      await assertPortReleased(occupiedKey === "DRAFTSPECT_HTTP_PORT" ? wsPort : httpPort);
    } finally {
      await closeServer(occupied);
    }
  }
});
