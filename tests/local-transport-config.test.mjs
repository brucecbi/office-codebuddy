import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import {
  mkdtemp,
  mkdir,
  readFile,
  writeFile,
  chmod,
  copyFile,
  symlink,
  rm,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { X509Certificate } from "node:crypto";
import {
  loadLocalTransport,
  resolveLocalTransport,
  inspectLocalCertificate,
} from "../daemon/local-transport.mjs";
import { localTransportStatus } from "../scripts/local-transport-status.mjs";
import { stateDirectory, stateFile } from "../daemon/state-paths.mjs";
import { runOpenSsl } from "./helpers/openssl.mjs";

const run = promisify(execFile);
const repo = fileURLToPath(new URL("../", import.meta.url));
let root, cert, fixtureEnv;
before(async () => {
  root = await mkdtemp(join(tmpdir(), "draftspect-transport-config-test-"));
  fixtureEnv = {
    DRAFTSPECT_TLS_KEY: join(root, "key.pem"),
    DRAFTSPECT_TLS_CERT: join(root, "cert.pem"),
  };
  await runOpenSsl([
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-days",
    "30",
    "-subj",
    "/CN=Draftspect synthetic configuration test",
    "-addext",
    "subjectAltName=IP:127.0.0.1",
    "-addext",
    "basicConstraints=critical,CA:FALSE",
    "-addext",
    "extendedKeyUsage=serverAuth",
    "-keyout",
    fixtureEnv.DRAFTSPECT_TLS_KEY,
    "-out",
    fixtureEnv.DRAFTSPECT_TLS_CERT,
  ]);
  await chmod(fixtureEnv.DRAFTSPECT_TLS_KEY, 0o600);
  cert = await readFile(fixtureEnv.DRAFTSPECT_TLS_CERT);
});
after(async () => {
  if (root) await rm(root, { recursive: true });
});

async function state(name, contents) {
  const path = join(root, name, "transport.json");
  await mkdir(dirname(path), { recursive: true });
  if (contents !== undefined)
    await writeFile(path, typeof contents === "string" ? contents : JSON.stringify(contents));
  return { path, env: { DRAFTSPECT_STATE_DIR: dirname(path) } };
}
const configuration = () => ({
  version: 1,
  protocol: "https",
  keyPath: "../key.pem",
  certPath: "../cert.pem",
});

test("saved transport is state-scoped, relative to its file, and defaults to HTTP only when absent", async () => {
  const empty = await state("missing");
  assert.equal(stateDirectory(empty.env), dirname(empty.path));
  assert.equal(stateFile("transport.json", empty.env), empty.path);
  const fallback = await loadLocalTransport(empty.env);
  assert.equal(fallback.source, "default");
  assert.equal(fallback.secure, false);
  const saved = await state("saved", configuration());
  const loaded = await loadLocalTransport(saved.env);
  assert.equal(loaded.source, "configuration");
  assert.equal(loaded.origin, "https://127.0.0.1:47834");
  assert.equal(loaded.keyPath, fixtureEnv.DRAFTSPECT_TLS_KEY);
  assert.equal(loaded.certPath, fixtureEnv.DRAFTSPECT_TLS_CERT);
  assert.equal(loaded.certificate.renewal_due, false);
  assert.deepEqual(loaded.warnings, []);
});

test("paired environment overrides take precedence, while partial pairs never use saved settings", async () => {
  const invalid = await state("override", "malformed synthetic configuration");
  assert.equal((await loadLocalTransport({ ...invalid.env, ...fixtureEnv })).source, "environment");
  for (const key of ["DRAFTSPECT_TLS_KEY", "DRAFTSPECT_TLS_CERT"])
    await assert.rejects(
      loadLocalTransport({ ...invalid.env, [key]: fixtureEnv[key] }),
      /Set both/,
    );
});

test("invalid, missing-explicit, dangling and unusable saved transport never downgrade", async () => {
  const variants = [
    "synthetic-secret malformed JSON",
    null,
    [],
    {},
    { ...configuration(), version: 2 },
    { ...configuration(), protocol: "http" },
    { ...configuration(), keyPath: " " },
    { ...configuration(), certPath: "bad\0path" },
    { ...configuration(), keypath: "typo" },
  ];
  for (let i = 0; i < variants.length; i++) {
    const candidate = await state(`invalid-${i}`, variants[i]);
    await assert.rejects(loadLocalTransport(candidate.env), (error) => {
      assert.match(error.message, /Invalid transport configuration/);
      assert.ok(!error.message.includes("synthetic-secret"));
      return true;
    });
  }
  const broken = await state("missing-material", { ...configuration(), keyPath: "missing.pem" });
  await assert.rejects(loadLocalTransport(broken.env), /ENOENT/);
  await assert.rejects(
    resolveLocalTransport({}, { configPath: join(root, "not-present.json") }),
    /ENOENT/,
  );
  const dangling = await state("dangling");
  await symlink(join(root, "missing-target.json"), dangling.path);
  await assert.rejects(loadLocalTransport(dangling.env), /ENOENT/);
});

test("certificate date boundaries fail closed and renewal warnings begin fourteen days before expiry", async () => {
  const x509 = new X509Certificate(cert);
  const from = Date.parse(x509.validFrom),
    to = Date.parse(x509.validTo);
  await assert.rejects(loadLocalTransport(fixtureEnv, { now: from - 1 }), /not valid yet/);
  await assert.rejects(loadLocalTransport(fixtureEnv, { now: to }), /expired/);
  assert.throws(() => inspectLocalCertificate(cert, { now: NaN }), /Cannot determine/);
  const warnings = [];
  const due = await loadLocalTransport(fixtureEnv, {
    now: to - 14 * 86_400_000,
    warn: (message) => warnings.push(message),
  });
  assert.equal(due.certificate.renewal_due, true);
  assert.equal(due.certificate.remaining_days, 14);
  assert.equal(warnings.length, 1);
  assert.equal(due.warnings.length, 1);
  assert.equal(inspectLocalCertificate(cert, { now: to - 14 * 86_400_000 - 1 }).renewal_due, false);
});

test("CA, wrong-IP and broadened-SAN certificates are refused", async () => {
  for (const [name, san, ca] of [
    ["ca", "IP:127.0.0.1", "TRUE"],
    ["wrong-ip", "IP:127.0.0.2", "FALSE"],
    ["extra-dns", "IP:127.0.0.1,DNS:localhost", "FALSE"],
  ]) {
    const path = join(root, `${name}.pem`);
    await runOpenSsl([
      "req",
      "-new",
      "-x509",
      "-key",
      fixtureEnv.DRAFTSPECT_TLS_KEY,
      "-days",
      "30",
      "-subj",
      "/CN=Synthetic scope rejection fixture",
      "-addext",
      `subjectAltName=${san}`,
      "-addext",
      `basicConstraints=critical,CA:${ca}`,
      "-out",
      path,
    ]);
    await assert.rejects(
      loadLocalTransport({ ...fixtureEnv, DRAFTSPECT_TLS_CERT: path }),
      /non-CA leaf with SAN only/,
    );
  }
});

test("private key permissions are checked and status exposes public metadata only", async () => {
  if (process.platform !== "win32") {
    const insecure = join(root, "world-readable-key.pem");
    await copyFile(fixtureEnv.DRAFTSPECT_TLS_KEY, insecure);
    await chmod(insecure, 0o644);
    await assert.rejects(
      loadLocalTransport({ ...fixtureEnv, DRAFTSPECT_TLS_KEY: insecure }),
      /permissions/,
    );
  }
  const saved = await state("status", configuration());
  const result = await localTransportStatus(saved.env);
  assert.equal(result.protocol, "https/wss");
  assert.equal(result.native_trust, "NOT_CHECKED");
  assert.equal(result.listeners_started, false);
  assert.equal(result.model_requests, 0);
  assert.equal(result.certificate.fingerprint_sha256, new X509Certificate(cert).fingerprint256);
  const output = JSON.stringify(result);
  for (const forbidden of ["PRIVATE KEY", "BEGIN CERTIFICATE", "key.pem", "cert.pem", '"tls"'])
    assert.ok(!output.includes(forbidden));
});

test("status CLI validates explicit candidates without creating state or listeners", async () => {
  const saved = await state("cli", configuration());
  const before = await readFile(saved.path);
  const commandEnv = { PATH: process.env.PATH, DRAFTSPECT_STATE_DIR: join(root, "unused-state") };
  const { stdout, stderr } = await run(
    process.execPath,
    ["scripts/local-transport-status.mjs", "--config", saved.path, "--require-https"],
    { cwd: repo, env: commandEnv },
  );
  assert.equal(stderr, "");
  assert.equal(JSON.parse(stdout).source, "configuration");
  assert.deepEqual(await readFile(saved.path), before);
  await assert.rejects(
    run(process.execPath, ["scripts/local-transport-status.mjs", "--require-https"], {
      cwd: repo,
      env: commandEnv,
    }),
    (error) => {
      assert.equal(error.code, 1);
      assert.match(error.stderr, /HTTPS is required/);
      return true;
    },
  );
});
