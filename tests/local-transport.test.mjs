import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, chmod, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { runOpenSsl } from "./helpers/openssl.mjs";
import { once } from "node:events";
import http from "node:http";
import https from "node:https";
import WebSocket from "ws";
import vm from "node:vm";
import {
  transportSettings,
  loadLocalTransport,
  renderManifestTransport,
  createLocalServer,
} from "../daemon/local-transport.mjs";
import { createBridge } from "../daemon/bridge.mjs";

let fixtureDir, cert, tls, fixtureEnv;
before(async () => {
  fixtureDir = await mkdtemp(join(tmpdir(), "draftspect-transport-test-"));
  fixtureEnv = {
    DRAFTSPECT_TLS_KEY: join(fixtureDir, "key.pem"),
    DRAFTSPECT_TLS_CERT: join(fixtureDir, "cert.pem"),
  };
  await runOpenSsl([
    "req",
    "-x509",
    "-newkey",
    "rsa:2048",
    "-nodes",
    "-days",
    "1",
    "-subj",
    "/CN=Draftspect synthetic transport test",
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
  ({ tls } = await loadLocalTransport(fixtureEnv));
  cert = await readFile(fixtureEnv.DRAFTSPECT_TLS_CERT);
});
after(async () => {
  if (fixtureDir) await rm(fixtureDir, { recursive: true });
});

function request(server, secure, headers = {}, trust = true) {
  return new Promise((resolve, reject) => {
    const req = (secure ? https : http).get(
      {
        hostname: "127.0.0.1",
        port: server.address().port,
        path: "/bridge-token",
        headers,
        agent: false,
        // Keep TLS identity tied to the explicit loopback connection, even when
        // testing an intentionally forged HTTP Host header. No SNI for an IP.
        servername: "",
        ...(secure && trust ? { ca: cert } : {}),
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () =>
          resolve({
            status: res.statusCode,
            headers: res.headers,
            text: Buffer.concat(chunks).toString(),
          }),
        );
      },
    );
    req.on("error", reject);
  });
}

test("TLS settings are paired, default to HTTP, and fail closed on unreadable/invalid/mismatched material", async () => {
  assert.equal(transportSettings({}).origin, "http://127.0.0.1:47834");
  assert.equal(transportSettings(fixtureEnv).origin, "https://127.0.0.1:47834");
  for (const name of ["DRAFTSPECT_TLS_KEY", "DRAFTSPECT_TLS_CERT"])
    await assert.rejects(loadLocalTransport({ [name]: "fixture" }), /Set both/);
  await assert.rejects(
    loadLocalTransport({ ...fixtureEnv, DRAFTSPECT_TLS_KEY: join(fixtureDir, "missing.pem") }),
    /ENOENT/,
  );
  const invalidKey = join(fixtureDir, "invalid-key.pem");
  await writeFile(invalidKey, cert, { mode: 0o600 });
  await assert.rejects(
    loadLocalTransport({ ...fixtureEnv, DRAFTSPECT_TLS_KEY: invalidKey }),
    /unsupported|no start line/i,
  );
  const otherKey = join(fixtureDir, "other-key.pem");
  await runOpenSsl([
    "genpkey",
    "-algorithm",
    "RSA",
    "-pkeyopt",
    "rsa_keygen_bits:2048",
    "-out",
    otherKey,
  ]);
  await chmod(otherKey, 0o600);
  await assert.rejects(
    loadLocalTransport({ ...fixtureEnv, DRAFTSPECT_TLS_KEY: otherKey }),
    /key values mismatch/i,
  );
});

test("manifest transport rendering changes only local service URLs, including ribbon resources", async () => {
  for (const host of ["word", "excel"]) {
    const original = await readFile(new URL(`../manifests/${host}.xml`, import.meta.url), "utf8");
    const source = original + '<bt:Url DefaultValue="http://127.0.0.1:47834/commands.html" />';
    const secure = renderManifestTransport(source, transportSettings(fixtureEnv));
    assert.ok(!secure.includes("http://127.0.0.1:47834"));
    assert.match(secure, /http:\/\/schemas\.microsoft\.com/);
    assert.equal(
      renderManifestTransport(secure, transportSettings({})),
      source.replaceAll("https://127.0.0.1:47834", "http://127.0.0.1:47834"),
    );
  }
});

test("taskpane selects exact WS/WSS endpoint and both host CSPs permit it", async () => {
  const source = await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");
  const declaration = source.match(/^const WS_URL = .*;$/m)?.[0];
  assert.ok(declaration);
  for (const [protocol, expected] of [
    ["http:", "ws"],
    ["https:", "wss"],
  ]) {
    const context = vm.createContext({ location: { protocol } });
    vm.runInContext(declaration + ";globalThis.result = WS_URL", context);
    assert.equal(context.result, `${expected}://127.0.0.1:47833`);
    for (const host of ["word", "excel"]) {
      const html = await readFile(
        new URL(`../taskpane/${host}/index.html`, import.meta.url),
        "utf8",
      );
      assert.ok(html.match(/connect-src 'self'[^;]+;/)[0].includes(context.result));
    }
  }
});

for (const secure of [false, true]) {
  test(`${secure ? "HTTPS/WSS" : "HTTP/WS"} round-trip enforces Host, Origin and token, then releases both listeners`, async () => {
    const assets = createLocalServer({
      port: 0,
      tls: secure ? tls : null,
      handler: (_req, res) => res.end("synthetic-token"),
    });
    let bridge;
    const clients = [];
    try {
      await assets.ready;
      const origin = `${secure ? "https" : "http"}://127.0.0.1:${assets.address().port}`;
      assert.equal(assets.address().address, "127.0.0.1");
      assert.equal((await request(assets, secure)).text, "synthetic-token");
      const own = await request(assets, secure, { Origin: origin });
      assert.equal(own.status, 200);
      assert.equal(own.headers["access-control-allow-origin"], origin);
      for (const headers of [
        { Host: "untrusted.invalid" },
        { Origin: "https://untrusted.invalid" },
        { Origin: "null" },
        { Origin: "" },
      ]) {
        const denied = await request(assets, secure, headers);
        assert.equal(denied.status, 403);
        assert.ok(!denied.text.includes("synthetic-token"));
        assert.equal(denied.headers["access-control-allow-origin"], undefined);
      }
      if (secure)
        await assert.rejects(request(assets, true, {}, false), /self-signed|unable to verify/i);
      bridge = createBridge({
        port: 0,
        tls: secure ? tls : null,
        token: "synthetic-token",
        allowedOrigins: [origin],
      });
      await bridge.ready;
      assert.equal(bridge.address().address, "127.0.0.1");
      const endpoint = `${secure ? "wss" : "ws"}://127.0.0.1:${bridge.address().port}`;
      const connect = (options = {}) => {
        const ws = new WebSocket(endpoint, {
          origin,
          ...(secure ? { ca: cert, servername: "" } : {}),
          ...options,
        });
        clients.push(ws);
        return ws;
      };
      for (const options of [
        { origin: "https://untrusted.invalid" },
        { origin: "" },
        { headers: { Host: "untrusted.invalid" } },
      ]) {
        await assert.rejects(once(connect(options), "open"), /403/);
      }
      const bad = connect();
      await once(bad, "open");
      const denied = once(bad, "close");
      bad.send(JSON.stringify({ type: "hello", token: "wrong", host: "word" }));
      assert.equal((await denied)[0], 4002);
      const ws = connect();
      await once(ws, "open");
      const welcomed = once(ws, "message");
      ws.send(
        JSON.stringify({
          type: "hello",
          token: "synthetic-token",
          host: "word",
          active_doc: "/synthetic.docx",
        }),
      );
      assert.equal(JSON.parse(String((await welcomed)[0])).type, "welcome");
      const pong = once(ws, "message");
      ws.send(JSON.stringify({ type: "ping" }));
      assert.equal(JSON.parse(String((await pong)[0])).type, "pong");
    } finally {
      clients.forEach((ws) => ws.terminate());
      await bridge?.close();
      await assets.close();
    }
    assert.equal(bridge.address(), null);
    assert.equal(assets.address(), null);
  });
}

test("bind failure rejects readiness and closes cleanly without a false ready signal", async () => {
  const occupied = createLocalServer({ port: 0, handler: (_req, res) => res.end() });
  await occupied.ready;
  const failed = createBridge({ port: occupied.address().port, tls, token: "fixture" });
  let readySent = false;
  try {
    await assert.rejects(
      (async () => {
        await Promise.all([occupied.ready, failed.ready]);
        readySent = true;
      })(),
      /EADDRINUSE/,
    );
    assert.equal(readySent, false);
  } finally {
    await failed.close();
    await occupied.close();
  }
  const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
  assert.ok(
    source.indexOf("await Promise.all([http.ready, bridge.ready])") <
      source.indexOf('process.send({ type: "daemon_ready" })'),
  );
});
