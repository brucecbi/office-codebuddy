// Local transport only. Importing this module never starts the agent/runtime,
// reads provider/user-agent settings, installs trust, or contacts an external service.
import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { createSecureContext } from "node:tls";
import { X509Certificate } from "node:crypto";
import { readFile, stat, lstat } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { stateFile } from "./state-paths.mjs";

// Ports are fixed so the manifest, CSP and pane can reference them. Tests may
// override via env to run an isolated daemon alongside a production one.
export const TASKPANE_PORT = Number(process.env.DRAFTSPECT_HTTP_PORT) || 47834;
export const BRIDGE_PORT = Number(process.env.DRAFTSPECT_WS_PORT) || 47833;
const RENEWAL_WARNING_DAYS = 14;
const DAY_MS = 86_400_000;

export function transportSettings(env = process.env) {
  const keyPath = env.DRAFTSPECT_TLS_KEY;
  const certPath = env.DRAFTSPECT_TLS_CERT;
  if (!!keyPath !== !!certPath)
    throw new Error("Set both DRAFTSPECT_TLS_KEY and DRAFTSPECT_TLS_CERT, or neither");
  if ([keyPath, certPath].some((path) => path && (typeof path !== "string" || !path.trim())))
    throw new Error("TLS file paths must be nonempty strings");
  const secure = !!keyPath;
  return {
    keyPath,
    certPath,
    secure,
    origin: `${secure ? "https" : "http"}://127.0.0.1:${TASKPANE_PORT}`,
    source: secure ? "environment" : "default",
  };
}

export async function resolveLocalTransport(
  env = process.env,
  { configPath: explicitConfigPath } = {},
) {
  const explicit = transportSettings(env);
  // An incomplete environment override is an error even if a saved pair exists.
  if (explicit.secure) return explicit;
  const configPath = explicitConfigPath ?? stateFile("transport.json", env);
  let raw;
  try {
    raw = await readFile(configPath, "utf8");
  } catch (error) {
    if (error.code === "ENOENT" && explicitConfigPath === undefined) {
      // A dangling saved-config symlink is a broken configuration, not absence.
      const entry = await lstat(configPath).catch((statError) => {
        if (statError.code === "ENOENT") return null;
        throw statError;
      });
      if (!entry) return { ...explicit, configPath };
    }
    throw error;
  }
  let config;
  try {
    config = JSON.parse(raw);
  } catch {
    // Do not echo the input: malformed local configuration may contain secrets.
    throw new Error(`Invalid transport configuration at ${configPath}: expected JSON`);
  }
  if (
    !config ||
    Array.isArray(config) ||
    config.version !== 1 ||
    config.protocol !== "https" ||
    Object.keys(config).some(
      (key) => !["version", "protocol", "keyPath", "certPath"].includes(key),
    ) ||
    [config.keyPath, config.certPath].some(
      (path) => typeof path !== "string" || !path.trim() || path.includes("\0"),
    )
  )
    throw new Error(
      `Invalid transport configuration at ${configPath}: expected version 1, protocol https, keyPath and certPath`,
    );
  return {
    ...transportSettings({
      DRAFTSPECT_TLS_KEY: resolve(dirname(configPath), config.keyPath),
      DRAFTSPECT_TLS_CERT: resolve(dirname(configPath), config.certPath),
    }),
    source: "configuration",
    configPath: resolve(configPath),
  };
}

export function inspectLocalCertificate(pem, { now = Date.now() } = {}) {
  const certificate = new X509Certificate(pem);
  const validFrom = Date.parse(certificate.validFrom);
  const validTo = Date.parse(certificate.validTo);
  if (!Number.isFinite(now) || !Number.isFinite(validFrom) || !Number.isFinite(validTo))
    throw new Error("Cannot determine TLS certificate validity");
  if (now < validFrom) throw new Error("TLS certificate is not valid yet");
  if (now >= validTo) throw new Error("TLS certificate has expired; renew it before starting");
  if (
    certificate.ca ||
    certificate.checkIP("127.0.0.1") !== "127.0.0.1" ||
    certificate.subjectAltName !== "IP Address:127.0.0.1"
  )
    throw new Error("TLS certificate must be a non-CA leaf with SAN only IP:127.0.0.1");
  const remainingDays = (validTo - now) / DAY_MS;
  return {
    valid_from: new Date(validFrom).toISOString(),
    expires_at: new Date(validTo).toISOString(),
    remaining_days: Math.round(remainingDays * 100) / 100,
    renewal_due: remainingDays <= RENEWAL_WARNING_DAYS,
    fingerprint_sha256: certificate.fingerprint256,
    // Parsing and matching do not prove native/macOS certificate trust.
    native_trust: "NOT_CHECKED",
  };
}

export async function loadLocalTransport(
  env = process.env,
  { configPath, now = Date.now(), warn = console.warn } = {},
) {
  const settings = await resolveLocalTransport(env, { configPath });
  let tls = null;
  let certificate = null;
  const warnings = [];
  if (settings.secure) {
    const keyStat = await stat(settings.keyPath);
    if (!keyStat.isFile()) throw new Error("TLS private key must be a regular file");
    if (process.platform !== "win32" && (keyStat.mode & 0o077) !== 0)
      throw new Error("TLS private key permissions must exclude group/other access (use 0600)");
    tls = {
      key: await readFile(settings.keyPath),
      cert: await readFile(settings.certPath),
      minVersion: "TLSv1.2",
    };
    // Reject invalid or mismatched material before opening any listener. Never
    // fall back to plaintext, generate a certificate, or install trust here.
    createSecureContext(tls);
    certificate = inspectLocalCertificate(tls.cert, { now });
    if (certificate.renewal_due) {
      const message = `Local TLS certificate expires ${certificate.expires_at}; renew within ${certificate.remaining_days} days and recheck native trust`;
      warnings.push(message);
      warn?.(`[transport] ${message}`);
    }
  }
  return { ...settings, tls, certificate, warnings };
}

export function renderManifestTransport(xml, settings) {
  // Restrict replacement to this application's service URLs. XML namespace
  // URIs, identifiers, requirements and third-party resources stay untouched.
  return xml.replace(/https?:\/\/127\.0\.0\.1:47834(?=[/"<\s])/g, settings.origin);
}

export function isLoopbackHost(req, port) {
  return req.headers.host === `127.0.0.1:${port}`;
}

export function createLocalServer({ port, tls = null, handler }) {
  const secure = !!tls;
  const guarded = (req, res) => {
    const boundPort = server.address()?.port;
    const origin = `${secure ? "https" : "http"}://127.0.0.1:${boundPort}`;
    const requestOrigin = req.headers.origin;
    if (
      !isLoopbackHost(req, boundPort) ||
      (requestOrigin !== undefined && requestOrigin !== origin)
    ) {
      res.writeHead(403).end("Forbidden local origin or host");
      return;
    }
    if (requestOrigin === origin) res.setHeader("Access-Control-Allow-Origin", origin);
    return handler(req, res);
  };
  const server = tls ? createHttpsServer(tls, guarded) : createHttpServer(guarded);
  // Track live sockets so close() can release the port immediately instead of
  // waiting for idle keep-alive connections to drain — a restart racing a
  // lingering connection crashed with EADDRINUSE (2026-09-20 restart loop).
  const sockets = new Set();
  const track = (socket) => {
    sockets.add(socket);
    socket.on("close", () => sockets.delete(socket));
  };
  server.on("connection", track);
  server.on("secureConnection", track);
  const ready = new Promise((resolve, reject) => {
    server.once("error", reject);
    server.once("listening", () => {
      server.removeListener("error", reject);
      resolve(server.address());
    });
    server.listen(port, "127.0.0.1");
  });
  // Callers may initialize other components before awaiting readiness. Keep a
  // rejection handled during that interval without changing the promise result.
  ready.catch(() => {});
  return {
    server,
    ready,
    address: () => server.address(),
    close: () =>
      new Promise((resolve, reject) => {
        for (const socket of sockets) socket.destroy();
        server.closeIdleConnections?.();
        server.close((error) => {
          if (error && error.code !== "ERR_SERVER_NOT_RUNNING") reject(error);
          else resolve();
        });
      }),
  };
}
