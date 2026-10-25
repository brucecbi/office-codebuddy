// Desktop-only startup checks. This module never imports an agent runtime,
// starts a model/CLI, installs trust, or changes saved configuration.
import { constants } from "node:fs";
import { access, realpath, stat } from "node:fs/promises";
import { createRequire } from "node:module";
import { createServer } from "node:net";
import { delimiter, isAbsolute, join } from "node:path";
import { loadLocalTransport } from "../daemon/local-transport.mjs";
import { agentAccess, codebuddyEnvironment } from "../daemon/preview-policy.mjs";

const REQUIRED_FILES = [
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
  "taskpane/shared/redline-audit.js",
  "taskpane/shared/word-ooxml-state.js",
  "taskpane/shared/word-write-safety.js",
  "taskpane/shared/styles.css",
  "manifests/word.xml",
  "manifests/excel.xml",
];
const REQUIRED_MODULES = ["ws", "zod"];
const PORT_CHECK_TIMEOUT_MS = 2_000;

function failure(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function validAbsolutePath(value) {
  return typeof value === "string" && isAbsolute(value) && !/[\0\r\n]/.test(value);
}

async function requireFile(filePath) {
  if (!(await stat(filePath)).isFile()) throw new Error("Not a regular file");
  await access(filePath, constants.R_OK);
}

async function executablePath(candidate) {
  if (!validAbsolutePath(candidate)) return null;
  try {
    const resolved = await realpath(candidate);
    if (!(await stat(resolved)).isFile()) return null;
    await access(resolved, constants.X_OK | constants.R_OK);
    return resolved;
  } catch {
    return null;
  }
}

async function findExecutable(env, override, name, label) {
  // Explicit but broken overrides must not silently select another install.
  if (env[override] !== undefined) {
    const resolved = await executablePath(env[override]);
    if (resolved) return resolved;
    throw failure(
      `DESKTOP_${label}_INVALID`,
      `${override} must identify an absolute, readable executable file. Check the desktop launcher configuration.`,
    );
  }
  for (const directory of (typeof env.PATH === "string" ? env.PATH : "").split(delimiter)) {
    // Finder must not discover executables relative to an accidental cwd.
    if (!validAbsolutePath(directory)) continue;
    const resolved = await executablePath(join(directory, name));
    if (resolved) return resolved;
  }
  throw failure(
    `DESKTOP_${label}_NOT_FOUND`,
    `${name} was not found on the launcher's absolute PATH. Configure ${override} explicitly; no bundled fallback was selected.`,
  );
}

function parsePort(value, fallback, name) {
  if (value === undefined) return fallback;
  if (typeof value !== "string" || !/^[0-9]+$/.test(value)) {
    throw failure("DESKTOP_PORT_INVALID", `${name} must be a decimal port number from 1 to 65535.`);
  }
  const port = Number(value);
  if (!Number.isSafeInteger(port) || port < 1 || port > 65535) {
    throw failure("DESKTOP_PORT_INVALID", `${name} must be a decimal port number from 1 to 65535.`);
  }
  return port;
}

function checkPort(port, label) {
  // This is a short availability check, not a lock. The daemon's later bind is
  // authoritative. Never fetch an existing server's tokens or terminate it.
  return new Promise((resolve, reject) => {
    const server = createServer((socket) => socket.destroy());
    let settled = false;
    const finish = (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (error) reject(error);
      else resolve();
    };
    const timer = setTimeout(() => {
      server.close(() => {});
      finish(
        failure(
          "DESKTOP_PORT_CHECK_TIMEOUT",
          `The ${label} loopback port check timed out. No existing process was stopped.`,
        ),
      );
    }, PORT_CHECK_TIMEOUT_MS);
    server.on("error", (error) => {
      finish(
        error.code === "EADDRINUSE"
          ? failure(
              "DESKTOP_PORT_IN_USE",
              `The ${label} loopback port ${port} is already in use. Stop its known owner explicitly, then retry. No existing process was stopped.`,
            )
          : failure(
              "DESKTOP_PORT_UNAVAILABLE",
              `The ${label} loopback port could not be checked. No existing process was stopped.`,
            ),
      );
    });
    // Also closes a hypothetical late successful bind after the timer fired.
    server.once("listening", () =>
      server.close((error) => {
        finish(
          error
            ? failure(
                "DESKTOP_PORT_UNAVAILABLE",
                `The ${label} loopback port check could not close cleanly.`,
              )
            : null,
        );
      }),
    );
    server.listen({ host: "127.0.0.1", port, exclusive: true });
  });
}

export async function desktopPreflight({ projectRoot, env = process.env, checkPorts = true } = {}) {
  if (!env || typeof env !== "object" || Array.isArray(env)) {
    throw failure(
      "DESKTOP_ENV_INVALID",
      "The desktop launch environment must be a settings object.",
    );
  }
  if (!validAbsolutePath(projectRoot)) {
    throw failure(
      "DESKTOP_PROJECT_INVALID",
      "The desktop launcher requires an absolute project directory.",
    );
  }
  for (const relative of REQUIRED_FILES) {
    try {
      await requireFile(join(projectRoot, relative));
    } catch {
      throw failure(
        "DESKTOP_ASSET_MISSING",
        `Required project file is missing or unreadable: ${relative}. Restore the project before launching.`,
      );
    }
  }
  const provider = env.DRAFTSPECT_RUNTIME || "codebuddy";
  if (!["codebuddy", "claude"].includes(provider))
    throw failure("DESKTOP_RUNTIME_INVALID", "Select a known runtime explicitly.");
  try {
    agentAccess(env);
    if (provider === "codebuddy") codebuddyEnvironment(env);
  } catch (error) {
    throw failure("DESKTOP_PREVIEW_CONFIG_REQUIRED", error.message);
  }
  const require = createRequire(join(projectRoot, "package.json"));
  for (const dependency of [
    ...REQUIRED_MODULES,
    provider === "codebuddy" ? "@tencent-ai/agent-sdk" : "@anthropic-ai/claude-agent-sdk",
  ]) {
    try {
      await requireFile(require.resolve(dependency));
    } catch {
      throw failure(
        "DESKTOP_DEPENDENCY_MISSING",
        `Required dependency is unavailable: ${dependency}. Restore this project's dependencies before launching.`,
      );
    }
  }
  const cliPath = await findExecutable(env, "CODEBUDDY_CODE_PATH", "codebuddy", "CLI");
  const nodePath = await findExecutable(env, "DRAFTSPECT_NODE_PATH", "node", "NODE");
  const httpPort = parsePort(env.DRAFTSPECT_HTTP_PORT, 47834, "DRAFTSPECT_HTTP_PORT");
  const wsPort = parsePort(env.DRAFTSPECT_WS_PORT, 47833, "DRAFTSPECT_WS_PORT");
  if (httpPort === wsPort) {
    throw failure(
      "DESKTOP_PORT_INVALID",
      "The HTTP and WebSocket loopback ports must be different.",
    );
  }
  let transport;
  try {
    transport = await loadLocalTransport(env, { warn: null });
  } catch {
    // Configuration and filesystem errors can contain secrets or user paths.
    // The UI receives a fixed diagnostic, never the underlying error/cause.
    throw failure(
      "DESKTOP_TLS_INVALID",
      "Local HTTPS configuration could not be validated. Check certificate expiry, key permissions, and saved transport settings. No trust settings were changed.",
    );
  }
  if (!transport.secure) {
    throw failure(
      "DESKTOP_HTTPS_REQUIRED",
      "Desktop launch requires the existing HTTPS/WSS configuration. Configure a valid local certificate first; plaintext fallback is disabled.",
    );
  }
  if (checkPorts) {
    await checkPort(httpPort, "HTTPS");
    await checkPort(wsPort, "WSS");
  }
  return {
    protocol: "https/wss",
    certificate: transport.certificate,
    warnings: transport.warnings,
    cliPath,
    nodePath,
  };
}
