import runtime from "./runtime/index.mjs";
import {
  acceptanceProfile,
  acceptanceAllowsDocument,
  assertAcceptanceWorkspace,
} from "./acceptance-profile.mjs";
import { stateFile } from "./state-paths.mjs";
import { forwardUserMcp, codebuddyEnvironment, agentAccess } from "./preview-policy.mjs";
import { createTurnSource } from "./runtime/turns.mjs";
import { classifyResult, classifyAgentError } from "./runtime/messages.mjs";
import { createLifecycle } from "./runtime/lifecycle.mjs";
import { createMessageRenderer } from "./runtime/render-messages.mjs";
import { pickModel, normalizeModelList } from "./runtime/models.mjs";
import { readFile, writeFile, mkdir, chmod, lstat } from "node:fs/promises";
import { writeRegularFile } from "./regular-file.mjs";
import { randomBytes } from "node:crypto";
import {
  loadLocalTransport,
  createLocalServer,
  TASKPANE_PORT,
  BRIDGE_PORT,
} from "./local-transport.mjs";
import { fileURLToPath } from "node:url";
import { dirname, resolve, extname, join, sep } from "node:path";
import { homedir } from "node:os";
import { createBridge } from "./bridge.mjs";
import { createOfficeBridgeMcp } from "./office-tools.mjs";
import { resolveWorkspaceRoot, suggestWorkspaceRoot, ensureWorkspaceMarker } from "./workspace.mjs";
import { randomUUID } from "node:crypto";
import { getSessionId, saveSessionId, touchFolder, getHandoff, setHandoff } from "./sessions.mjs";
import { archiveSession, resumeCommands } from "./session-archive.mjs";
import {
  captureTranscriptFingerprint,
  sameTranscriptFingerprint,
} from "./session-transcript-state.mjs";
import { readTranscript } from "./transcript.mjs";
import { installTimestampedConsole } from "./log.mjs";
import { customPermissionHandler } from "./permission-guard.mjs";

// Long-running process: every log line carries an ISO timestamp so incidents
// can be reconstructed from daemon.log (they could not, on 2026-09-20).
installTimestampedConsole();
import { createRuntimeCloseRegistry } from "./runtime/close-registry.mjs";
import { createSessionPersistenceRegistry } from "./session-persistence.mjs";
import { diag } from "./diag.mjs";
import { getContextEntries, setContextEntries } from "./context.mjs";
import { stat } from "node:fs/promises";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, "..");
let shuttingDown = false;

// Only an IPC-spawned daemon belongs to an Electron parent. Observe loss
// before the first startup await; standalone `npm run dev` is unaffected.
if (process.send) process.once("disconnect", handleParentDisconnect);

function handleParentDisconnect() {
  const forced = setTimeout(() => {
    console.error("[daemon] RUNTIME_EXIT_UNCONFIRMED: parent-loss shutdown deadline exceeded");
    process.exit(1);
  }, 3500);
  // During early startup some resources may still be uninitialized. A safe
  // self-exit releases this process's sockets; no other PID is inspected/killed.
  void shutdown()
    .catch(() => process.exit(0))
    .finally(() => clearTimeout(forced));
}

// Draftspect deliberately uses 47833/47834. Another local Office add-in
// on this machine may bind 47823/47824, so a distinct pair avoids a port
// clash when both run at once. Keep WS_PORT = HTTP_PORT - 1 (the taskpane
// derives nothing; the WS port is referenced explicitly in taskpane.js
// and the index.html CSP — change all three together).
const WS_PORT = BRIDGE_PORT;
const HTTP_PORT = TASKPANE_PORT;
agentAccess();
if (runtime.id === "codebuddy" && !acceptanceProfile) codebuddyEnvironment();
const localTransport = await loadLocalTransport();
const HTTP_ORIGIN = localTransport.origin;
const TOKEN_FILE = stateFile("bridge-token");

// Bridge token — random per-daemon-start. The taskpane fetches it from the
// HTTP server's /bridge-token endpoint (same-origin, CORS-restricted) and
// includes it in the first WS hello. Any WS that doesn't present this token
// (or comes from an unknown origin) is closed.
const BRIDGE_TOKEN = randomBytes(24).toString("hex");
{
  await mkdir(dirname(TOKEN_FILE), { recursive: true, mode: 0o700 });
  const directory = await lstat(dirname(TOKEN_FILE));
  if (!directory.isDirectory() || directory.isSymbolicLink())
    throw new Error("Bridge token state directory must be a real directory");
  await writeRegularFile(TOKEN_FILE, BRIDGE_TOKEN, { mode: 0o600 });
  console.log(`[daemon] Bridge token written to ${TOKEN_FILE}`);
}

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const matterFolder =
  acceptanceProfile?.workspace ??
  (process.argv[2]
    ? resolve(process.argv[2])
    : process.env.MATTER_FOLDER
      ? resolve(process.env.MATTER_FOLDER)
      : process.cwd());
if (acceptanceProfile)
  console.log(
    "[acceptance] synthetic Word document only; CodeBuddy pinned model; built-ins, user MCP/settings/hooks disabled",
  );

console.log(`[daemon] Workspace folder (agent cwd): ${matterFolder}`);
console.log(`[daemon] Runtime: ${runtime.id}${runtime.id === "codebuddy" ? " (default)" : ""}`);
// Availability probe — logged only when something is off (import/CLI missing).
runtime
  .probe?.()
  .then((p) => {
    if (p && !p.ok) console.warn(`[runtime] probe failed: ${p.detail}`);
  })
  .catch(() => {});

// ---------------------------------------------------------------------------
// HTTP server: serve the taskpane assets so Word can load them.
// ---------------------------------------------------------------------------
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const taskpaneDir = join(PROJECT_ROOT, "taskpane");

// A branded, actionable error page. Office renders whatever the manifest's
// SourceLocation returns inside the task pane, so a bare "Not found" (the
// old body) left users staring at two unhelpful words. The common real
// causes are all recoverable; spell them out. Served only to document
// navigations (Accept: text/html) — asset fetches still get terse text so
// nothing tries to parse HTML as JS/CSS.
const htmlEscape = (s) =>
  String(s).replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );

function errorPageHtml(status, headline, requestedPath) {
  requestedPath = htmlEscape(requestedPath);
  return `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Draftspect — ${status}</title>
<style>
  body{font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
       color:#1a1a1a;margin:0;padding:28px 24px;background:#fff}
  h1{font-size:18px;margin:0 0 4px}
  .sub{color:#666;margin:0 0 18px}
  ol{padding-left:20px;margin:0 0 18px} li{margin:6px 0}
  code{background:#f2f2f2;padding:1px 5px;border-radius:4px;font-size:12px}
  .foot{color:#999;font-size:12px;border-top:1px solid #eee;padding-top:12px}
</style></head><body>
<h1>Draftspect couldn't load this panel</h1>
<p class="sub">${headline}</p>
<p>This usually means one of:</p>
<ol>
  <li>The <strong>Draftspect tray app isn't running</strong> — start it, then reopen this panel.</li>
  <li>Word/Excel cached an old add-in — <strong>fully quit the app (⌘Q / Alt+F4) and reopen it</strong> so it re-reads the add-in.</li>
  <li>This add-in's manifest points at a <strong>different port</strong> than the running Draftspect daemon (e.g. another add-in's daemon answered). Relaunch Draftspect, then quit &amp; reopen Word/Excel.</li>
  <li>If it persists, reinstall the add-in from the Draftspect tray menu.</li>
</ol>
<p class="foot">Draftspect daemon on <code>127.0.0.1:${HTTP_PORT}</code> · requested <code>${requestedPath}</code> · ${status}</p>
</body></html>`;
}

function sendError(req, res, status, headline, requestedPath) {
  const wantsHtml = (req.headers.accept || "").includes("text/html");
  if (wantsHtml) {
    res.writeHead(status, { "Content-Type": "text/html; charset=utf-8" });
    res.end(errorPageHtml(status, headline, requestedPath));
  } else {
    res.writeHead(status).end(headline);
  }
}

const http = createLocalServer({
  port: HTTP_PORT,
  tls: localTransport.tls,
  handler: async (req, res) => {
    // createLocalServer rejects a foreign Host or explicit Origin before this
    // handler runs. Same-origin GETs may omit Origin; explicit matching origins
    // receive only the exact taskpane origin, never a wildcard CORS grant.
    const reqOrigin = req.headers.origin || "";
    if (reqOrigin === HTTP_ORIGIN) {
      res.setHeader("Access-Control-Allow-Origin", reqOrigin);
    }
    // No-cache: Word's webview likes to cache aggressively. During dev we want
    // every reload to pick up the latest taskpane JS/CSS/HTML.
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    if (req.method === "OPTIONS") {
      res.writeHead(204).end();
      return;
    }

    try {
      const urlPath = (req.url || "/").split("?")[0];

      // Token endpoint: serves the per-daemon bridge token to the taskpane.
      // Host and any supplied Origin are checked before this handler runs.
      if (urlPath === "/bridge-token") {
        res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
        res.end(BRIDGE_TOKEN);
        return;
      }

      const relPath = urlPath === "/" ? "/index.html" : urlPath;
      const fsPath = join(taskpaneDir, relPath);
      // Containment check. `join` already normalizes `../`, so the obvious
      // traversal is blocked — but a bare startsWith(taskpaneDir) would also
      // accept a sibling like `<…>/taskpane-evil/x`. Require an exact match
      // OR a path under `taskpaneDir` + separator.
      if (fsPath !== taskpaneDir && !fsPath.startsWith(taskpaneDir + sep)) {
        res.writeHead(403).end("Forbidden");
        return;
      }
      const ext = extname(fsPath).toLowerCase();
      const mime = MIME[ext] || "application/octet-stream";
      if (ext === ".html") {
        // Office's webview caches the JS/CSS bundle aggressively and ignores
        // our no-store headers — so a taskpane code change wouldn't take
        // effect even after reopening the pane. Cache-bust the local asset
        // refs with the per-start bridge token: every daemon restart yields a
        // fresh URL, forcing a re-fetch. (The handler strips the query string
        // before resolving the file, so `?v=` doesn't affect routing.)
        let html = await readFile(fsPath, "utf8");
        html = html.replace(/(\/shared\/(?:taskpane\.js|styles\.css))"/g, `$1?v=${BRIDGE_TOKEN}"`);
        res.writeHead(200, { "Content-Type": mime });
        res.end(html);
      } else if (ext === ".js") {
        // Bust transitive module imports too: a new pane must not import an old
        // cached tools-word.js that lacks the matching protocol implementation.
        const js = (await readFile(fsPath, "utf8")).replace(
          /(from\s*["'])(\.[^"'?]*\.(?:js|mjs))(["'])/g,
          `$1$2?v=${BRIDGE_TOKEN}$3`,
        );
        res.writeHead(200, { "Content-Type": mime });
        res.end(js);
      } else {
        const data = await readFile(fsPath);
        res.writeHead(200, { "Content-Type": mime });
        res.end(data);
      }
    } catch (err) {
      const reqPath = (req.url || "/").split("?")[0];
      if (err.code === "ENOENT") {
        sendError(
          req,
          res,
          404,
          "That page or file isn’t served by this Draftspect daemon.",
          reqPath,
        );
      } else {
        console.error("[http]", err);
        sendError(
          req,
          res,
          500,
          "The Draftspect daemon hit an internal error serving this page.",
          reqPath,
        );
      }
    }
  },
});

http.ready.then(
  () => {
    console.log(
      `[daemon] ${localTransport.secure ? "HTTPS" : "HTTP"} server listening on ${HTTP_ORIGIN}/`,
    );
  },
  () => {},
);

// ---------------------------------------------------------------------------
// IPC channel back to the Electron main process. Used to ask main.mjs to
// show a native macOS open panel for folder/file picking — synthetic
// in-page modals can't navigate to Google Drive, iCloud, recent items, or
// any of the other sources macOS users expect in NSOpenPanel. The channel
// is fd 3 (added in main.mjs's spawn options), wired through Node IPC.
// ---------------------------------------------------------------------------
const pendingPicks = new Map(); // id -> { resolve, reject, timer }
const PICK_TIMEOUT_MS = 5 * 60_000; // 5 min; the user might leave the dialog open
let daemonHealthSnapshot = () => ({ phase: "starting" });

if (process.send) {
  process.on("message", (msg) => {
    if (msg?.type === "health_check") {
      process.send({ type: "health_result", id: msg.id, ...daemonHealthSnapshot() });
      return;
    }
    if (msg?.type !== "pick_path_result") return;
    const entry = pendingPicks.get(msg.id);
    if (!entry) return;
    pendingPicks.delete(msg.id);
    clearTimeout(entry.timer);
    entry.resolve(msg);
  });
}

function pickPathFromMain({ include_files, default_path, title, button_label }) {
  if (!process.send) {
    return Promise.reject(new Error("No IPC channel to Electron main process"));
  }
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingPicks.delete(id);
      reject(new Error("Picker timed out"));
    }, PICK_TIMEOUT_MS);
    pendingPicks.set(id, { resolve, reject, timer });
    process.send({ type: "pick_path", id, include_files, default_path, title, button_label });
  });
}

// Word and Excel are fully independent surfaces: each host gets its own
// agent loop, session, message queue, transcript AND workspace. Nothing
// one host does ever touches the other (no shared "current session", no
// cross-host interrupt). Both maps are hoisted above createBridge so
// bridge handlers that fire during the top-level awaits below
// (preflightHttpMcpServers, etc.) don't hit a TDZ on these bindings.
//
//   sessions:        paneKey -> live session { key, host, cwd, sessionId,
//                    cancel, close, settled, interrupted }
//   workspaceByKey:  paneKey -> last-known cwd, so a pane that connects
//                    before its first message still resolves a workspace
//                    (and survives that pane's loop ending).
//
// paneKey identifies one open document (host + doc path), so two Word
// docs (or two workbooks) open at once each get their own independent
// loop/session/queue/workspace. The host is still carried for the tool
// family + per-host behavior. Persisted SDK resume and serial handoff state
// are scoped by provider, workspace, host AND document identity.
const sessions = new Map();
const workspaceByKey = new Map();
// Panes whose workspace the user set explicitly (Change workspace). Their
// workspace must NOT be re-derived from the document folder on reconnect.
const explicitWorkspaceKeys = new Set();
// Per-pane sticky model choice (set by the taskpane composer). The SDK
// model is fixed per agent loop, so a change while a loop is live triggers
// a resuming restart (see the set_model handler).
const modelByKey = new Map();
const handoffInProgress = new Set();
const runtimeClosures = createRuntimeCloseRegistry();
const sessionPersistence = createSessionPersistenceRegistry({
  save: saveSessionId,
  notify: (event, key) => bridge.sendAssistantEvent(event, key),
});

async function sessionInfo(key, host) {
  const cwd = cwdForKey(key);
  const session_id = await resolveReplaySessionId(key, host, cwd);
  const handoff = await getHandoff(host, cwd, runtime.id, key);
  const persistence = sessionPersistence.status(key, { cwd, provider: runtime.id, host });
  return {
    provider: runtime.id,
    host,
    cwd,
    session_id,
    persistence_status: persistence?.status || "saved_or_none",
    persistence_error: persistence?.error || null,
    runtime_processes: runtimeClosures.status(key),
    state: handoff?.state || (sessionFor(key)?.lifecycle.active ? "working" : "word"),
    archive_path: handoff?.archive_path ?? null,
    return_confirmation: handoff?.return_confirmation?.id ?? null,
    commands:
      session_id &&
      (!persistence || persistence.status === "saved") &&
      handoff?.state === "released" &&
      handoff?.archive_path
        ? resumeCommands({ provider: runtime.id, session_id, cwd })
        : null,
  };
}
// Static fallback for the experimental Claude runtime (standard aliases) when the
// pane's stored choice is missing or stale. CodeBuddy ids come from the
// discovered catalog instead — see getRuntimeModels().
const DEFAULT_CLAUDE_MODEL = "sonnet";

// Discovered model catalog (CodeBuddy runtime), cached with a TTL and
// in-flight dedup. The taskpane's dropdown and set_model validation read
// this; the Claude runtime serves its static catalog through the same
// listModels() path (instant).
const RUNTIME_MODELS_TTL_MS = 10 * 60 * 1000;
let runtimeModelsCache = null; // { at, list }
let runtimeModelsInflight = null;

async function getRuntimeModels({ force = false } = {}) {
  const now = Date.now();
  if (!force && runtimeModelsCache && now - runtimeModelsCache.at < RUNTIME_MODELS_TTL_MS) {
    return runtimeModelsCache.list;
  }
  if (runtimeModelsInflight) return runtimeModelsInflight;
  runtimeModelsInflight = (async () => {
    try {
      const raw = await runtime.listModels?.({ force: true });
      const list = normalizeModelList(raw);
      if (list) runtimeModelsCache = { at: Date.now(), list };
      if (!list) runtimeModelsCache = null;
      return list;
    } catch (err) {
      console.warn(`[daemon] listModels failed: ${err?.message ?? err}`);
      runtimeModelsCache = null;
      return null;
    } finally {
      runtimeModelsInflight = null;
    }
  })();
  return runtimeModelsInflight;
}

// The model to pass to a session. Runtime-aware: the Claude runtime pins an
// explicit legacy provider id (cost-control UI; never silently inherits the CLI
// default); the CodeBuddy runtime validates against the discovered catalog
// and otherwise lets the runtime default apply. A stale id from the OTHER
// runtime is dropped explicitly (logged), never silently mis-sent.
function modelArgFor(key) {
  const requested = modelByKey.get(key);
  const { model, dropped } = pickModel({
    requested,
    runtimeId: runtime.id,
    knownModels: runtime.knownModels,
    dynamicModels: runtimeModelsCache?.list ?? null,
    fallback: DEFAULT_CLAUDE_MODEL,
  });
  if (dropped) {
    throw new Error(
      `Model "${requested}" is not currently validated for ${runtime.id}. Refresh the catalog and select a model.`,
    );
  }
  if (runtime.id === "codebuddy" && !model)
    throw new Error("Select and confirm a CodeBuddy model before sending a request.");
  return model;
}

function sessionFor(key) {
  return sessions.get(key) ?? null;
}

// This pane's workspace: its live session's cwd, else the last cwd we
// recorded for it, else the launch default.
function cwdForKey(key) {
  return sessionFor(key)?.cwd ?? workspaceByKey.get(key) ?? matterFolder;
}

// Resolve the session id to replay for this pane. Prefer the pane's live
// session id (set once the SDK reports init); otherwise the id persisted
// for (host, cwd) — covers the window before the SDK has re-inited.
async function resolveReplaySessionId(key, host, cwd) {
  const recorded = sessionPersistence.status(key, { cwd, provider: runtime.id, host });
  if (recorded) return recorded.session_id;
  const live = sessionFor(key);
  if (live?.sessionId) return live.sessionId;
  if (!host || !cwd) return null;
  try {
    return (await getSessionId(host, cwd, runtime.id, key)) ?? null;
  } catch {
    return null;
  }
}

// Reconstruct this pane's prior conversation from its .jsonl and push it
// to THAT pane only. Sent on every taskpane hello and after a workspace
// switch (cwd_changed). Empty events => fresh chat, no divider.
async function sendTranscriptReplayTo(key, host, cwd) {
  try {
    const sessionId = await resolveReplaySessionId(key, host, cwd);
    const { events, truncated } = sessionId
      ? await readTranscript(sessionId, { maxEvents: 200, provider: runtime.id })
      : { events: [], truncated: false };
    bridge.sendToTaskpane(
      {
        type: "transcript_replay",
        session_id: sessionId ?? null,
        session_info: await sessionInfo(key, host),
        truncated,
        events,
        legacy_history_available: !sessionId && !!(await getSessionId(host, cwd, runtime.id)),
      },
      key,
    );
  } catch (err) {
    console.warn("[daemon] transcript replay failed:", err?.message ?? err);
    bridge.sendToTaskpane(
      {
        type: "transcript_replay",
        events: [],
        truncated: false,
        error: "Prior history could not be read: " + (err?.message ?? err),
      },
      key,
    );
  }
}

// Start (or resume) this pane's session on the next tick — past the
// current message handler / the agent loop's finally block (which clears
// sessions.get(key)), so startSessionForFolder builds cleanly. Every
// path that (re)starts a pane's session — first message, post-stream-end,
// post-Stop — funnels through here; `reason` only flavors the failure log.
// Per-key serialization + coalescing of (re)starts. Every restart trigger —
// first message, post-Stop, post-stream-end, model/context change, workspace
// switch — funnels through here. startSessionForFolder aborts whatever
// session is current for the key, so if several triggers fire close together
// (e.g. "slow first turn → Stop → switch model → resend") their setImmediate
// callbacks used to interleave: each abort killed the next start before it
// could `init`, leaving an orphaned loop with no consumer (the Excel hang).
//
// Now each key has a single-flight runner. While a start is in progress,
// later requests only overwrite `latest` — intermediate restarts collapse,
// and exactly one final start runs to completion uncontested, so it always
// reaches `init`.
const startQueue = new Map(); // key -> { running, latest, generation }

function scheduleSessionStart(cwd, sessionId, key, host, reason, { replay = true } = {}) {
  let q = startQueue.get(key);
  if (!q) {
    q = { running: false, latest: null, generation: 0 };
    startQueue.set(key, q);
  }
  q.latest = {
    generation: ++q.generation,
    cwd,
    resumeId: sessionId,
    host: host ?? null,
    reason,
    replay,
    turnId: bridge.activeTurn(key),
  };
  if (q.running) return; // the active runner will pick up `latest`
  q.running = true;
  setImmediate(() => runStartQueue(key));
}

async function runStartQueue(key) {
  const q = startQueue.get(key);
  if (!q) return;
  // startSessionForFolder returns once setup is done (it kicks the agent
  // loop off detached), so awaiting it serializes only the abort+create
  // step — exactly the part that must not interleave.
  while (startQueue.get(key) === q && q.latest) {
    const req = q.latest;
    q.latest = null;
    try {
      await startSessionForFolder(req.cwd, req.resumeId, {
        key,
        host: req.host,
        replay: req.replay,
        turnId: req.turnId,
        isCurrentStart: () => startQueue.get(key) === q && q.generation === req.generation,
      });
    } catch (err) {
      console.error(`[daemon] ${req.reason} session start failed:`, err?.message ?? err);
    }
  }
  q.running = false;
}

// Called on every taskpane hello — for ANY open document, many panes
// possibly connected at once. Connecting a pane must NOT start the agent
// loop: connect-driven starts amplified the old connect/disconnect
// ping-pong and burn an unasked turn. We only re-render this pane's own
// transcript. The loop starts lazily on the first user message
// (onUserMessage → ensureLoopForMessage).
async function onPaneConnect(key, host, doc) {
  if (!key) return;
  const reconnecting = sessionFor(key);
  reconnecting?.cancel();
  // Resolve this pane's workspace from the open document's own folder,
  // server-side and immediately — deterministic, no loop start. Without
  // this, cwdForKey() falls back to matterFolder (the daemon's launch
  // cwd, often the repo root) until a message or the taskpane's set_cwd
  // round-trip lands, so get_context / replay read the WRONG folder's
  // CLAUDE.md (the symptom: context files bleeding across documents).
  // Skip if the user explicitly pinned this pane's workspace, or a live
  // session already owns the cwd.
  if (doc && !explicitWorkspaceKeys.has(key) && !sessionFor(key)) {
    try {
      const folder = await resolveWorkspaceRoot(doc);
      if (folder) workspaceByKey.set(key, folder);
    } catch {
      /* unresolvable (cloud/unsaved) — keep the fallback */
    }
  }
  const cwd = cwdForKey(key);
  diag(`hello → replay key=${key} cwd=${cwd} (no loop start on connect)`);
  await sendTranscriptReplayTo(key, host, cwd);
  if (reconnecting)
    bridge.sendAssistantEvent(
      {
        event: "info",
        message:
          "The previous connection's running request was stopped. Inspect any in-flight edits before resuming.",
      },
      key,
    );
}

// Called when a pane's WebSocket closes for good (the bridge already
// freed its own pane/queue state). Prune the daemon-side per-key Maps so
// they don't grow for the life of the daemon — but ONLY when no live
// session owns this key. A session is deliberately kept alive across a
// transient disconnect (it resumes on reconnect with the same stable
// key); its own `finally` clears `sessions` when its loop actually ends.
function onPaneClose(key) {
  if (!key) return;
  sessionFor(key)?.cancel();
  const turnId = bridge.activeTurn(key);
  if (turnId)
    bridge.sendAssistantEvent(
      { event: "turn_complete", turn_id: turnId, subtype: "disconnected", interrupted: true },
      key,
    );
  bridge.clearUserMessages(key);
  startQueue.delete(key);
  // Preserve lightweight workspace/model choice through a short reconnect.
  const rememberedWorkspace = workspaceByKey.get(key);
  const expiry = setTimeout(() => {
    if (
      !bridge.isTaskpaneConnected(key) &&
      !sessionFor(key) &&
      workspaceByKey.get(key) === rememberedWorkspace
    ) {
      workspaceByKey.delete(key);
      explicitWorkspaceKeys.delete(key);
      modelByKey.delete(key);
    }
  }, 10 * 60_000);
  expiry.unref?.();
}

// Called when a user message arrives from a pane, BEFORE it's queued.
// Each pane has its OWN loop — independent of every other pane. If this
// pane's loop is already live, do nothing (its turn source will
// consume the message). Otherwise start it, resuming this (host, cwd)
// conversation. Deferred via setImmediate so it lands after the message
// is queued and after any in-flight finally; the new loop then drains
// this pane's queue. No other pane's loop is ever touched.
async function ensureLoopForMessage(key, host) {
  if (!key) return;
  const turnId = bridge.activeTurn(key);
  const cwd = cwdForKey(key);
  if (
    handoffInProgress.has(key) ||
    (await getHandoff(host, cwd, runtime.id, key))?.state === "released"
  )
    throw new Error(
      "This session was released for CLI/archive use. Exit the CLI and select Return to Word before continuing.",
    );
  const live = sessionFor(key);
  if (live && !live.settled) {
    return; // this pane's loop is live and will consume the message
  }
  // Stop can remove the live loop before its latest ID reaches disk. Apply
  // the same persistence barrier as CLI handoff; never resume an older ID.
  const persisted = await sessionPersistence.awaitSaved(key, {
    host,
    cwd,
    provider: runtime.id,
  });
  const resumeId = persisted?.session_id ?? (await getSessionId(host, cwd, runtime.id, key));
  if (bridge.activeTurn(key) !== turnId) return;
  diag(`message → ensure loop key=${key} cwd=${cwd} resume=${resumeId ?? "(new)"}`);
  // replay:false — the pane already shows the chat (incl. the message
  // that just triggered this). An empty transcript_replay here (a
  // brand-new session has no .jsonl yet) would wipe the user's prompt.
  scheduleSessionStart(cwd, resumeId, key, host, "user message", { replay: false });
}

// ---------------------------------------------------------------------------
// WebSocket bridge.
// ---------------------------------------------------------------------------
const bridge = createBridge({
  port: WS_PORT,
  tls: localTransport.tls,
  runtimeId: runtime.id,
  token: BRIDGE_TOKEN,
  allowedOrigins: [HTTP_ORIGIN],
  allowHello: acceptanceProfile ? acceptanceAllowsDocument : undefined,
  onHello: (key, host, doc) => onPaneConnect(key, host, doc),
  onUserMessage: (key, host) => ensureLoopForMessage(key, host),
  onClose: (key) => onPaneClose(key),
  // Office tool events are turn activity too: a long confirmation wait or a
  // slow Office.js write produces no SDK messages, but the turn is healthy.
  onActivity: (key) => sessionFor(key)?.lifecycle?.activity?.(),
  extraHandlers: {
    pick_path: async (msg, reply) => {
      if (acceptanceProfile) throw new Error("Acceptance does not open arbitrary path pickers");
      try {
        const result = await pickPathFromMain({
          include_files: !!msg.include_files,
          default_path: msg.default_path || null,
          title: msg.title || null,
          button_label: msg.button_label || null,
        });
        reply({ type: "pick_path_result", request_id: msg.request_id, ...result });
      } catch (e) {
        reply({
          type: "pick_path_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_cwd: async (msg, reply, key, host) => {
      try {
        let cwd;
        let explicitPick = false;
        if (msg.autodetect_from_doc) {
          const detected = await resolveWorkspaceRoot(msg.autodetect_from_doc);
          if (!detected)
            throw new Error("Could not auto-detect a workspace folder from that doc path");
          cwd = detected;
        } else if (msg.cwd) {
          cwd = msg.cwd;
          explicitPick = true;
        } else {
          throw new Error("set_cwd requires `cwd` or `autodetect_from_doc`");
        }
        assertAcceptanceWorkspace(cwd);
        // Remember an explicit pick so a later reconnect doesn't re-derive
        // this pane's workspace from the doc folder; an autodetect switch
        // clears that pin (the doc folder is authoritative again).
        if (explicitPick) explicitWorkspaceKeys.add(key);
        else explicitWorkspaceKeys.delete(key);
        const resolved = await switchFolder(cwd, key, host);
        // Drop a CLAUDE.md marker on explicit user pick so the next open of
        // any doc in this folder auto-detects silently.
        let markerCreated = false;
        if (explicitPick && !acceptanceProfile) {
          try {
            markerCreated = await ensureWorkspaceMarker(resolved);
          } catch (e) {
            console.warn(`[daemon] could not create CLAUDE.md in ${resolved}: ${e.message}`);
          }
        }
        reply({
          type: "set_cwd_result",
          ok: true,
          cwd: resolved,
          marker_created: markerCreated,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({ type: "set_cwd_result", ok: false, error: e.message, request_id: msg.request_id });
      }
    },
    suggest_workspace: async (msg, reply) => {
      if (acceptanceProfile) {
        reply({
          type: "suggest_workspace_result",
          ok: true,
          suggestion: { cwd: acceptanceProfile.workspace, confidence: "doc" },
          request_id: msg.request_id,
        });
        return;
      }
      try {
        const suggestion = await suggestWorkspaceRoot(msg.doc_path || null);
        reply({
          type: "suggest_workspace_result",
          ok: true,
          suggestion,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({
          type: "suggest_workspace_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_cwd_state: async (msg, reply, key) => {
      reply({
        type: "get_cwd_state_result",
        ok: true,
        // This pane's own workspace, resolvable even before its first
        // message (lazy start ⇒ no session yet).
        current_cwd: cwdForKey(key),
        request_id: msg.request_id,
      });
    },
    stop_agent: async (msg, reply, key) => {
      // User clicked Stop in this pane. Cancel only THIS pane's loop; the
      // runtime's terminal-state path (Claude: AbortError; CodeBuddy: a
      // result consumed with the interrupted flag) emits turn_complete
      // interrupted so the taskpane flips to Ready. Every other pane's
      // loop is untouched.
      const s = sessionFor(key);
      const turnId = bridge.activeTurn(key);
      if (s || turnId) {
        const q = startQueue.get(key);
        if (q) q.latest = null;
        if (s) s.cancel();
        else {
          bridge.clearUserMessages(key);
          bridge.cancelTools(key);
          bridge.sendAssistantEvent(
            { event: "turn_complete", turn_id: turnId, subtype: "interrupted", interrupted: true },
            key,
          );
        }
        reply({ type: "stop_agent_result", ok: true, request_id: msg.request_id });
      } else {
        reply({
          type: "stop_agent_result",
          ok: false,
          error: "No active agent turn",
          request_id: msg.request_id,
        });
      }
    },
    get_session_info: async (msg, reply, key, host) => {
      reply({
        type: "get_session_info_result",
        ok: true,
        request_id: msg.request_id,
        info: await sessionInfo(key, host),
      });
    },
    retry_session_save: async (msg, reply, key, host) => {
      if (handoffInProgress.has(key)) throw new Error("Session handoff is in progress.");
      const cwd = cwdForKey(key);
      await sessionPersistence.retry(key, { cwd, provider: runtime.id, host });
      reply({
        type: "retry_session_save_result",
        request_id: msg.request_id,
        ok: true,
        info: await sessionInfo(key, host),
      });
    },
    recover_runtime: async (msg, reply, key) => {
      if (msg.confirm !== true || bridge.activeTurn(key) || handoffInProgress.has(key))
        throw new Error(
          "Stop the request and explicitly confirm recovery of the displayed process first.",
        );
      handoffInProgress.add(key);
      try {
        const result = await runtimeClosures.recover(key, {
          expectedPid: msg.expected_pid,
          signal: msg.force === true ? "SIGKILL" : "SIGTERM",
          confirmForce: msg.force === true && msg.confirm_force === true,
        });
        reply({ type: "recover_runtime_result", request_id: msg.request_id, ok: true, ...result });
      } finally {
        handoffInProgress.delete(key);
      }
    },
    release_session: async (msg, reply, key, host) => {
      if (handoffInProgress.has(key)) throw new Error("A session handoff is already running");
      if (bridge.activeTurn(key))
        throw new Error("Finish or stop the current turn before archiving");
      handoffInProgress.add(key);
      try {
        await sessionPersistence.awaitSaved(key, {
          cwd: cwdForKey(key),
          provider: runtime.id,
          host,
        });
        const info = await sessionInfo(key, host);
        if (!info.session_id) throw new Error("No persisted session yet");
        const s = sessionFor(key);
        await setHandoff(host, info.cwd, runtime.id, key, {
          state: "released",
          session_id: info.session_id,
          released_at: new Date().toISOString(),
        });
        const queued = startQueue.get(key);
        if (queued) queued.latest = null;
        s?.cancel();
        if (runtime.id === "codebuddy") await runtimeClosures.wait(key);
        else if (s?.done) {
          let timer;
          try {
            await Promise.race([
              s.done,
              new Promise((_, reject) => {
                timer = setTimeout(
                  () =>
                    reject(
                      new Error(
                        "Runtime close not confirmed; session remains released. Retry after it exits.",
                      ),
                    ),
                  5000,
                );
              }),
            ]);
          } finally {
            clearTimeout(timer);
          }
        }
        // Preserve the baseline even if archiving later fails. The owned SDK
        // process must have exited before observing an external handoff file.
        const released = {
          state: "released",
          session_id: info.session_id,
          released_at: new Date().toISOString(),
          transcript_fingerprint: await captureTranscriptFingerprint(info),
        };
        await setHandoff(host, info.cwd, runtime.id, key, released);
        const archive = await archiveSession(info);
        await setHandoff(host, info.cwd, runtime.id, key, {
          ...released,
          released_at: archive.created_at,
          archive_path: archive.archive_path,
        });
        reply({
          type: "release_session_result",
          ok: true,
          request_id: msg.request_id,
          info: await sessionInfo(key, host),
        });
      } finally {
        handoffInProgress.delete(key);
      }
    },
    return_session: async (msg, reply, key, host) => {
      if (msg.cli_closed !== true) throw new Error("Confirm the CLI conversation has exited first");
      if (handoffInProgress.has(key) || bridge.activeTurn(key)) throw new Error("Session is busy");
      handoffInProgress.add(key);
      try {
        const info = await sessionInfo(key, host);
        if (!info.session_id || msg.session_id !== info.session_id)
          throw new Error("Session identity changed; refresh first");
        if (runtime.id === "codebuddy") await runtimeClosures.wait(key);
        const handoff = await getHandoff(host, info.cwd, runtime.id, key);
        if (handoff?.state !== "released" || handoff.session_id !== info.session_id)
          throw new Error("Session is not released; refresh its handoff state first");
        const latest = await captureTranscriptFingerprint(info);
        const confirmation = handoff.return_confirmation;
        const confirmedChange =
          confirmation &&
          msg.confirm_transcript_change === confirmation.id &&
          sameTranscriptFingerprint(latest, confirmation.fingerprint);
        if (
          !confirmedChange &&
          (confirmation || !sameTranscriptFingerprint(latest, handoff.transcript_fingerprint))
        ) {
          const pending =
            confirmation && sameTranscriptFingerprint(latest, confirmation.fingerprint)
              ? confirmation
              : { id: randomUUID(), fingerprint: latest };
          await setHandoff(host, info.cwd, runtime.id, key, {
            ...handoff,
            return_confirmation: pending,
          });
          reply({
            type: "return_session_result",
            request_id: msg.request_id,
            ok: false,
            code: "TRANSCRIPT_CHANGED",
            error:
              "TRANSCRIPT_CHANGED: session history was externally modified or its release snapshot is unavailable. Return was refused. Exit the CLI, inspect the history change, then explicitly confirm again; no operations were replayed.",
            info: await sessionInfo(key, host),
          });
          return;
        }
        await setHandoff(host, info.cwd, runtime.id, key, null);
        await sendTranscriptReplayTo(key, host, info.cwd);
        reply({
          type: "return_session_result",
          ok: true,
          request_id: msg.request_id,
          info: await sessionInfo(key, host),
        });
      } finally {
        handoffInProgress.delete(key);
      }
    },
    get_context: async (msg, reply, key) => {
      try {
        const cwd = cwdForKey(key);
        const entries = !acceptanceProfile && cwd ? await getContextEntries(cwd) : [];
        reply({ type: "get_context_result", ok: true, cwd, entries, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "get_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    list_models: async (msg, reply) => {
      // Serve the ACTIVE runtime's model catalog to the taskpane dropdown.
      // Claude: the fixed legacy provider list. CodeBuddy: the discovered catalog
      // (cached; a cold-cache first call may take a few seconds).
      const list = await getRuntimeModels();
      reply({
        type: "list_models_result",
        ok: !!list,
        runtime: runtime.id,
        models: list ?? [],
        request_id: msg.request_id,
      });
    },
    set_model: async (msg, reply, key, host) => {
      if (bridge.activeTurn(key))
        throw new Error("Stop or finish the current request before changing models.");
      const requested = String(msg.model || "").trim();
      if (!(await getRuntimeModels())?.some((m) => m.value === requested))
        throw new Error(
          "Model is not in the current catalog. Refresh and choose an available model.",
        );
      let model;
      if (runtime.id === "claude") {
        // Unchanged behavior for the default runtime: unknown ids migrate
        // to the static default so existing panes keep working across
        // provider model retirements.
        model = (runtime.knownModels ?? []).includes(requested) ? requested : DEFAULT_CLAUDE_MODEL;
      } else {
        // Discovery-based runtime: unknown ids are REJECTED explicitly
        // (plan §2 — no silent fallback) and the current model is kept.
        if (!requested) {
          reply({
            type: "set_model_result",
            ok: false,
            error: "No model specified.",
            model: modelByKey.get(key) ?? "",
            request_id: msg.request_id,
          });
          return;
        }
        const list = await getRuntimeModels();
        if (!list || !list.some((m) => m.value === requested)) {
          reply({
            type: "set_model_result",
            ok: false,
            error: `"${requested}" is not offered by the ${runtime.id} runtime. Pick a model from the dropdown list.`,
            model: modelByKey.get(key) ?? "",
            request_id: msg.request_id,
          });
          return;
        }
        model = requested;
      }
      const prev = modelByKey.get(key);
      modelByKey.set(key, model);
      reply({ type: "set_model_result", ok: true, model, request_id: msg.request_id });
      // Only relaunch if the model actually changed for an already-known
      // pane. On the initial connect `prev` is undefined (no live session
      // yet — the lazy first-message start will read modelByKey), so we
      // just record it. restartSession itself no-ops when no session.
      if (prev !== undefined && prev !== model) {
        console.log(`[daemon] model changed → ${model} (key bound); restarting loop`);
        restartSession(key, host, { reason: "model_changed" }).catch((err) =>
          console.warn("[daemon] restart after model change failed:", err.message),
        );
      }
    },
    set_context: async (msg, reply, key, host) => {
      if (acceptanceProfile) throw new Error("Acceptance does not change context files");
      try {
        const cwd = cwdForKey(key);
        if (!cwd) throw new Error("No workspace selected");
        const { saved, errors } = await setContextEntries(cwd, msg.entries || []);
        reply({
          type: "set_context_result",
          ok: errors.length === 0,
          cwd,
          saved,
          errors,
          request_id: msg.request_id,
        });
        // Restart THIS pane's loop so its agent re-reads CLAUDE.md and
        // picks up the updated context block on the next turn. Every
        // other pane is unaffected.
        restartSession(key, host, { reason: "context_changed" }).catch((err) =>
          console.warn("[daemon] restart failed:", err.message),
        );
      } catch (e) {
        reply({
          type: "set_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
  },
});

// ---------------------------------------------------------------------------
// Office-bridge MCP server (in-process; forwards tool calls to the taskpane).
// Built per session so the registered tool family matches the connected
// host (see startSessionForFolder / the host re-narrow on hello).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Load the user's MCP servers from ~/.claude.json — the same file the
// interactive `claude` CLI uses. The Agent SDK doesn't read this file by
// default (it reads `~/.claude/settings.json` instead), so without this step
// the user's configured servers (visio, etc.) would be invisible to the
// daemon.
// ---------------------------------------------------------------------------
async function loadUserMcpServers() {
  if (acceptanceProfile || !forwardUserMcp()) return {};
  const configPath = join(homedir(), ".claude.json");
  try {
    const raw = await readFile(configPath, "utf8");
    const parsed = JSON.parse(raw);
    const servers = parsed?.mcpServers ?? {};
    return servers;
  } catch (err) {
    if (err.code === "ENOENT") return {};
    console.warn(`[daemon] Could not load MCP servers from ${configPath}:`, err.message);
    return {};
  }
}

// Hoisted ABOVE the top-level awaits below on purpose. Module evaluation
// suspends at the first top-level `await` (loadUserMcpServers /
// preflightHttpMcpServers). The bridge WS server is already listening by
// then, so a taskpane `hello` can arrive mid-suspension and drive
// startSessionForFolder before the rest of the module body runs. Anything
// that path touches must be initialized first, or it hits a TDZ
// ReferenceError. Keep this (and any other start-path consts) up here.
//
// Disallow the out-of-process word-mcp tools. The user has word-mcp configured
// in ~/.claude (it surfaces as `mcp__word-mcp__*`), but those tools drive Word
// via AppleScript/COM and cause screen flicker on every edit — disqualifying
// for the live-edit path. We replace them with the in-process office_* tools
// that go through Office.js. The SDK's disallowedTools doesn't support globs,
// so we have to enumerate.
const WORD_MCP_DISALLOWED = [
  "mcp__word-mcp__word_accept_revisions",
  "mcp__word-mcp__word_add_comment",
  "mcp__word-mcp__word_apply_style",
  "mcp__word-mcp__word_begin_transaction",
  "mcp__word-mcp__word_commit_transaction",
  "mcp__word-mcp__word_delete_comment",
  "mcp__word-mcp__word_delete_paragraphs",
  "mcp__word-mcp__word_delete_snapshot",
  "mcp__word-mcp__word_diff_snapshots",
  "mcp__word-mcp__word_emergency_recover",
  "mcp__word-mcp__word_export_pdf",
  "mcp__word-mcp__word_find_text",
  "mcp__word-mcp__word_get_document_info",
  "mcp__word-mcp__word_get_outline",
  "mcp__word-mcp__word_get_paragraph",
  "mcp__word-mcp__word_get_paragraphs",
  "mcp__word-mcp__word_get_section",
  "mcp__word-mcp__word_get_selection",
  "mcp__word-mcp__word_get_styles",
  "mcp__word-mcp__word_insert_paragraphs",
  "mcp__word-mcp__word_list_comments",
  "mcp__word-mcp__word_list_open_documents",
  "mcp__word-mcp__word_list_revisions",
  "mcp__word-mcp__word_list_transactions",
  "mcp__word-mcp__word_prune_snapshots",
  "mcp__word-mcp__word_reject_revisions",
  "mcp__word-mcp__word_replace_paragraphs",
  "mcp__word-mcp__word_replace_range",
  "mcp__word-mcp__word_replace_section",
  "mcp__word-mcp__word_replace_text",
  "mcp__word-mcp__word_restore_snapshot",
  "mcp__word-mcp__word_rollback_transaction",
  "mcp__word-mcp__word_toggle_track_changes",
  "mcp__word-mcp__word_undo_last_edit",
];

const userMcpServers = await loadUserMcpServers();
const userMcpNames = Object.keys(userMcpServers);
if (userMcpNames.length > 0) {
  console.log(
    `[daemon] Loaded ${userMcpNames.length} MCP server(s) from ~/.claude.json: ${userMcpNames.join(", ")}`,
  );
}

// Preflight HTTP MCP servers. The SDK will silently drop any server whose
// initial handshake fails, with no retry for the lifetime of the session
// (see memory: sdk-silently-drops-failed-mcp). We can't fix that here, but
// we can surface the failure in the daemon log so it's obvious why a tool
// is missing — "restart the daemon when the server is back up" instead of
// "no idea why Visio doesn't work".
async function preflightHttpMcpServers(servers) {
  const entries = Object.entries(servers).filter(([, cfg]) => cfg?.type === "http" && cfg.url);
  await Promise.all(
    entries.map(async ([name, cfg]) => {
      const body = JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "claude-code-office-preflight", version: "0.1" },
        },
      });
      try {
        const res = await fetch(cfg.url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json, text/event-stream",
          },
          body,
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          console.log(`[daemon] MCP preflight: ${name} reachable at ${cfg.url}`);
        } else {
          console.warn(
            `[daemon] MCP preflight: ${name} returned HTTP ${res.status} — tools may be missing until daemon restart`,
          );
        }
      } catch (err) {
        const reason =
          err.name === "TimeoutError" ? "timeout after 5s" : err.cause?.code || err.message;
        console.warn(
          `[daemon] MCP preflight: ${name} unreachable (${reason}) — tools will be missing until daemon restart`,
        );
      }
    }),
  );
}
await preflightHttpMcpServers(userMcpServers);

// ---------------------------------------------------------------------------
// System prompt: Claude Code default + Office-specific append.
// ---------------------------------------------------------------------------
// Shared base + the active host's section only. An Excel session never
// carries Word's surgical-edit / paragraph-reference rules (pure noise
// for it) and vice-versa. Re-read fresh on every session start so edits
// take effect when a session restarts (no daemon restart required).
async function buildSystemPromptAppend(host) {
  const base = await readFile(join(__dirname, "system-prompt.md"), "utf8");
  const files = host === "excel" ? ["system-prompt-excel.md"] : ["system-prompt-word.md"];
  // Degraded pre-bind case (host null) shouldn't happen now that sessions
  // start per pane, but fall back to including both to stay safe.
  if (host !== "word" && host !== "excel") {
    files.length = 0;
    files.push("system-prompt-word.md", "system-prompt-excel.md");
  }
  const parts = [base];
  for (const f of files) parts.push(await readFile(join(__dirname, f), "utf8"));
  return parts.join("\n");
}

// ---------------------------------------------------------------------------
// Per-turn message preparation (context header, slash-command handling,
// pull-based turn source) lives in daemon/runtime/turns.mjs, shared by
// every runtime.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Session management. Each open document (paneKey) runs its own
// independent agent loop with its own cwd; switching workspace or
// reloading config restarts only that pane's loop. The transport loop
// itself lives in daemon/runtime/ (Claude default; CodeBuddy via
// DRAFTSPECT_RUNTIME); this file owns the session lifecycle, restart and
// recovery semantics around it. Histories live in ~/.claude/projects/
// (Claude) or ~/.codebuddy/projects/ (CodeBuddy) per each SDK's normal
// persistence, with resume ids namespaced per provider in sessions.json.
// `sessions` / `workspaceByKey` are declared at the top of the module
// (above createBridge) so bridge handlers that fire during top-level
// awaits (e.g. preflightHttpMcpServers, which can block for 5s) don't
// hit a TDZ on those bindings.
// ---------------------------------------------------------------------------

async function startSessionForFolder(
  cwd,
  resumeSessionId = null,
  {
    key = null,
    host = null,
    replay = true,
    turnId = bridge.activeTurn(key),
    isCurrentStart = () => true,
  } = {},
) {
  // A queued start owns the turn captured when it was scheduled. Stop/reconnect
  // can replace both that turn and its queue while teardown is still pending.
  if (!isCurrentStart() || bridge.activeTurn(key) !== turnId) return null;
  const prior = sessionFor(key);
  prior?.cancel();
  if (!isCurrentStart() || bridge.activeTurn(key) !== turnId) return null;
  // No idle eager loop: all transports start on a real accepted user request.
  if (!turnId) {
    workspaceByKey.set(key, cwd);
    if (replay) await sendTranscriptReplayTo(key, host, cwd);
    return null;
  }
  const session = { key, cwd, sessionId: resumeSessionId, settled: false, host };
  const lifecycle = createLifecycle({
    emit: (event) => {
      if (event.event === "turn_complete") {
        const elapsed = session.turnBeganAt ? Date.now() - session.turnBeganAt : null;
        console.log?.(
          `[turn] ${event.turn_id ?? turnId} end subtype=${event.subtype}` +
            (event.error ? ` error=${JSON.stringify(String(event.error).slice(0, 200))}` : "") +
            ` elapsed=${elapsed}ms`,
        );
      }
      bridge.sendAssistantEvent(event, key);
    },
    cleanup: () => {
      session.settled = true;
      if (sessionFor(key) === session) {
        const activeTurn = bridge.activeTurn(key);
        if (isCurrentStart() && (!activeTurn || activeTurn === lifecycle.turnId)) {
          bridge.clearUserMessages(key);
          bridge.cancelTools(key);
        }
        sessions.delete(key);
      }
    },
    // Drives the timeout copy: only warn about unknown edit outcomes when a
    // mutating Office tool call is actually outstanding.
    writesPending: () => bridge.hasPendingWrites?.(key) ?? true,
  });
  session.lifecycle = lifecycle;
  session.cancel = () => lifecycle.cancel();
  session.close = () => lifecycle.close();
  session.turnBeganAt = Date.now();
  lifecycle.begin(turnId);
  sessions.set(key, session);
  workspaceByKey.set(key, cwd);
  const ownsStart = () =>
    !shuttingDown &&
    !lifecycle.closed &&
    sessionFor(key) === session &&
    bridge.activeTurn(key) === turnId &&
    isCurrentStart();
  const abandonStaleStart = () => {
    if (ownsStart()) return false;
    lifecycle.close(); // Close its timer/resources without touching a successor's queue.
    return true;
  };
  try {
    // The new turn already owns a cancellable lifecycle/timer. A close timeout
    // is therefore its controlled terminal error, never only a queue-runner log.
    if (runtime.id === "codebuddy" && runtimeClosures.has(key)) await runtimeClosures.wait(key);
    if (abandonStaleStart()) return session;
    await touchFolder(cwd);
    if (abandonStaleStart()) return session;
    if (
      handoffInProgress.has(key) ||
      (await getHandoff(host, cwd, runtime.id, key))?.state === "released"
    )
      throw new Error("Session is released for serial CLI reuse; return it to Word first");
    const append = await buildSystemPromptAppend(host);
    if (abandonStaleStart()) return session;
    const officeMcp = createOfficeBridgeMcp(
      {
        callTaskpaneTool: (name, args, paneKey) => {
          if (lifecycle.closed || !lifecycle.active || sessionFor(key) !== session)
            throw new Error("Expired session: Office tool call refused.");
          return bridge.callTaskpaneTool(name, args, paneKey);
        },
      },
      host,
      key,
      runtime.mcp,
    );
    bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeSessionId }, key);
    const handle = runtime.runSession({
      cwd,
      resumeId: resumeSessionId,
      model: modelArgFor(key),
      systemPromptAppend: append,
      mcpServers: { ...userMcpServers, office: officeMcp },
      disallowedTools: WORD_MCP_DISALLOWED,
      canUseTool: (name, input) =>
        lifecycle.closed || !lifecycle.active
          ? Promise.resolve({ behavior: "deny", message: "The owning request ended." })
          : customPermissionHandler(name, input),
      includePartialMessages: true,
      turnSource: createTurnSource({
        pull: () => bridge.nextUserMessage(key),
        session,
        onTurnStart: (msg) => {
          session.turnBeganAt = Date.now();
          if (!lifecycle.active) lifecycle.begin(msg.turn_id);
          session.renderMessage = createMessageRenderer((event) => {
            if (event.text || event.event === "tool_use_announce")
              session.turnProducedOutput = true;
            bridge.sendAssistantEvent(event, key);
          }, msg.turn_id);
        },
      }),
      isStale: () => lifecycle.closed || sessionFor(key) !== session,
      onMessage: (msg) => {
        if (!lifecycle.closed && sessionFor(key) === session) handleAgentMessage(msg, session);
      },
      onStderr: (data) => {
        for (const line of String(data).split(/\r?\n/))
          if (line.trim()) console.error("[sdk]", line);
      },
      onTurnEnd: ({ sawResult } = {}) => {
        if (sawResult === false)
          lifecycle.fail(
            "The stream ended without a final result. Send a new request to resume; verify any in-flight edit first.",
            "stream_ended",
          );
      },
      onEnd: ({ reason, error }) => {
        if (lifecycle.closed) return;
        if (reason === "error") {
          const failure = classifyAgentError(error);
          lifecycle.fail(
            failure.text,
            error?.subtype || "runtime_error",
            failure.isAuth ? { auth: true } : {},
          );
        } else if (lifecycle.active)
          lifecycle.fail("The stream ended without a final result.", "stream_ended");
        else lifecycle.close();
      },
    });
    if (runtime.id === "codebuddy") runtimeClosures.track(key, handle.closed, handle);
    lifecycle.attach(handle);
    session.done = handle.done;
  } catch (error) {
    // A stopped/replaced start must not emit for, clear, or configure its successor.
    if (ownsStart()) lifecycle.fail(error?.message ?? error, "startup_error");
    else lifecycle.close();
  }
  return session;
}

async function switchFolder(rawCwd, key, host = null) {
  const cwd = resolve(rawCwd);
  assertAcceptanceWorkspace(cwd);
  // Validate the path is a directory.
  const s = await stat(cwd);
  if (!s.isDirectory()) throw new Error(`Not a directory: ${cwd}`);
  if (handoffInProgress.has(key) || bridge.activeTurn(key))
    throw new Error("Stop or finish the current request before changing workspace.");
  const previousCwd = cwdForKey(key);
  // A failed unsaved ID remains recoverable in its original scope. Do not
  // replace that scope with a new workspace until explicit save retry succeeds.
  await sessionPersistence.awaitSaved(key, {
    ...(host ? { host } : {}),
    cwd: previousCwd,
    provider: runtime.id,
  });
  if (cwdForKey(key) !== previousCwd || handoffInProgress.has(key) || bridge.activeTurn(key))
    throw new Error("Session changed while saving; refresh before changing workspace.");
  // Switch ONLY the requesting pane to the target folder; resume that
  // (provider, host, cwd)'s prior conversation if one is on record. Every
  // other pane stays in its own workspace, untouched.
  workspaceByKey.set(key, cwd);
  sessionFor(key)?.cancel();
  const q = startQueue.get(key);
  if (q) q.latest = null;
  bridge.clearUserMessages(key);
  await sendTranscriptReplayTo(key, host, cwd);
  return cwd;
}

// Re-launch one pane's loop (same cwd, resuming via session_id) so that
// changes to CLAUDE.md, the drafting setup, or other config loaded at
// session-init take effect without losing conversation history. No-op if
// that pane has no live loop.
async function restartSession(key, host, { reason = "config_changed" } = {}) {
  const s = sessionFor(key);
  if (!s) return;
  if (s.lifecycle.active) {
    s.restartAfterTurn = true;
    bridge.sendAssistantEvent(
      { event: "info", message: "Configuration saved; it will apply after this request finishes." },
      key,
    );
    return;
  }
  const { cwd, sessionId } = s;
  console.log(`[daemon] Restarting session for ${cwd} (reason: ${reason})`);
  bridge.sendAssistantEvent({ event: "config_reloaded", reason }, key);
  // Funnel through the serialized per-key queue (not a direct
  // startSessionForFolder) so a config/model restart coalesces with any
  // concurrent post-Stop / first-message start instead of racing it.
  scheduleSessionStart(cwd, sessionId, key, host, reason, { replay: false });
}

function handleAgentMessage(msg, session) {
  // Every SDK message is proof of life: it resets the turn's idle timer.
  session.lifecycle?.activity?.();
  switch (msg.type) {
    case "system": {
      if (msg.subtype === "init") {
        const requestedModel = modelByKey.get(session.key);
        if (requestedModel && msg.model && requestedModel !== msg.model) {
          session.lifecycle.fail(
            `Model mismatch: selected ${requestedModel}, runtime reports ${msg.model}. No further tools will be dispatched.`,
            "model_mismatch",
          );
          return;
        }
        // The CodeBuddy transport re-emits init per stream segment, so one
        // turn can show many inits for the same session id; the count tells
        // reconnect churn apart from fresh sessions.
        session.initCount = (session.initCount ?? 0) + 1;
        console.log(
          `[agent] init session ${msg.session_id} (model: ${msg.model}) #${session.initCount}`,
        );
        try {
          const tn = Array.isArray(msg.tools) ? msg.tools : Object.keys(msg.tools ?? {});
          diag(
            `init tools (${tn.length}):`,
            tn.filter((t) => /office|excel|mcp__/.test(String(t))).join(", ") ||
              "(no office/excel/mcp tools in init list!)",
          );
        } catch (e) {
          diag("init tools introspection failed:", e?.message);
        }
        bridge.sendAssistantEvent(
          {
            event: "session_init",
            session_id: msg.session_id,
            model: msg.model,
          },
          session?.key,
        );
        // Record this session_id for THIS (provider, host, cwd) so the next
        // time this pane connects (or you switch back) it resumes here.
        if (session && msg.session_id && msg.session_id !== session.sessionId) {
          session.sessionId = msg.session_id;
          sessionPersistence.persist(session.key, {
            host: session.host,
            cwd: session.cwd,
            session_id: msg.session_id,
            provider: runtime.id,
          });
        }
      } else if (msg.subtype === "api_retry") {
        // Transport-level retry (network blip / upstream 5xx / rate limit).
        // Not a terminal state, but the pane must know the agent is stalled
        // instead of silently waiting. One event per turn, not per attempt —
        // retry storms (10× observed historically) would otherwise spam.
        const attempt = msg.attempt ?? msg.retry_count ?? null;
        console.log(`[agent] api_retry${attempt ? ` (attempt ${attempt})` : ""}`);
        if (session && !session.retryNotified) {
          session.retryNotified = true;
          bridge.sendAssistantEvent(
            {
              event: "retry",
              attempt,
              message: typeof msg.message === "string" ? msg.message : null,
            },
            session.key,
          );
        }
      } else {
        console.log(`[agent] system/${msg.subtype}`);
      }
      break;
    }
    case "stream_event":
    case "assistant":
    case "user":
      session.renderMessage?.(msg);
      break;
    case "result": {
      const outcome = classifyResult(msg);
      if (session.slashCommandPending && !session.turnProducedOutput) {
        bridge.sendAssistantEvent(
          {
            event: "info",
            message:
              "Command ended without chat output. Some CLI-only commands do not render here.",
          },
          session.key,
        );
      } else if (outcome.terminal === "success" && !session.turnProducedOutput) {
        bridge.sendAssistantEvent(
          {
            event: "info",
            message:
              "The runtime completed without an answer. No summary was produced; inspect the tool results or retry.",
          },
          session.key,
        );
      }
      session.lifecycle.finish({
        subtype: msg.subtype || outcome.terminal,
        ...(outcome.terminal === "error" ? { error: outcome.errorText } : {}),
        ...(outcome.kind === "auth" ? { auth: true } : {}),
        ...(msg.subtype === "interrupted" ? { interrupted: true } : {}),
      });
      if (outcome.terminal === "error" || session.restartAfterTurn) session.lifecycle.close();
      break;
    }
    default:
      break;
  }
}

// ---------------------------------------------------------------------------
// Kick off.
//
// No eager session: an agent session needs a host, and the host is only
// known once a taskpane connects and says hello (→ ensureSessionFor
// ActivePane). Starting both-tool sessions pre-hello is exactly what
// caused the per-host tool/transcript churn. The bridge buffers any
// user_message until a session's loop consumes it, so nothing is lost.
//
// Tell the Electron shell we're up now (servers listening) so the tray
// flips to "Ready" without waiting for a pane — independent of, and
// idempotent with, the daemon_ready that startSessionForFolder also
// emits on the first real session.
// ---------------------------------------------------------------------------
// Warm the runtime model catalog (CodeBuddy: one throwaway session to fetch
// it; Claude: instant static list) so the pane's first list_models call —
// fired on hello — answers without a cold-cache delay.
try {
  await Promise.all([http.ready, bridge.ready]);
} catch (error) {
  await Promise.allSettled([http.close(), bridge.close()]);
  if (error?.code === "EADDRINUSE") {
    // A bare node stack trace left the 2026-09-20 restart loop unexplained.
    console.error(
      `[daemon] ${error.address}:${error.port} is already in use — another Draftspect daemon (or a stale one) is probably running. Quit the other tray/instance first, then restart.`,
    );
  }
  throw error;
}
getRuntimeModels().catch(() => {});

console.log(`[daemon] Ready; waiting for a taskpane. Default workspace: ${matterFolder}`);
if (process.send) {
  try {
    daemonHealthSnapshot = () => ({ phase: "ready", ...bridge.diagnosticSummary() });
    process.send({ type: "daemon_ready" });
  } catch {
    /* no IPC channel (npm run dev) */
  }
}

// Keep process alive even when nothing is happening.
async function shutdown() {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log("\n[daemon] Shutting down");
  const owned = [...sessions.values()];
  for (const session of owned) session.cancel();
  // Close both listeners (destroying held connections) BEFORE exiting, so a
  // supervisor restarting us 1s later does not hit EADDRINUSE — the cause of
  // the 2026-09-20 crash loop that left no daemon running. Force-exit anyway
  // if a close hangs.
  const forced = setTimeout(() => {
    console.error("[daemon] RUNTIME_EXIT_UNCONFIRMED: shutdown deadline exceeded");
    process.exit(1);
  }, 3000);
  let exitCode = 0;
  try {
    const results = await Promise.allSettled([
      http.close(),
      bridge.close(),
      runtime.shutdown?.(2500),
      runtimeClosures.closeAll(2500),
      ...owned.map((session) =>
        runtime.id === "codebuddy" && runtimeClosures.has(session.key)
          ? runtimeClosures.wait(session.key, 2500)
          : Promise.resolve(session.done).catch(() => {}),
      ),
    ]);
    if (results.some((result) => result.status === "rejected")) {
      console.error("[daemon] RUNTIME_EXIT_UNCONFIRMED: one or more shutdown barriers failed");
      exitCode = 1;
    }
  } catch {
    exitCode = 1;
  }
  clearTimeout(forced);
  process.exit(exitCode);
}
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);
