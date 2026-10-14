import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";
import { createLocalServer, isLoopbackHost } from "./local-transport.mjs";
import { OFFICE_READ_ONLY_TOOLS } from "./office-read-only.mjs";

// 120s per tool call; a confirmation card re-arms it (see CONFIRM_WINDOW_MS).
// Some Office.js ops run long —
// office_clear_highlights / a wide office_replace_section over a large
// document, batched office_highlight, a whole-sheet excel_read_range.
// 30s was tripping on real-sized documents. (Per-tool timeouts would be
// finer-grained but aren't worth the complexity yet.)
const TOOL_TIMEOUT_MS = 120_000;
// When the pane shows a confirmation card, the deadline is re-armed so slow
// preview preparation cannot consume the user's authorization window.
export const CONFIRM_WINDOW_MS = 45_000;
const WRITE_RESERVE_MS = 60_000;

function normalizeHost(h) {
  return h === "word" || h === "excel" ? h : null;
}

// A pane is keyed by (host, document), not just host: a user can have two
// Word docs (or two workbooks) open at once, each with its own task pane.
// Keying by host alone made them fight over a single slot — the same
// connect/disconnect ping-pong the host split originally fixed, just one
// level down. The host stays embedded in the key so the tool family and
// per-host behavior are still derivable from it.
const PANE_SEP = "\u0000";
function makePaneKey(host, doc, paneId) {
  let d;
  if (doc && String(doc).trim()) {
    d = String(doc);
  } else if (paneId && String(paneId).trim()) {
    // Unsaved / cloud doc with no filesystem path: key by the taskpane's
    // stable per-instance id (sent in hello) so a reconnect REUSES this
    // pane's key instead of minting a fresh random one every ~1.5s while
    // the daemon is down — which leaked per-pane state unbounded.
    d = "anon:" + String(paneId);
  } else {
    d = "anon:" + randomUUID();
  }
  return host + PANE_SEP + d;
}
function hostOfKey(key) {
  return typeof key === "string" ? key.split(PANE_SEP, 1)[0] : null;
}

export function createBridge({
  port,
  tls = null,
  extraHandlers = {},
  token,
  allowedOrigins = [],
  allowHello,
  onHello,
  onUserMessage,
  onClose,
  onActivity,
  runtimeId = "codebuddy",
  toolTimeoutMs = TOOL_TIMEOUT_MS,
}) {
  if (!token) throw new Error("createBridge requires a token");
  const transport = createLocalServer({
    port,
    tls,
    handler: (_req, res) => res.writeHead(426).end("WebSocket upgrade required"),
  });
  const wss = new WebSocketServer({
    server: transport.server,
    // The supplied HTTP(S) listener binds only 127.0.0.1. TLS changes its
    // protocol, never the local-only boundary or the admission gates below.
    // First gate: only allow upgrades from known origins (the taskpane's
    // origin = our own HTTP server's). Browsers honor the Origin header on
    // WebSocket upgrades; rejecting unknown origins blocks malicious local
    // web pages from driving the agent.
    verifyClient: (info, cb) => {
      if (!isLoopbackHost(info.req, transport.address()?.port)) {
        cb(false, 403, "Forbidden host");
        return;
      }
      const origin = info.req.headers.origin || "";
      if (allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
        cb(true);
      } else {
        console.warn(`[bridge] Rejecting WS upgrade from origin: ${origin || "(none)"}`);
        cb(false, 403, "Forbidden origin");
      }
    },
  });
  // ws forwards errors from the supplied HTTP(S) server. The transport.ready
  // promise carries bind failure to the owner; retain an explicit diagnostic
  // instead of throwing an unhandled EventEmitter error before it can await.
  wss.on("error", (error) => console.error("[bridge] Server error:", error.message));

  // One pane per (host, document). A connection is "unbound" until its
  // hello names a host + active doc; a reconnect with the SAME key
  // replaces the prior pane (Office reloaded that document's panel), any
  // DIFFERENT key — different host OR a different document of the same
  // host — coexists.
  const panes = new Map(); // paneKey -> { ws, context }

  function emptyContext(host = null, activeDoc = null) {
    return { activeDoc, selection: null, trackChangesMode: "always", host };
  }

  function paneWs(key) {
    const p = panes.get(key);
    return p && p.ws.readyState === p.ws.OPEN ? p.ws : null;
  }

  // Tool calls awaiting their tool_result. Keyed by id; each remembers the
  // ws it was dispatched on so a stale pane's disconnect only rejects its
  // own in-flight calls.
  const pendingTools = new Map();
  const activeTurns = new Map();
  // Bounded, connection-local ledger. No prompt/document/tool payloads. Health
  // RPCs bypass assistant-event filtering, so a lost terminal can be reconciled.
  function health(key) {
    const pane = panes.get(key);
    return {
      connection_epoch: pane?.epoch ?? null,
      active: pane?.turn
        ? {
            ...pane.turn,
            pending_tools: [...pendingTools.entries()]
              .filter(([, p]) => p.key === key && p.ws === pane.ws)
              .map(([id, p]) => ({
                id,
                name: p.name,
                phase: p.phase,
                deadline: p.deadline,
                write_started: p.writeStarted,
              })),
          }
        : null,
      terminal: pane?.terminal ?? null,
    };
  }
  function progress(key, phase) {
    const turn = panes.get(key)?.turn;
    if (!turn || turn.turn_id !== activeTurns.get(key)) return;
    turn.phase = phase;
    turn.last_progress_at = Date.now();
  }

  // Per-pane queue of pending user messages. Waiters are {resolve, reject}
  // so we can reject pending awaits when that pane's session is aborted
  // (otherwise the suspended generator from the aborted session would
  // consume the next user message, starving the new session).
  const queues = new Map(); // paneKey -> { queue: [], waiters: [] }
  function queueFor(key) {
    let q = queues.get(key);
    if (!q) {
      q = { queue: [], waiters: [] };
      queues.set(key, q);
    }
    return q;
  }

  function pushUserMessage(text, key, turnId) {
    const ctx = panes.get(key)?.context ?? emptyContext(hostOfKey(key));
    const payload = {
      text,
      turn_id: turnId,
      context: { ...ctx, host: ctx.host ?? hostOfKey(key) },
    };
    const q = queueFor(key);
    if (q.waiters.length > 0) {
      q.waiters.shift().resolve(payload);
    } else {
      q.queue.push(payload);
    }
  }

  function nextUserMessage(key) {
    const q = queueFor(key);
    if (q.queue.length > 0) return Promise.resolve(q.queue.shift());
    return new Promise((resolve, reject) => q.waiters.push({ resolve, reject }));
  }

  // Reject all pending waiters and drop queued messages for ONE pane.
  // Called before that pane's session swaps so the new session starts
  // clean with no zombie consumer — and, crucially, without touching any
  // OTHER pane's queue (a message just enqueued for the pane we're
  // switching TO must survive).
  function clearUserMessages(key) {
    const q = queues.get(key);
    if (!q) return;
    while (q.waiters.length > 0) {
      q.waiters.shift().reject(new Error("Session aborted"));
    }
    q.queue.length = 0;
  }

  // Merge inbound doc/selection/track-changes fields into a pane's
  // context. `hello` resets doc+selection (a fresh taskpane connection —
  // absent fields mean null), while `context_update` patches (only fields
  // present on the message change). track-changes is "set only when
  // valid" in both modes. host is fixed at bind time, never re-derived.
  function mergeContext(ctx, msg, { reset = false } = {}) {
    if (reset) {
      ctx.activeDoc = msg.active_doc ?? null;
      ctx.selection = msg.selection ?? null;
    } else {
      if (msg.active_doc !== undefined) ctx.activeDoc = msg.active_doc;
      if (msg.selection !== undefined) ctx.selection = msg.selection;
    }
    if (typeof msg.track_changes_mode === "string") {
      ctx.trackChangesMode = msg.track_changes_mode;
    }
  }

  function send(obj, key) {
    const ws = paneWs(key);
    if (!ws) {
      console.warn(`[bridge] No pane for ${key ?? "?"}; dropping`, obj.type);
      return;
    }
    ws.send(JSON.stringify(obj));
  }

  async function callTaskpaneTool(name, args, key = null) {
    const ws = paneWs(key);
    if (!ws) {
      throw new Error(`Cannot call tool ${name}: no taskpane for ${key ?? "?"}`);
    }
    const id = randomUUID();
    // Capture the WS this call is dispatched on. On close we only reject
    // pending calls belonging to that specific WS — so a stale pane
    // disconnecting after a fresh one is active doesn't kill live work.
    const ownerWs = ws;
    const startedAt = Date.now();
    console.log(`[bridge] tool_call ${name} (${hostOfKey(key) ?? "?"})`);
    onActivity?.(key);
    const promise = new Promise((resolve, reject) => {
      const expire = () => {
        pendingTools.delete(id);
        send({ type: "cancel_office_tools" }, key);
        reject(
          new Error(
            `RESULT_UNKNOWN: Tool ${name} timed out. Inspect the document before retrying any write.`,
          ),
        );
      };
      const timer = setTimeout(expire, toolTimeoutMs);
      pendingTools.set(id, {
        expire,
        ws: ownerWs,
        key,
        resolve: (value) => {
          console.log(`[bridge] tool_result ${name} elapsed=${Date.now() - startedAt}ms ok`);
          onActivity?.(key);
          resolve(value);
        },
        reject: (error) => {
          console.log(`[bridge] tool_result ${name} elapsed=${Date.now() - startedAt}ms error`);
          onActivity?.(key);
          reject(error);
        },
        timer,
        name,
        phase: "office_pending",
        deadline: Date.now() + toolTimeoutMs,
        writeStarted: false,
      });
    });
    progress(key, "office_pending");
    ownerWs.send(
      JSON.stringify({
        type: "tool_call",
        id,
        name,
        args,
        turn_id: activeTurns.get(key),
        deadline: Date.now() + toolTimeoutMs,
      }),
    );
    return await promise;
  }

  function sendAssistantText(delta, key) {
    progress(key, "streaming");
    send({ type: "assistant_text", delta, turn_id: activeTurns.get(key) }, key);
  }

  function sendAssistantEvent(event, key) {
    if (event.turn_id && activeTurns.get(key) !== event.turn_id) return;
    if (event.event === "turn_complete") {
      const pane = panes.get(key);
      if (pane) {
        pane.terminal = {
          event: "turn_complete",
          turn_id: event.turn_id || activeTurns.get(key),
          subtype: event.subtype,
          error: event.error ? String(event.error).slice(0, 1000) : null,
          interrupted: !!event.interrupted,
          ...(event.auth ? { auth: true } : {}),
          finished_at: Date.now(),
        };
        pane.turn = null;
      }
      activeTurns.delete(key);
    } else if (event.event === "assistant_message") {
      progress(key, "streaming");
    } else if (event.event === "tool_use_announce") {
      progress(key, "tool_wait");
    } else if (event.event === "session_init") {
      progress(key, "waiting_model");
    }
    send({ type: "assistant_event", ...event }, key);
  }

  function cancelTools(key) {
    send({ type: "cancel_office_tools" }, key);
    for (const [id, p] of pendingTools) {
      if (p.key !== key) continue;
      clearTimeout(p.timer);
      pendingTools.delete(id);
      p.reject(
        new Error(
          "RESULT_UNKNOWN: request ended; verify any Office operation already started before retrying.",
        ),
      );
    }
  }

  wss.on("connection", (ws, req) => {
    console.log(
      `[bridge] WS connection from ${req.socket.remoteAddress} (origin: ${req.headers.origin || "(none)"})`,
    );

    // Second gate: first message must be a hello with the correct token.
    // Until that arrives we don't trust the connection — any other message
    // type is rejected and the socket is closed. The taskpane fetches the
    // token from /bridge-token over the same-origin HTTP server before
    // opening the WS. A connection is bound to exactly one (host,
    // document) pane key for its lifetime.
    let authed = false;
    let boundKey = null;
    let boundHost = null;

    // True iff this ws is still the live pane for the key it bound to.
    // A superseded (same-key reconnect) ws stays open just long enough
    // to be force-closed; drop any stale frames it sends meanwhile.
    function isLivePane() {
      return authed && panes.get(boundKey)?.ws === ws;
    }

    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch (e) {
        console.error("[bridge] Bad JSON from taskpane:", e.message);
        return;
      }

      if (!authed) {
        if (msg.type !== "hello") {
          console.warn(`[bridge] Pre-auth message ${msg.type}; closing`);
          ws.close(4001, "Authenticate first");
          return;
        }
        if (msg.token !== token) {
          console.warn("[bridge] hello with invalid token; closing");
          ws.close(4002, "Invalid token");
          return;
        }
        const host = normalizeHost(msg.host);
        if (!host) {
          console.warn(`[bridge] hello without a valid host (${msg.host}); closing`);
          ws.close(4005, "hello must name host word|excel");
          return;
        }
        if (allowHello) {
          let allowed = false;
          try {
            allowed = allowHello(host, msg.active_doc) === true;
          } catch {}
          if (!allowed) {
            ws.close(4007, "Document is outside this bridge's allowed scope");
            return;
          }
        }
        boundHost = host;
        boundKey = makePaneKey(host, msg.active_doc, msg.pane_id);
        // Replace only a prior pane with the SAME key (same document's
        // panel reloaded / reopened). A different document — even of the
        // same host — gets its own pane and coexists.
        const prior = panes.get(boundKey);
        if (prior && prior.ws !== ws && prior.ws.readyState === prior.ws.OPEN) {
          console.log(`[bridge] New pane authed for ${boundKey}; closing prior`);
          try {
            prior.ws.close(4003, "Replaced by new pane");
          } catch {}
          // End the old generation before installing its replacement. Its
          // later close event must not cancel a newly admitted request.
          cancelTools(boundKey);
          onClose?.(boundKey, boundHost);
          activeTurns.delete(boundKey);
          clearUserMessages(boundKey);
        }
        authed = true;
        panes.set(boundKey, {
          ws,
          context: emptyContext(host, msg.active_doc ?? null),
          epoch: randomUUID(),
          protocol: msg.protocol_version,
          turn: null,
          terminal: null,
          lastSeenAt: Date.now(),
        });
      } else if (!isLivePane()) {
        // Superseded by a newer same-key pane; its close is in flight.
        // Drop stale frames rather than letting them mutate context or
        // queue a user message for a pane that's going away.
        console.warn(`[bridge] Post-auth ${msg.type} from superseded pane ${boundKey}; ignoring`);
        return;
      }

      const pane = panes.get(boundKey);
      pane.lastSeenAt = Date.now();

      switch (msg.type) {
        case "hello": {
          mergeContext(pane.context, msg, { reset: true });
          ws.send(
            JSON.stringify({
              type: "welcome",
              session_id: randomUUID(),
              server_version: "0.2.0-cb",
              protocol_version: 2,
              capabilities: ["turn_health_v1", "confirm_window_v1"],
              connection_epoch: pane.epoch,
              runtime: runtimeId,
            }),
          );
          console.log(
            `[bridge] hello received; host: ${boundHost}; active doc: ${pane.context.activeDoc}`,
          );
          // Let the daemon replay this pane's transcript (and anything
          // else it wants on connect). Fire-and-forget; never block.
          if (onHello) {
            Promise.resolve(onHello(boundKey, boundHost, pane.context.activeDoc)).catch((err) =>
              console.warn("[bridge] onHello failed:", err?.message ?? err),
            );
          }
          break;
        }
        case "user_message": {
          const ack = (ok, error = null) => {
            if (msg.request_id)
              send(
                {
                  type: "user_message_result",
                  request_id: msg.request_id,
                  turn_id: msg.turn_id,
                  connection_epoch: pane.epoch,
                  ok,
                  error,
                },
                boundKey,
              );
          };
          if (
            pane.protocol === 2 &&
            (!msg.turn_id ||
              typeof msg.turn_id !== "string" ||
              msg.turn_id.length > 100 ||
              !msg.request_id ||
              typeof msg.request_id !== "string" ||
              msg.request_id.length > 100)
          ) {
            ack(false, "PROTOCOL_ERROR: a request and turn identity are required.");
            send(
              {
                type: "protocol_error",
                code: "MISSING_REQUEST_IDENTITY",
                message: "Request not accepted: refresh this panel before sending again.",
              },
              boundKey,
            );
            return;
          }
          if (typeof msg.text !== "string" || !msg.text.trim()) {
            ack(false, "EMPTY_REQUEST");
            return;
          }
          if (activeTurns.has(boundKey)) {
            ack(false, "TURN_BUSY: another request is active.");
            send(
              {
                type: "assistant_event",
                event: "info",
                message: "A request is already running. Stop it or wait before sending another.",
              },
              boundKey,
            );
            return;
          }
          const turnId =
            typeof msg.turn_id === "string" && msg.turn_id.length <= 100
              ? msg.turn_id
              : randomUUID();
          activeTurns.set(boundKey, turnId);
          pane.turn = {
            turn_id: turnId,
            request_id: msg.request_id ?? null,
            phase: "starting",
            started_at: Date.now(),
            last_progress_at: Date.now(),
          };
          console.log(`[bridge] user_message (${boundHost}), turn=${turnId}`);
          // Admission must complete BEFORE publishing to an idle SDK waiter.
          // In particular, persisted serial-CLI release is an async check.
          Promise.resolve()
            .then(() => onUserMessage?.(boundKey, boundHost, pane.context.activeDoc, msg.text))
            .then(() => {
              if (panes.get(boundKey)?.ws !== ws || activeTurns.get(boundKey) !== turnId) return;
              ack(true);
              progress(boundKey, "waiting_model");
              pushUserMessage(msg.text, boundKey, turnId);
            })
            .catch((err) => {
              if (panes.get(boundKey)?.ws !== ws || activeTurns.get(boundKey) !== turnId) return;
              ack(false, String(err?.message ?? err));
              clearUserMessages(boundKey);
              sendAssistantEvent(
                {
                  event: "turn_complete",
                  turn_id: turnId,
                  subtype: "startup_error",
                  error: String(err?.message ?? err),
                },
                boundKey,
              );
            });
          break;
        }
        case "context_update": {
          if (msg.active_doc !== undefined && msg.active_doc !== pane.context.activeDoc) {
            ws.close(4006, "Document changed; reconnect with its new identity");
            break;
          }
          mergeContext(pane.context, msg);
          break;
        }
        case "tool_result": {
          const pending = pendingTools.get(msg.id);
          if (!pending || pending.ws !== ws) {
            console.warn("[bridge] tool_result for unknown id:", msg.id);
            return;
          }
          clearTimeout(pending.timer);
          pendingTools.delete(msg.id);
          progress(boundKey, "waiting_model");
          if (msg.ok) {
            pending.resolve(msg.result);
          } else {
            pending.reject(new Error(msg.error ?? "Unknown tool error"));
          }
          break;
        }
        case "tool_phase": {
          const pending = pendingTools.get(msg.id);
          if (!pending || pending.ws !== ws || msg.turn_id !== activeTurns.get(boundKey)) break;
          if (!["office_read", "awaiting_confirmation", "office_write"].includes(msg.phase)) break;
          pending.phase = msg.phase;
          if (msg.phase === "office_write") pending.writeStarted = true;
          if (msg.phase === "awaiting_confirmation") {
            const deadline = Date.now() + CONFIRM_WINDOW_MS + WRITE_RESERVE_MS;
            if (deadline > pending.deadline) {
              clearTimeout(pending.timer);
              pending.deadline = deadline;
              pending.timer = setTimeout(pending.expire, deadline - Date.now());
            }
          }
          onActivity?.(boundKey);
          progress(boundKey, msg.phase);
          break;
        }
        case "get_turn_state": {
          ws.send(
            JSON.stringify({
              type: "get_turn_state_result",
              request_id: msg.request_id,
              ok: !msg.connection_epoch || msg.connection_epoch === pane.epoch,
              ...health(boundKey),
            }),
          );
          break;
        }
        case "ping": {
          ws.send(JSON.stringify({ type: "pong" }));
          break;
        }
        default: {
          // Custom handlers registered by the daemon (e.g. settings /
          // refs). They receive the message, a `reply(obj)` shortcut that
          // sends back over this same WS, and the bound pane key + host so
          // per-pane actions (set_cwd, stop_agent, …) act on the right
          // document.
          const handler = extraHandlers[msg.type];
          if (handler) {
            const reply = (obj) => ws.send(JSON.stringify(obj));
            Promise.resolve()
              .then(() => handler(msg, reply, boundKey, boundHost))
              .catch((err) => {
                console.error(`[bridge] handler for ${msg.type} threw:`, err);
                reply({
                  type: msg.type + "_result",
                  ok: false,
                  error: err.message ?? String(err),
                  request_id: msg.request_id,
                });
              });
          } else {
            console.warn("[bridge] Unknown message type:", msg.type);
          }
        }
      }
    });

    ws.on("close", () => {
      console.log(`[bridge] Taskpane disconnected${boundKey ? ` (${boundHost})` : ""}`);
      // Only clear the pane mapping if THIS ws still owns it (a superseded
      // ws closing must not evict the fresh same-key pane).
      if (boundKey && panes.get(boundKey)?.ws === ws) {
        panes.delete(boundKey);
        // End this connection's generation. The next request resumes by
        // document/provider ID; no suspended consumer survives disconnect.
        if (onClose) {
          try {
            onClose(boundKey, boundHost);
          } catch (err) {
            console.warn("[bridge] onClose failed:", err?.message ?? err);
          }
        }
        activeTurns.delete(boundKey);
        clearUserMessages(boundKey);
        queues.delete(boundKey);
      }
      // Reject only the pending tool calls dispatched on THIS ws. A
      // different (newer) WS may have its own in-flight calls; leave them.
      for (const [id, p] of pendingTools) {
        if (p.ws === ws) {
          clearTimeout(p.timer);
          p.reject(
            new Error(
              "RESULT_UNKNOWN: taskpane disconnected; inspect any write already started before retrying.",
            ),
          );
          pendingTools.delete(id);
        }
      }
    });

    ws.on("error", (err) => {
      console.error("[bridge] WS error:", err.message);
    });
  });

  transport.ready.then(
    (address) =>
      console.log(
        `[bridge] WebSocket server listening on ${tls ? "wss" : "ws"}://127.0.0.1:${address.port} (origin allowlist: ${allowedOrigins.join(", ") || "<empty>"})`,
      ),
    () => {},
  );

  return {
    health,
    diagnosticSummary: () => ({
      connected_panes: panes.size,
      connected_panes_by_host: {
        word: [...panes.values()].filter((pane) => pane.context.host === "word").length,
        excel: [...panes.values()].filter((pane) => pane.context.host === "excel").length,
      },
      active_turns: activeTurns.size,
      unresponsive_active_panes: [...panes.values()].filter(
        (p) => p.protocol === 2 && p.turn && Date.now() - p.lastSeenAt > 90000,
      ).length,
    }),
    activeTurn: (key) => activeTurns.get(key) ?? null,
    // A turn interrupted while one of these is outstanding may have edits in
    // flight; read-only tools never modify the document. Drives the wording
    // of the turn-timeout terminal message. False is NOT evidence of no prior
    // edit or of an earlier RESULT_UNKNOWN being resolved.
    hasPendingWrites: (key) => {
      for (const p of pendingTools.values()) {
        if (p.key === key && !OFFICE_READ_ONLY_TOOLS.has(p.name)) return true;
      }
      return false;
    },
    cancelTools,
    nextUserMessage,
    clearUserMessages,
    callTaskpaneTool,
    sendAssistantText,
    sendAssistantEvent,
    sendToTaskpane: send,
    getContext: (key) => ({ ...(panes.get(key)?.context ?? emptyContext(hostOfKey(key))) }),
    isTaskpaneConnected: (key) =>
      key ? !!paneWs(key) : [...panes.keys()].some((k) => !!paneWs(k)),
    // Test/observability surface: the bound address (null until listening)
    // and a clean shutdown. Used by the loopback-bind unit test.
    ready: transport.ready,
    address: transport.address,
    // ws does not close an externally supplied HTTP(S) server. Close both,
    // including the listener on bind failure, so tests/restarts do not leak it.
    close: async () => {
      for (const ws of wss.clients) ws.terminate();
      await Promise.all([new Promise((resolve) => wss.close(resolve)), transport.close()]);
    },
  };
}
