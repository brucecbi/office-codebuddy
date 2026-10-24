/* global Office, Word, Excel */

import {
  toolGetSelection,
  toolReadParagraphs,
  toolInsertParagraphs,
  toolReplaceParagraphs,
  toolApplyTextReplacement,
  toolReplaceSection,
  toolDeleteParagraphs,
  toolHighlight,
  toolClearHighlights,
  toolAddComment,
  toolClearComments,
  toolApplyStyle,
  toolSetFont,
  toolSetParagraphFormatting,
  toolInsertTable,
  toolSetTableCell,
  toolGetDocumentText,
  toolGetOutline,
  toolSetList,
  toolInsertImage,
  toolInsertHyperlink,
  toolInsertBookmark,
  toolFind,
  toolListComments,
  toolReplyToComment,
  toolResolveComment,
  toolHeaderFooter,
  prepareSafeWordWrite,
  prepareTextReplacement,
  resolveWordHeadingArgs,
  captureWordFormattingPrecondition,
} from "./tools-word.js";
import {
  toolExcelGetSelectedRange,
  toolExcelListSheets,
  toolExcelReadRange,
  toolExcelWriteRange,
  toolExcelFindValue,
  toolExcelInsertRows,
  toolExcelDeleteRows,
  toolExcelSelectRange,
  toolExcelWriteFormula,
  toolExcelSetFormat,
  toolExcelInsertColumns,
  toolExcelDeleteColumns,
  toolExcelAddSheet,
  toolExcelDeleteSheet,
  toolExcelRenameSheet,
  toolExcelClearRange,
  toolExcelSortRange,
  toolExcelAutoFilter,
  toolExcelCreateTable,
  toolExcelAddTableRows,
  toolExcelCreateChart,
  toolExcelSetColumnWidth,
  toolExcelSetRowHeight,
  captureExcelPrecondition,
} from "./tools-excel.js";
import { isInOrUnder, docDirFromActiveUrl } from "./paths.js";
import { planReplacement } from "./word-edit-plan.js";
import { createOperationQueue } from "./operation-queue.js";
import {
  toolDocumentMap,
  toolReadStory,
  toolReadFormatting,
  toolReadRevisions,
  toolReadComments,
  toolAuditRedline,
  prepareReviewWrite,
} from "./word-review.js";

// Daemon endpoints. The HTTP server that loaded this taskpane is on
// HTTP_PORT; the WebSocket bridge is on WS_PORT (one less by daemon convention).
const WS_URL = `${location.protocol === "https:" ? "wss" : "ws"}://127.0.0.1:47833`;
const TOKEN_URL = "/bridge-token";

// The bridge token — fetched at boot from the same-origin HTTP server. The
// bridge rejects any WS that doesn't present this token in its first hello.
let bridgeToken = null;

async function fetchBridgeToken() {
  try {
    const res = await fetch(TOKEN_URL, { credentials: "same-origin" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    bridgeToken = (await res.text()).trim();
  } catch (e) {
    console.warn("Failed to fetch bridge token:", e);
    bridgeToken = null;
  }
}

let ws = null;
let wsReady = false;
let lastSelection = null;
let activeDocUrl = null;

// Host: "word" or "excel". Set from <body data-host="..."> (the Word and
// Excel taskpanes each load their own index.html which sets this), with a
// fallback to Office.context.host when Office.onReady fires.
let HOST = document.body.dataset.host === "excel" ? "excel" : "word";
function hostLabel() {
  return HOST === "excel" ? "Excel" : "Word";
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
const $messages = document.getElementById("messages");
// Two status indicators, semantically separate:
//   $connectionStatus (topbar) — WS bridge / daemon reachability + auth.
//     Stays put while the user is reading the chat history.
//   $agentStatus (above composer) — agent activity: Ready / Working —
//     <tool>… / Stopped. Lives right where the user's eye is when they
//     send a message.
const $connectionStatus = document.getElementById("connection-status");
const $agentStatus = document.getElementById("agent-status");
const $stopAgent = document.getElementById("stop-agent");
const $activeDoc = document.getElementById("active-doc");
const $input = document.getElementById("input");
const $send = document.getElementById("send");
const $composer = document.getElementById("composer");
const $chip = document.getElementById("selection-chip");
const $chipText = document.getElementById("selection-chip-text");
const $chipDetach = document.getElementById("selection-chip-detach");

let assistantTurnElem = null;
let attachSelection = true;

// True while a user turn is mid-flight (we've sent user_message and are
// waiting for turn_complete / error / auth_error). The composer is disabled
// during this window so a second submit can't break turn ordering — without
// the gate, appendUserMessage would reset assistantTurnElem and subsequent
// streaming deltas would land in a new (wrong) bubble.
let turnInFlight = false;
let currentTurnId = null;
let modelConfirmed = false;
let modelChoiceRequired = false;
let replayReady = false;
let activeRuntime = null;
let currentSessionInfo = null;
let sessionReleased = false;
let handoffBusy = false;
let connectionEpoch = null;
let healthSupported = false;
let confirmWindowSupported = false;
let healthTimer = null;
let healthPolling = false;
let healthFailures = 0;
let healthBlocked = false;
let persistenceWarning = null;
let lastHealthNotice = null;
let connectionInterruptedTurn = false;

function showHealthNotice(message, code = "", retry = false) {
  const box = document.getElementById("health-notice");
  const body = document.getElementById("health-notice-text");
  if (box && body) {
    box.hidden = !message;
    body.textContent = message || "";
    const detail = document.getElementById("health-notice-detail");
    if (detail)
      detail.textContent = code ? `${code} · ${currentTurnId || "no active request"}` : "";
    const button = document.getElementById("retry-models");
    if (button) button.hidden = !retry;
  } else if (message && message !== lastHealthNotice) appendError(message);
  lastHealthNotice = message;
}
function refreshReadyStatus() {
  if (turnInFlight) return;
  if (wsReady && modelConfirmed && replayReady && !healthBlocked) {
    if (connectionInterruptedTurn) {
      showHealthNotice(
        `连接已恢复。原请求未被自动重发；请重新读取并核对 ${hostLabel()}，再手动提出新请求。`,
        "CONNECTION_RECOVERED",
      );
      connectionInterruptedTurn = false;
    } else if (
      document
        .getElementById("health-notice-detail")
        ?.textContent?.startsWith("CONNECTION_UNAVAILABLE")
    )
      showHealthNotice(null);
  }
  if (operationStorage?.getItem(inspectionKey(activeDocUrl))) {
    setAgentStatus("err", `修改结果待核实 — 请检查 ${hostLabel()}`);
  } else if (persistenceWarning) setAgentStatus("err", "会话保存失败 — 恢复能力受影响");
  else if (healthBlocked) setAgentStatus("err", "状态异常 — 请检查连接");
  else if (!wsReady)
    setAgentStatus(
      connState === "err" ? "err" : "idle",
      connState === "err" ? "连接已断开 — 正在重试…" : "正在连接…",
    );
  else if (modelChoiceRequired) setAgentStatus("err", "请选择可用模型 — 尚未提交模型请求");
  else if (!modelConfirmed) setAgentStatus("idle", "等待模型确认…");
  else if (!replayReady) setAgentStatus("idle", "正在加载历史…");
  else if (sessionReleased) setAgentStatus("idle", "会话已交接至 CLI");
  else setAgentStatus("idle", "可发送 · 模型连通性与鉴权待实际请求验证");
}
function scheduleHealthCheck() {
  clearTimeout(healthTimer);
  if (!healthSupported || !wsReady) return;
  healthTimer = setTimeout(
    async () => {
      await checkTurnHealth();
      scheduleHealthCheck();
    },
    turnInFlight ? 5000 : 25000,
  );
}
async function checkTurnHealth() {
  if (!healthSupported || !wsReady || healthPolling) return;
  const poll = {};
  healthPolling = poll;
  const owner = ws,
    epoch = connectionEpoch;
  const queriedTurn = currentTurnId,
    queriedInFlight = turnInFlight;
  const stillCurrent = () =>
    owner === ws &&
    epoch === connectionEpoch &&
    queriedTurn === currentTurnId &&
    queriedInFlight === turnInFlight;
  try {
    const result = await sendRequest("get_turn_state", { connection_epoch: epoch });
    if (!stillCurrent()) return;
    if (!result.ok || result.connection_epoch !== epoch)
      throw new Error("CONNECTION_EPOCH_MISMATCH");
    healthFailures = 0;
    healthBlocked = false;
    if (!turnInFlight) {
      if (result.active) {
        healthBlocked = true;
        showHealthNotice(
          "服务端存在未同步的请求。请停止该请求或重新连接，不会自动接管。",
          "TURN_STATE_CONFLICT",
        );
      }
      setComposerDisabled(false);
      refreshReadyStatus();
      return;
    }
    if (result.terminal?.turn_id === currentTurnId && !result.active) {
      await handleServerMessage({ type: "assistant_event", ...result.terminal });
      showHealthNotice(
        result.terminal.error
          ? `请求已结束，状态已同步：${result.terminal.error}`
          : "请求已结束，界面状态已同步。",
        "TERMINAL_RECONCILED",
      );
      return;
    }
    if (!result.active) {
      cancelOfficeOperations();
      endTurn();
      showHealthNotice(
        "服务端没有这次请求的状态记录。请检查文档后重新提出请求；原请求未被自动重发。",
        "REQUEST_STATE_MISSING",
      );
      setAgentStatus("err", "请求状态失配 — 查看提示");
      return;
    }
    if (result.active.turn_id !== currentTurnId) {
      healthBlocked = true;
      cancelOfficeOperations();
      showHealthNotice(
        "请求标识冲突。请停止或重新连接；不会执行其他请求的文档操作。",
        "TURN_STATE_CONFLICT",
      );
      setAgentStatus("working", "请求状态冲突 — 请停止");
      return;
    }
    const labels = {
      starting: "正在启动",
      waiting_model: "等待模型回复",
      streaming: "正在接收回复",
      tool_wait: "等待工具返回",
      office_pending: `等待 ${hostLabel()} 操作`,
      office_read: `正在读取 ${hostLabel()}`,
      awaiting_confirmation: "等待你确认修改",
      office_write: `正在写入 ${hostLabel()}`,
    };
    const elapsed = Math.max(0, Math.floor((Date.now() - result.active.started_at) / 1000));
    setAgentStatus("working", `${labels[result.active.phase] || "正在处理"} · ${elapsed} 秒`);
    if (
      Date.now() - result.active.last_progress_at > 45000 &&
      result.active.phase !== "awaiting_confirmation"
    )
      showHealthNotice(
        "本阶段耗时较长，本地服务仍有响应。可以继续等待或停止；文档操作不会自动重试。",
        "SLOW_PROGRESS",
      );
    else if (
      ["SLOW_PROGRESS", "SERVICE_UNRESPONSIVE"].some((c) =>
        document.getElementById("health-notice-detail")?.textContent?.startsWith(c),
      )
    )
      showHealthNotice(null);
  } catch (error) {
    if (!stillCurrent()) return;
    if (++healthFailures >= 2) {
      healthBlocked = true;
      cancelOfficeOperations();
      setComposerDisabled(turnInFlight);
      setAgentStatus("err", "本地服务未响应 — 可停止或重新连接");
      // An unresponsive connection cannot prove the request ended. Keep its
      // admission lock and Stop control, with an error rather than busy UI.
      $stopAgent.hidden = !turnInFlight;
      showHealthNotice(
        `连续健康检查未收到响应。连接或本地服务可能异常；写入结果不明时请先检查 ${hostLabel()}。`,
        "SERVICE_UNRESPONSIVE",
      );
    }
  } finally {
    if (healthPolling === poll) healthPolling = false;
  }
}
document.getElementById("retry-models")?.addEventListener("click", () => requestModelList());
document.getElementById("health-check")?.addEventListener("click", () => checkTurnHealth());
document.getElementById("health-reconnect")?.addEventListener("click", () => {
  cancelOfficeOperations();
  ws?.close();
});
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "visible") checkTurnHealth();
});

// Use the daemon's welcome identity, never model names or assistant self-reports.
// Before handshake (or after disconnect), keep the copy provider-neutral.
function renderRuntimeCopy() {
  const label =
    activeRuntime === "codebuddy" ? "CodeBuddy" : activeRuntime === "claude" ? "Claude" : null;
  $input.placeholder = label ? `Ask ${label}…` : "Ask your agent…";
  const model = document.getElementById("composer-model");
  if (model)
    model.title = label
      ? `Model — provided by ${label}`
      : "Model — provided by the connected runtime";
}
renderRuntimeCopy();

function setComposerDisabled(disabled) {
  const blocked =
    disabled ||
    !wsReady ||
    !modelConfirmed ||
    !replayReady ||
    sessionReleased ||
    handoffBusy ||
    healthBlocked;
  $send.disabled = blocked;
  $input.disabled = blocked;
  const select = document.getElementById("composer-model");
  if (select) select.disabled = disabled || !wsReady;
  renderSessionControls();
}

function renderSessionControls() {
  const info = currentSessionInfo;
  const label = document.getElementById("session-identity");
  if (label)
    label.textContent = info?.session_id
      ? `${info.provider} · ${info.state}\n${info.session_id}`
      : "No session yet — send a message first.";
  const commands = document.getElementById("session-commands");
  if (commands)
    commands.textContent = info?.commands
      ? `Continue the same session:\n${info.commands.resume}\n\nReuse in a new session:\n${info.commands.reuse_as_new_session}\n\nArchive: ${info.archive_path || "not yet created"}`
      : "";
  const release = document.getElementById("session-release");
  if (release) release.disabled = !wsReady || !info?.session_id || turnInFlight || handoffBusy;
  const back = document.getElementById("session-return");
  if (back) {
    back.hidden = !sessionReleased;
    back.disabled = !wsReady || handoffBusy;
    back.textContent = info?.return_confirmation
      ? "再次确认 CLI 已退出，接受记录变化并返回"
      : "CLI exited — return to Office";
  }
  const retry = document.getElementById("session-retry-save");
  if (retry) {
    retry.hidden = !persistenceWarning;
    retry.disabled = !wsReady || handoffBusy;
  }
  const recovery = document.getElementById("session-process-recovery");
  if (recovery) {
    recovery.innerHTML = "";
    for (const process of info?.runtime_processes || []) {
      if (!process.close_requested) continue;
      const row = document.createElement("div");
      row.textContent = `本插件子进程 ${process.pid || "不可观察"}：${process.phase}。`;
      if (process.can_terminate && process.pid) {
        const button = document.createElement("button");
        button.textContent = `结束本插件进程 ${process.pid} 并等待退出`;
        button.disabled = turnInFlight || handoffBusy;
        button.addEventListener("click", () => recoverRuntime(process.pid, false));
        row.appendChild(button);
        const label = document.createElement("label"),
          force = document.createElement("input");
        force.type = "checkbox";
        label.append(
          force,
          document.createTextNode
            ? document.createTextNode("普通结束无效时，允许强制结束该进程（可能丢失未保存历史）")
            : Object.assign(document.createElement("span"), {
                textContent: "允许强制结束该进程（可能丢失未保存历史）",
              }),
        );
        const kill = document.createElement("button");
        kill.textContent = "强制结束已显示的插件进程";
        kill.disabled = true;
        force.addEventListener("change", () => {
          kill.disabled = !force.checked || turnInFlight || handoffBusy;
        });
        kill.addEventListener("click", () => {
          if (force.checked) recoverRuntime(process.pid, true);
        });
        row.append(label, kill);
      }
      recovery.appendChild(row);
    }
  }
}
async function recoverRuntime(pid, force) {
  if (!wsReady || turnInFlight || handoffBusy || inFlightWrites.size) return;
  handoffBusy = true;
  setComposerDisabled(true);
  try {
    const result = await sendRequest("recover_runtime", {
      confirm: true,
      expected_pid: pid,
      force,
      confirm_force: force,
    });
    if (!result.ok) throw new Error(result.error);
    appendNotice("已确认本插件子进程退出。可重新归档或返回 Word；没有跳过退出屏障。");
  } catch (error) {
    appendError(error.message);
  } finally {
    handoffBusy = false;
    await refreshSessionInfo().catch((error) => appendError(error.message));
    setComposerDisabled(turnInFlight);
  }
}
document.getElementById("session-retry-save")?.addEventListener("click", async () => {
  if (!wsReady || handoffBusy) return;
  try {
    const result = await sendRequest("retry_session_save");
    if (!result.ok) throw new Error(result.error);
    receiveSessionInfo(result.info);
  } catch (error) {
    showHealthNotice(`会话仍未保存：${error.message}`, "SESSION_PERSISTENCE_FAILED");
  }
});
function receiveSessionInfo(info) {
  currentSessionInfo = info || null;
  if (info?.persistence_status === "failed") {
    persistenceWarning = info.persistence_error || "SESSION_PERSISTENCE_FAILED";
    showHealthNotice(
      "会话关联未保存。可在 Setup 中检查并重试保存；不要假定历史已可恢复。",
      persistenceWarning,
    );
  } else if (info?.persistence_status === "saved") persistenceWarning = null;
  sessionReleased = info?.state === "released";
  setComposerDisabled(turnInFlight);
}
async function refreshSessionInfo() {
  if (!wsReady) return;
  const reply = await sendRequest("get_session_info");
  if (!reply.ok) throw new Error(reply.error);
  receiveSessionInfo(reply.info);
}
document.getElementById("session-release")?.addEventListener("click", async () => {
  if (turnInFlight || inFlightWrites.size || handoffBusy) {
    appendNotice("Finish the current turn and inspect any pending Word edits before archiving.");
    return;
  }
  handoffBusy = true;
  setComposerDisabled(true);
  try {
    cancelOfficeOperations();
    const reply = await sendRequest("release_session");
    if (!reply.ok) throw new Error(reply.error);
    receiveSessionInfo(reply.info);
    appendNotice(
      "Session closed and archived locally. CLI commands are in Setup. Exit the CLI before returning to Word; the archive is conversation history, not a document backup.",
    );
  } catch (error) {
    appendError("Archive/handoff: " + error.message);
  } finally {
    handoffBusy = false;
    try {
      await refreshSessionInfo();
    } catch (error) {
      appendError(error.message);
    }
    setComposerDisabled(turnInFlight);
  }
});
document.getElementById("session-return")?.addEventListener("click", async () => {
  if (!currentSessionInfo?.session_id || handoffBusy) return;
  handoffBusy = true;
  setComposerDisabled(true);
  try {
    const reply = await sendRequest("return_session", {
      session_id: currentSessionInfo.session_id,
      cli_closed: true,
      ...(currentSessionInfo.return_confirmation
        ? { confirm_transcript_change: currentSessionInfo.return_confirmation }
        : {}),
    });
    if (!reply.ok) {
      if (reply.info) receiveSessionInfo(reply.info);
      if (reply.code === "TRANSCRIPT_CHANGED")
        throw new Error(
          "会话记录已被外部修改，已拒绝返回。请确认 CLI 已退出并核对变化，再次点击返回按钮；不会重放任何操作。",
        );
      throw new Error(reply.error);
    }
    receiveSessionInfo(reply.info);
    appendNotice(
      "Returned to Word. History was reloaded; reread the live document before analysis or editing.",
    );
  } catch (error) {
    appendError(error.message);
  } finally {
    handoffBusy = false;
    setComposerDisabled(turnInFlight);
  }
});
document.getElementById("session-copy-id")?.addEventListener("click", async () => {
  if (!currentSessionInfo?.session_id) return;
  try {
    await navigator.clipboard.writeText(currentSessionInfo.session_id);
  } catch {
    appendNotice("Copy the complete session ID shown in Setup.");
  }
});

function beginTurn() {
  turnInFlight = true;
  setComposerDisabled(true);
  scheduleHealthCheck();
}

function endTurn() {
  turnInFlight = false;
  currentTurnId = null;
  if (currentSessionInfo && !sessionReleased) currentSessionInfo.state = "word";
  setComposerDisabled(false);
  scheduleHealthCheck();
}

// The connection status doubles as the live-model indicator. Horizontal
// space in the topbar is tight (the workspace chip takes the rest), so when
// connected we show ONLY the short model name and let the colored dot carry
// connection state; the full SDK model id goes in the tooltip. The model
// shown is the trustworthy `session_init` value from the daemon — never the
// agent's own (unreliable) self-report.
let connState = "idle";
let connLabel = "Connecting…";
let liveModel = null; // SDK-reported model id of the running loop
let pendingModelShort = null; // target short name while a switch restarts

// Pre-connect fallback for the experimental Claude adapter's standard aliases.
// The AUTHORITATIVE list now comes from the daemon (list_models — the
// active runtime's catalog: static for Claude, discovered for CodeBuddy);
// these values keep the dropdown sane if the daemon can't answer yet, and
// mirror runtime-claude.mjs STATIC_MODELS.
const MODELS = ["haiku", "sonnet", "opus"];
const DEFAULT_MODEL = "sonnet";

function shortModelName(idOrAlias) {
  const s = String(idOrAlias || "");
  if (/glm-5\.2/i.test(s)) return "GLM 5.2";
  if (/deepseek/i.test(s) && /flash/i.test(s)) return "DeepSeek v4 Flash";
  if (/deepseek/i.test(s) && /pro/i.test(s)) return "DeepSeek v4 Pro";
  if (/deepseek/i.test(s)) return "DeepSeek";
  if (/hy3/i.test(s)) return "HY3";
  if (/sonnet-5/i.test(s)) return "Sonnet 5";
  if (/sonnet/i.test(s)) return "Sonnet";
  if (/opus/i.test(s)) return "Opus";
  if (/haiku/i.test(s)) return "Haiku";
  if (/kimi/i.test(s)) return "Kimi";
  if (/gemini/i.test(s)) return "Gemini";
  if (/minimax/i.test(s)) return "MiniMax";
  if (/qwen/i.test(s)) return "Qwen";
  if (/gpt/i.test(s)) return s.replace(/-ioa$/i, "").toUpperCase();
  return s.replace(/^claude-/, "").split(/[-[]/)[0] || s;
}

function renderConnection() {
  let text = connLabel;
  let title = "Connection to the Draftspect daemon";
  if (connState === "ok") {
    if (pendingModelShort) {
      text = `↻ ${pendingModelShort}`;
      title = `Switching model → ${pendingModelShort}…`;
    } else if (liveModel) {
      text = shortModelName(liveModel);
      title = `Connected · ${activeRuntime} · runtime reported ${liveModel}`;
    } else {
      // Connected but no turn has run yet — show what the next message
      // will use (the sticky dropdown choice) rather than a bare "Connected".
      text = shortModelName(settings?.model || DEFAULT_MODEL);
      title = `Connected · ${activeRuntime} · requested ${text} on next message (runtime init pending)`;
    }
  }
  $connectionStatus.className = `status ${connState}`;
  $connectionStatus.textContent = text;
  $connectionStatus.title = title;
}

function setConnectionStatus(state, label) {
  connState = state;
  connLabel = label;
  renderConnection();
}

function setAgentStatus(state, label) {
  $agentStatus.className = `agent-status ${state}`;
  $agentStatus.textContent = label;
  // Stop button is meaningful only while the agent is mid-turn.
  $stopAgent.hidden = state !== "working";
}

// Stop button — abort the current agent turn. The daemon picks up the
// abort, emits turn_complete with interrupted=true (flipping this
// indicator to "Stopped"), and auto-restarts a fresh resuming loop.
$stopAgent?.addEventListener("click", async () => {
  if (!wsReady) return;
  setAgentStatus("working", "Stopping…");
  cancelOfficeOperations();
  try {
    const reply = await sendRequest("stop_agent");
    if (!reply.ok) throw new Error(reply.error);
  } catch (error) {
    appendError(
      `Stop was not confirmed: ${error.message}. Pending edits may have an unknown outcome.`,
    );
    // Keep the turn until the daemon reports its state: ending it locally would
    // unlock a composer the daemon rejects and drop the turn's late result.
    if (healthSupported && wsReady) {
      setAgentStatus("working", "Stop not confirmed — checking request state");
      checkTurnHealth();
      return;
    }
    setAgentStatus("idle", "Stop not confirmed");
    endTurn();
  }
});

// Auth-failure banner. Shown across the top of the panel when the daemon
// emits event: "auth_error". Persists until the user dismisses it; recovery
// uses the active runtime's sign-in flow, which cannot run in the taskpane.
function showAuthErrorBanner(rawError) {
  let banner = document.getElementById("auth-error-banner");
  if (!banner) {
    banner = document.createElement("div");
    banner.id = "auth-error-banner";
    banner.className = "auth-error-banner";
    banner.innerHTML = `
      <div class="auth-error-head">
        <strong>Sign-in required</strong>
        <button type="button" class="auth-error-dismiss" title="Dismiss">×</button>
      </div>
      <div class="auth-error-body"></div>
      <details class="auth-error-raw">
        <summary>Raw error</summary>
        <pre></pre>
      </details>
    `;
    document.body.insertBefore(banner, document.body.firstChild);
    banner.querySelector(".auth-error-dismiss").addEventListener("click", () => {
      banner.hidden = true;
    });
  }
  // Refresh an existing banner too, in case a reconnect changed runtimes.
  const guidance =
    activeRuntime === "codebuddy"
      ? `CodeBuddy could not authenticate. Run <code>codebuddy</code> in a terminal and follow its sign-in prompts.`
      : activeRuntime === "claude"
        ? `Claude Code could not authenticate. Run <code>claude</code> in a terminal and follow its sign-in prompts, or configure <code>ANTHROPIC_API_KEY</code> for your runtime.`
        : `The agent could not authenticate. Check the active runtime's sign-in configuration.`;
  banner.querySelector(".auth-error-body").innerHTML =
    guidance + " After signing in, quit Draftspect (tray icon) and reopen it.";
  banner.querySelector(".auth-error-raw pre").textContent = String(rawError || "(no detail)");
  banner.hidden = false;
}

// Map a tool name to a short, user-friendly status label shown in the topbar
// while the agent is mid-turn. Falls back to a generic "Working…" for tools
// the user hasn't seen named before — most filesystem/Bash/MCP tools.
const TOOL_STATUS_LABELS = {
  // Word
  office_get_selection: "Reading your selection…",
  office_read_paragraphs: "Reading the document…",
  office_insert_paragraphs: "Inserting paragraphs…",
  office_replace_text: "Editing text…",
  office_replace_paragraphs: "Replacing paragraphs…",
  office_replace_section: "Rewriting section…",
  office_delete_paragraphs: "Deleting paragraphs…",
  office_highlight: "Highlighting…",
  office_clear_highlights: "Clearing highlights…",
  office_add_comment: "Adding comment…",
  office_clear_comments: "Clearing comments…",
  office_apply_style: "Applying style…",
  office_set_font: "Formatting text…",
  office_set_paragraph_formatting: "Formatting paragraphs…",
  office_insert_table: "Inserting a table…",
  office_set_table_cell: "Editing a table cell…",
  office_get_document_text: "Reading the document…",
  office_get_outline: "Reading the outline…",
  office_set_list: "Formatting a list…",
  office_insert_image: "Inserting an image…",
  office_insert_hyperlink: "Adding a link…",
  office_insert_bookmark: "Adding a bookmark…",
  office_find: "Searching the document…",
  office_list_comments: "Reading comments…",
  office_audit_redline: "Auditing revisions…",
  office_reply_to_comment: "Replying to a comment…",
  office_resolve_comment: "Resolving a comment…",
  office_header_footer: "Editing header/footer…",
  // Excel
  excel_get_selected_range: "Reading your selection…",
  excel_list_sheets: "Listing sheets…",
  excel_read_range: "Reading cells…",
  excel_write_range: "Writing cells…",
  excel_find_value: "Searching…",
  excel_insert_rows: "Inserting rows…",
  excel_delete_rows: "Deleting rows…",
  excel_select_range: "Selecting cells…",
  excel_write_formula: "Writing formulas…",
  excel_set_format: "Formatting cells…",
  excel_insert_columns: "Inserting columns…",
  excel_delete_columns: "Deleting columns…",
  excel_add_sheet: "Adding a sheet…",
  excel_delete_sheet: "Deleting a sheet…",
  excel_rename_sheet: "Renaming a sheet…",
  excel_clear_range: "Clearing cells…",
  excel_sort_range: "Sorting…",
  excel_autofilter: "Filtering…",
  excel_create_table: "Creating a table…",
  excel_add_table_rows: "Adding table rows…",
  excel_create_chart: "Creating a chart…",
  excel_set_column_width: "Resizing columns…",
  excel_set_row_height: "Resizing rows…",
  // Common Claude Code tools
  Read: "Reading a file…",
  Write: "Writing a file…",
  Edit: "Editing a file…",
  MultiEdit: "Editing a file…",
  Bash: "Running a command…",
  Glob: "Searching files…",
  Grep: "Searching files…",
  WebFetch: "Fetching from the web…",
  WebSearch: "Searching the web…",
};
// Raw MCP server name → user-facing display name. The in-process bridge
// server is named "office" in our code, but to the user it's the active
// host's add-in — "Word" or "Excel". Computed per-call (not cached) since
// HOST is finalized in Office.onReady, after module load.
function mcpServerDisplayName(raw) {
  if (raw === "office") return hostLabel();
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}
function statusForTool(name) {
  if (TOOL_STATUS_LABELS[name]) return TOOL_STATUS_LABELS[name];
  // External and in-process MCP tools both surface with the
  // mcp__<server>__<tool> convention. Strip the prefix and retry the map
  // — catches our own office_*/excel_* tools if the SDK ever wraps them.
  const inner = /^mcp__[^_]+__(.+)$/.exec(name || "");
  if (inner && TOOL_STATUS_LABELS[inner[1]]) return TOOL_STATUS_LABELS[inner[1]];
  const m = /^mcp__([^_]+)__/.exec(name || "");
  if (m) return `Calling ${mcpServerDisplayName(m[1])}…`;
  return "Working…";
}

// Autoscroll is "pinned" only while the user is at (or near) the bottom of
// the chat. The previous behavior unconditionally snapped to bottom on
// every append (streaming deltas, events, tool chips, replay), which made
// scrolling up to read mid-stream impossible — the next token snapped you
// back. nearBottom uses a 48px tolerance so a small natural-scroll wobble
// while reading the latest message still counts as pinned.
let scrollPinned = true;
function nearBottom() {
  return $messages.scrollHeight - $messages.scrollTop - $messages.clientHeight <= 48;
}
function maybeScrollToBottom() {
  if (scrollPinned) $messages.scrollTop = $messages.scrollHeight;
}
function forceScrollToBottom() {
  scrollPinned = true;
  $messages.scrollTop = $messages.scrollHeight;
}
$messages.addEventListener(
  "scroll",
  () => {
    scrollPinned = nearBottom();
  },
  { passive: true },
);

function appendUserMessage(text) {
  processGroup = null;
  const el = document.createElement("div");
  el.className = "msg user";
  el.textContent = text;
  $messages.appendChild(el);
  // The user just sent a message — they want to see what comes next.
  // Override any prior scroll-up.
  forceScrollToBottom();
  assistantTurnElem = null;
}

function appendAssistantDelta(delta) {
  if (!delta) return;
  if (!assistantTurnElem) {
    assistantTurnElem = document.createElement("div");
    assistantTurnElem.className = "msg assistant";
    $messages.appendChild(assistantTurnElem);
  }
  assistantTurnElem.textContent += delta;
  maybeScrollToBottom();
}

function appendEvent(text) {
  const el = document.createElement("div");
  el.className = "msg event";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// Like appendEvent, but a user-facing notice that is always visible (not
// gated by the "Show diagnostics" toggle, which hides .msg.event).
function appendNotice(text) {
  const el = document.createElement("div");
  el.className = "msg notice";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// An always-visible alert bubble for errors that stop the turn (usage-limit
// hit, stream ended without a result, etc.). NOT routed through appendEvent
// because .msg.event is hidden when diagnostics-off — a user hitting their
// limit would otherwise see nothing, the turn would silently stop, and the
// status would flip back to "Ready". This bubble is rendered with a
// prominent border and a ⚠ prefix and is never hidden by the toggle.
function appendError(text) {
  const el = document.createElement("div");
  el.className = "msg error";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

const toolCards = new Map();
const assistantCards = new Map();
let processGroup = null;

// One disclosure per user turn, shared by SDK and native Office receipts.
// Keep each card's owner: a late result may arrive after the next turn starts.
function ensureProcessGroup() {
  if (processGroup) return processGroup;
  const el = document.createElement("div");
  el.className = "process-turn";
  const details = document.createElement("details");
  details.className = "process-log";
  const summary = document.createElement("summary");
  const body = document.createElement("div");
  body.className = "process-records";
  const alert = document.createElement("div");
  alert.className = "msg error process-alert";
  alert.hidden = true;
  alert.setAttribute("role", "status");
  alert.setAttribute("aria-live", "polite");
  const alertText = document.createElement("span");
  const inspect = document.createElement("button");
  inspect.textContent = "查看详情";
  inspect.addEventListener("click", () => {
    details.open = true;
    summary.focus();
  });
  alert.append(alertText, inspect);
  details.append(summary, body);
  el.append(details, alert);
  $messages.appendChild(el);
  processGroup = { el, details, summary, body, alert, alertText, cards: new Set() };
  return processGroup;
}

function refreshProcessGroup(group) {
  let pending = 0,
    running = 0,
    failed = 0,
    partial = 0,
    unconfirmed = 0;
  for (const card of group.cards) {
    if (card.state.startsWith("awaiting your confirmation")) pending++;
    else if (card.state === "running") running++;
    else if (card.state === "completed") continue;
    else if (card.state === "partially applied") partial++;
    else if (/unknown|not reported|not recorded/.test(card.state)) unconfirmed++;
    else failed++;
  }
  const statuses = [];
  if (pending) statuses.push(`${pending} 项待确认`);
  if (running) statuses.push(`${running} 项进行中`);
  if (failed) statuses.push(`${failed} 项异常`);
  if (partial) statuses.push(`${partial} 项部分完成`);
  if (unconfirmed) statuses.push(`${unconfirmed} 项结果未确认`);
  if (!statuses.length) statuses.push("记录已完成");
  group.summary.textContent = `工作过程 · ${group.cards.size} 条记录 · ${statuses.join(" / ")}`;
  // Provider tool errors do not necessarily emit a separate health notice.
  // A single visible warning per turn avoids burying them in the disclosure.
  group.alert.hidden = !failed && !unconfirmed && !partial;
  group.alertText.textContent = [
    failed ? `${failed} 项工具操作异常。` : "",
    partial ? `${partial} 项只完成了一部分，请核对文档。` : "",
    unconfirmed ? `${unconfirmed} 项工具结果未确认；这不表示执行成功。` : "",
  ].join("");
}

function displayValue(value) {
  return typeof value === "string"
    ? value
    : (JSON.stringify(
        value,
        (key, v) =>
          /^(authorization|api[_-]?key|token|password|secret|access_token|refresh_token)$/i.test(
            key,
          )
            ? "[redacted]"
            : v,
        2,
      ) ?? "");
}
function appendToolUse(name, args, id = uuid(), state = "running") {
  if (toolCards.has(id)) return toolCards.get(id);
  const el = document.createElement("details");
  el.className = "msg tool";
  const summary = document.createElement("summary");
  const input = document.createElement("pre");
  const output = document.createElement("pre");
  input.textContent = "Input\n" + displayValue(args);
  const group = ensureProcessGroup();
  const card = { el, summary, output, name, state, group };
  summary.textContent = name + " — " + state;
  el.append(summary, input, output);
  group.body.appendChild(el);
  group.cards.add(card);
  toolCards.set(id, card);
  refreshProcessGroup(group);
  maybeScrollToBottom();
  assistantTurnElem = null;
  return card;
}
function officeResultState(result) {
  if (result?.status === "partial") return "partially applied";
  if (result?.status === "not_applied") return "not applied";
  return "completed";
}
function updateToolResult(id, value, state) {
  const card = toolCards.get(id) || appendToolUse("Tool " + id, {}, id);
  state = typeof state === "string" && state ? state : "result not reported";
  card.state = state;
  card.summary.textContent = card.name + " — " + state;
  card.output.textContent = "Result\n" + displayValue(value);
  // Approval must remain outside the closed disclosure. Return its receipt to
  // the log afterwards, without changing the user's expanded/collapsed choice.
  const target = state.startsWith("awaiting your confirmation") ? $messages : card.group.body;
  if (card.el.parentNode !== target) target.appendChild(card.el);
  refreshProcessGroup(card.group);
}
function settleToolCards(state) {
  for (const [id, card] of toolCards) {
    if (card.state === "running")
      updateToolResult(id, "No confirmed final result was received.", state);
  }
}
function upsertAssistantMessage(id, text) {
  let el = assistantCards.get(id);
  if (!text && !el) return;
  if (!el || !el.isConnected) {
    el = document.createElement("div");
    el.className = "msg assistant";
    assistantCards.set(id, el);
    $messages.appendChild(el);
  }
  el.textContent = text;
  el.hidden = !text;
  maybeScrollToBottom();
}

// A complete assistant bubble (replay path — full text, not streamed
// deltas). Resets assistantTurnElem so a subsequent live delta starts a
// fresh bubble rather than appending onto a replayed one.
function appendAssistantMessage(text) {
  if (!text) return;
  const el = document.createElement("div");
  el.className = "msg assistant";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// Rebuild the chat panel from a replayed transcript. Clears whatever is in
// #messages, renders each event through the same bubble helpers the live
// path uses (so the diagnostics CSS gate applies identically), and — only
// when there is prior history — appends a divider so replayed history is
// visually distinct from the live session that follows.
function renderTranscriptReplay(events, truncated) {
  $messages.innerHTML = "";
  toolCards.clear();
  assistantCards.clear();
  assistantTurnElem = null;
  processGroup = null;

  if (truncated) {
    const t = document.createElement("div");
    t.className = "transcript-truncated";
    t.textContent = "⋯ earlier messages not shown";
    $messages.appendChild(t);
  }

  for (const ev of events) {
    if (ev.kind === "user") appendUserMessage(ev.text);
    else if (ev.kind === "assistant") appendAssistantMessage(ev.text);
    else if (ev.kind === "tool")
      appendToolUse(ev.name, ev.input, ev.id, "historical / result not recorded");
    else if (ev.kind === "tool_result") updateToolResult(ev.id, ev.output, ev.status);
  }

  if (events.length > 0) {
    const d = document.createElement("div");
    d.className = "transcript-divider";
    d.textContent = "end of earlier conversation";
    $messages.appendChild(d);
    // End of a replay is an explicit "take me to latest" — the user just
    // resumed, they want the cursor at the live tail.
    forceScrollToBottom();
  }
  processGroup = null;
}

function refreshSelectionChip() {
  if (attachSelection && lastSelection && lastSelection.text) {
    const preview =
      lastSelection.text.length > 60 ? lastSelection.text.slice(0, 57) + "..." : lastSelection.text;
    $chipText.textContent = `Selection: "${preview}"`;
    $chip.hidden = false;
  } else {
    $chip.hidden = true;
  }
}

$chipDetach.addEventListener("click", () => {
  attachSelection = false;
  refreshSelectionChip();
});

// ---------------------------------------------------------------------------
// Settings — persisted to localStorage. New settings get added here and
// applied via applySettings().
// ---------------------------------------------------------------------------
const SETTINGS_KEY = "draftspect-preview-settings-v1";

function defaultSettings() {
  return {
    showDiagnostics: false,
    trackChangesMode: "always", // "always" | "modifications" | "never"
    // Global, sticky. Picked from the connected provider catalog (see MODELS).
    model: DEFAULT_MODEL,
    modelsByRuntime: {},
  };
}

function loadSettings() {
  try {
    const raw =
      localStorage.getItem(SETTINGS_KEY) || localStorage.getItem("claude-code-office-settings-v1");
    if (!raw) return defaultSettings();
    const s = { ...defaultSettings(), ...JSON.parse(raw) };
    // Migrate any stale/retired model choice to the default so the
    // dropdown/indicator stay valid.
    if (!s.modelsByRuntime || typeof s.modelsByRuntime !== "object") s.modelsByRuntime = {};
    return s;
  } catch {
    return defaultSettings();
  }
}

function saveSettings(s) {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  } catch {
    /* ignore */
  }
}

let settings = loadSettings();

function applySettings() {
  $messages.dataset.showDiagnostics = String(settings.showDiagnostics);
  const $showDiag = document.getElementById("setting-show-diagnostics");
  if ($showDiag) $showDiag.checked = settings.showDiagnostics;
  const $tcMode = document.getElementById("setting-track-changes-mode");
  if ($tcMode) $tcMode.value = settings.trackChangesMode || "always";
  const $model = document.getElementById("composer-model");
  if ($model) $model.value = settings.model || DEFAULT_MODEL;
}

// Push the chosen model to the daemon. The SDK model is fixed per agent
// loop, so changing it mid-conversation triggers a resuming restart
// (daemon side); on first connect it's just recorded for the lazy start.
async function sendModel() {
  modelConfirmed = false;
  modelChoiceRequired = false;
  setComposerDisabled(turnInFlight);
  pendingModelShort = shortModelName(settings.model);
  try {
    const r = await sendRequest("set_model", { model: settings.model });
    if (!r.ok) throw new Error(r.error || "Model selection rejected");
    settings.model = r.model;
    settings.modelsByRuntime[activeRuntime] = r.model;
    saveSettings(settings);
    modelConfirmed = true;
    if (
      document
        .getElementById("health-notice-detail")
        ?.textContent?.startsWith("MODEL_CHOICE_REQUIRED")
    )
      showHealthNotice(null);
    liveModel = null;
  } catch (error) {
    appendNotice(`Model not confirmed: ${error.message}. Choose a model to retry.`);
  } finally {
    pendingModelShort = null;
    applySettings();
    renderConnection();
    setComposerDisabled(turnInFlight);
    refreshReadyStatus();
  }
}

// Ask the daemon for the ACTIVE runtime's model catalog (authoritative —
// static for the Claude runtime, discovered for CodeBuddy).
function requestModelList() {
  if (!wsReady) return;
  modelConfirmed = false;
  modelChoiceRequired = false;
  setComposerDisabled(turnInFlight);
  const owner = ws;
  sendRequest("list_models")
    .then((result) => {
      if (owner === ws && wsReady) handleServerMessage({ ...result, request_id: undefined });
    })
    .catch((error) => {
      if (owner === ws && wsReady) {
        showHealthNotice(`模型列表加载失败：${error.message}`, "MODEL_CATALOG_UNAVAILABLE", true);
        refreshReadyStatus();
      }
    });
}

// Populate the dropdown from the daemon's catalog. Keeps the stored choice
// when it's still offered; otherwise migrates to the first entry. That
// migration is how a legacy provider id becomes a CodeBuddy id after a runtime
// switch — visible in the dropdown and pushed to the daemon, never silent.
function populateModelOptions(models, runtimeId) {
  const $model = document.getElementById("composer-model");
  if (!$model || !Array.isArray(models) || models.length === 0) return;
  setComposerDisabled(turnInFlight);
  $model.innerHTML = "";
  for (const m of models) {
    const opt = document.createElement("option");
    opt.value = m.value;
    opt.textContent = m.label || shortModelName(m.value);
    $model.appendChild(opt);
  }
  const current = settings.modelsByRuntime[runtimeId] || settings.model;
  if (models.some((m) => m.value === current)) {
    settings.model = current;
    $model.value = current;
    sendModel();
    return;
  }
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Choose a model…";
  $model.prepend(placeholder);
  $model.value = "";
  settings.model = "";
  modelConfirmed = false;
  modelChoiceRequired = true;
  showHealthNotice(
    "上次选择的模型已不在当前目录中。请在输入框下方重新选择可用模型；尚未提交模型请求。",
    "MODEL_CHOICE_REQUIRED",
  );
  appendNotice(
    `Choose a ${runtimeId} model. The previous choice (${current}) is not offered by this runtime.`,
  );
  renderConnection();
  refreshReadyStatus();
}

document.getElementById("setting-show-diagnostics").addEventListener("change", (e) => {
  settings.showDiagnostics = e.target.checked;
  saveSettings(settings);
  applySettings();
});

// Word-only control; the Excel pane omits this element entirely.
document.getElementById("setting-track-changes-mode")?.addEventListener("change", (e) => {
  settings.trackChangesMode = e.target.value;
  saveSettings(settings);
  applySettings();
  // Push to daemon so the agent sees the new mode on the next turn.
  if (wsReady) {
    wsSend({ type: "context_update", track_changes_mode: settings.trackChangesMode });
  }
});

document.getElementById("composer-model")?.addEventListener("change", (e) => {
  settings.model = e.target.value;
  saveSettings(settings);
  applySettings();
  sendModel();
  // If a loop is already live (a turn has run), the daemon will restart it
  // to apply the new model — show that until the next session_init confirms.
  // If nothing has run yet, renderConnection already shows the new choice.
  if (liveModel) pendingModelShort = shortModelName(settings.model);
  renderConnection();
});

applySettings();

// ---------------------------------------------------------------------------
// WebSocket
// ---------------------------------------------------------------------------
function wsConnect() {
  activeRuntime = null;
  renderRuntimeCopy();
  setConnectionStatus("idle", "Connecting…");
  refreshReadyStatus();
  ws = new WebSocket(WS_URL);
  const owner = ws;

  ws.onopen = () => {
    if (ws !== owner) return;
    wsReady = true;
    modelConfirmed = false;
    replayReady = false;
    setComposerDisabled(true);
    refreshReadyStatus();
    sendHello();
    // Record the sticky model for this pane key right after the hello
    // binds it, so the lazy first-message session start uses it.
  };

  ws.onclose = () => {
    if (ws !== owner) return;
    connectionInterruptedTurn ||= turnInFlight;
    clearTimeout(healthTimer);
    healthSupported = false;
    confirmWindowSupported = false;
    connectionEpoch = null;
    healthBlocked = false;
    wsReady = false;
    activeRuntime = null;
    renderRuntimeCopy();
    modelConfirmed = false;
    replayReady = false;
    cancelOfficeOperations();
    settleToolCards("interrupted / outcome unknown");
    for (const p of pendingRequests.values()) {
      clearTimeout(p.timer);
      p.reject(new Error("Connection closed"));
    }
    pendingRequests.clear();
    setConnectionStatus("err", "Disconnected — retrying…");
    showHealthNotice(
      `本地连接不可用，原因尚未确认。请查看 Draftspect 桌面状态与本地诊断记录；侧栏会尝试重新连接，原请求和文档修改不会自动重发。写入结果不明时先重新读取并核对 ${hostLabel()}。`,
      "CONNECTION_UNAVAILABLE",
    );
    // Release the composer if a turn was mid-flight when the connection
    // dropped — otherwise the user is stuck waiting for a turn_complete
    // that will never arrive.
    endTurn();
    refreshReadyStatus();
    // Daemon may have been restarted, in which case it has rotated the
    // bridge token. Re-fetch from /bridge-token before each reconnect
    // attempt so the next hello carries the current token. fetchBridgeToken
    // tolerates the HTTP server being briefly unreachable too (sets the
    // token to null, the hello fails, we loop again).
    setTimeout(() => {
      fetchBridgeToken().then(wsConnect);
    }, 1500);
  };

  ws.onerror = (err) => {
    console.warn("WS error:", err);
  };

  ws.onmessage = (evt) => {
    if (ws !== owner) return;
    let msg;
    try {
      msg = JSON.parse(evt.data);
    } catch {
      return;
    }
    handleServerMessage(msg).catch(() => {
      healthBlocked = true;
      cancelOfficeOperations();
      showHealthNotice(
        "处理服务端消息时发生异常。已暂停新的文档操作，请检查状态或重新连接。",
        "PANE_MESSAGE_ERROR",
      );
      setComposerDisabled(turnInFlight);
    });
  };
}

function wsSend(obj) {
  if (!wsReady) return;
  ws.send(JSON.stringify(obj));
}

// Stable per-pane-instance id. sessionStorage persists for the lifetime
// of this task pane (survives WS reconnects and panel reloads, cleared
// when the pane is closed), so an unsaved/cloud doc with no filesystem
// path still gets a STABLE bridge key across the 1.5s reconnect loop —
// without this the bridge minted a fresh random key every retry and
// leaked per-pane state. Falls back to an in-memory id if
// sessionStorage is unavailable.
let panePersistId;
try {
  panePersistId = sessionStorage.getItem("cc-pane-id");
  if (!panePersistId) {
    panePersistId = crypto?.randomUUID?.() ?? "p_" + Math.random().toString(36).slice(2, 12);
    sessionStorage.setItem("cc-pane-id", panePersistId);
  }
} catch {
  panePersistId = crypto?.randomUUID?.() ?? "p_" + Math.random().toString(36).slice(2, 12);
}

function sendHello() {
  activeDocUrl = Office.context.document.url || null;
  wsSend({
    type: "hello",
    protocol_version: 2,
    token: bridgeToken,
    host: HOST,
    active_doc: activeDocUrl,
    pane_id: panePersistId,
    selection: attachSelection ? lastSelection : null,
    track_changes_mode: settings.trackChangesMode || "always",
  });
}

async function handleServerMessage(msg) {
  const pending = msg.request_id && pendingRequests.get(msg.request_id);
  if (pending) {
    if (msg.type !== pending.type + "_result") return;
    pendingRequests.delete(msg.request_id);
    clearTimeout(pending.timer);
    pending.resolve(msg);
    return;
  }
  switch (msg.type) {
    case "welcome":
      if (msg.protocol_version !== 2) {
        appendError(
          "This panel needs the matching Draftspect service version. Close the older installation and start the current Draftspect service.",
        );
        return;
      }
      activeRuntime = msg.runtime;
      connectionEpoch = msg.connection_epoch || null;
      healthSupported = !!connectionEpoch && msg.capabilities?.includes("turn_health_v1");
      confirmWindowSupported = !!msg.capabilities?.includes("confirm_window_v1");
      healthFailures = 0;
      healthBlocked = false;
      renderRuntimeCopy();
      setConnectionStatus("ok", "Connected");
      setAgentStatus("idle", "正在检查模型与历史…");
      refreshWorkspaceFromDaemon();
      // Fetch the active runtime's model catalog (authoritative dropdown).
      requestModelList();
      scheduleHealthCheck();
      break;

    case "transcript_replay":
      receiveSessionInfo(
        msg.session_info || { session_id: msg.session_id, provider: activeRuntime, state: "word" },
      );
      if (!turnInFlight) renderTranscriptReplay(msg.events || [], !!msg.truncated);
      if (msg.error) appendNotice(msg.error);
      if (msg.legacy_history_available)
        appendNotice(
          "Legacy folder-level history is retained separately. It was not automatically attached to this document, to prevent cross-document history mixing.",
        );
      replayReady = true;
      setComposerDisabled(turnInFlight);
      refreshReadyStatus();
      break;

    case "assistant_text":
      if (msg.turn_id && msg.turn_id !== currentTurnId) break;
      appendAssistantDelta(msg.delta);
      break;

    case "assistant_event":
      if (msg.turn_id && msg.turn_id !== currentTurnId) break;
      if (msg.event === "assistant_message") {
        upsertAssistantMessage(msg.message_id, msg.text);
      } else if (msg.event === "tool_result_announce") {
        updateToolResult(msg.tool_id, msg.output, msg.status);
      } else if (msg.event === "tool_use_announce") {
        appendToolUse(msg.tool, msg.input, msg.tool_id);
        setAgentStatus("working", statusForTool(msg.tool));
      } else if (msg.event === "turn_complete") {
        if (msg.error) appendError(msg.error);
        if (msg.auth) showAuthErrorBanner(msg.error);
        settleToolCards(msg.interrupted ? "interrupted / outcome unknown" : "result not reported");
        cancelOfficeOperations();
        setAgentStatus(
          "idle",
          msg.error ? "Stopped — see message" : msg.interrupted ? "Stopped" : "Ready",
        );
        endTurn();
        if (
          (!msg.error && !msg.interrupted) ||
          operationStorage?.getItem(inspectionKey(activeDocUrl)) ||
          persistenceWarning
        )
          refreshReadyStatus();
      } else if (msg.event === "retry") {
        // Transport-level retry in flight (network blip / upstream error).
        // Status-only signal — deliberately no chat bubble, or a retry
        // storm would spam the conversation.
        setAgentStatus("working", "Retrying…");
      } else if (msg.event === "info") {
        appendNotice(msg.message);
      } else if (msg.event === "persistence_error") {
        persistenceWarning = msg.code || "SESSION_PERSISTENCE_FAILED";
        showHealthNotice(
          "会话关联保存失败。模型回复仍可能完成，但不能确认之后能恢复这段历史。请先检查会话存储。",
          persistenceWarning,
        );
        renderSessionControls();
      } else if (msg.event === "persistence_saved") {
        persistenceWarning = null;
        if (lastHealthNotice?.includes("会话")) showHealthNotice(null);
        renderSessionControls();
        refreshReadyStatus();
      } else if (msg.event === "error") {
        // Always-visible bubble (NOT .msg.event, which the diagnostics
        // toggle hides). Status reads "Stopped — see message" so the
        // user knows the turn ended on this error, not normally.
        appendError(msg.error || "The agent stopped without producing a result.");
        setAgentStatus("idle", "Stopped — see message");
        endTurn();
      } else if (msg.event === "auth_error") {
        showAuthErrorBanner(msg.error);
        if (wsReady) setConnectionStatus("err", "Sign-in required");
        endTurn();
      } else if (msg.event === "session_init") {
        receiveSessionInfo({
          ...currentSessionInfo,
          session_id: msg.session_id,
          provider: activeRuntime,
          state: "working",
        });
        appendEvent(`Session ${msg.session_id?.slice(0, 8)}… (${msg.model})`);
        // Authoritative: this is the model the SDK actually started with.
        liveModel = msg.model || liveModel;
        pendingModelShort = null;
        renderConnection();
      } else if (msg.event === "cwd_changed") {
        setWorkspaceDisplay(msg.cwd);
        appendEvent(
          `Switched to workspace: ${msg.cwd.split(/[\\/]/).filter(Boolean).pop()}${msg.resumed ? " (resumed prior session)" : ""}`,
        );
        // Per-workspace context must be re-read for the new workspace.
        contextCache = null;
        if (document.body.dataset.activeTab === "setup") loadContext(true);
      } else if (msg.event === "config_reloaded") {
        const what = msg.reason === "context_changed" ? "context files" : "config";
        appendEvent(`Session reloaded — ${what} updated.`);
      }
      break;

    case "list_models_result":
      if (!msg.ok || !msg.models?.length) {
        modelConfirmed = false;
        showHealthNotice(
          "模型列表不可用。请点击重试加载，尚未提交新的模型请求。",
          "MODEL_CATALOG_UNAVAILABLE",
          true,
        );
        setComposerDisabled(true);
        break;
      }
      // Authoritative catalog for the active runtime (see populateModelOptions).
      populateModelOptions(msg.models || [], msg.runtime);
      if (document.getElementById("health-notice-detail")?.textContent?.startsWith("MODEL_CATALOG"))
        showHealthNotice(null);
      break;

    case "set_model_result":
      if (!msg.ok) {
        // The daemon rejected the choice (e.g. a stale id after a runtime
        // switch, before the fresh catalog arrives). Surface it and roll the
        // dropdown back to the model the daemon is actually using.
        appendNotice(msg.error || "Model switch rejected.");
        if (msg.model) {
          settings.model = msg.model;
          saveSettings(settings);
          applySettings();
        }
        renderConnection();
      }
      break;

    case "tool_call":
      runOfficeTool(msg);
      break;

    case "cancel_office_tools":
      cancelOfficeOperations();
      break;

    case "pong":
      break;
    case "protocol_error":
      healthBlocked = true;
      cancelOfficeOperations();
      showHealthNotice(
        msg.message || "请求协议不兼容，请重新打开侧栏。",
        msg.code || "PROTOCOL_ERROR",
      );
      setComposerDisabled(turnInFlight);
      checkTurnHealth();
      break;

    default:
      // Request/response messages keyed by request_id end in "_result".
      // Resolve the matching pending request.
      if (typeof msg.type === "string" && msg.type.endsWith("_result") && msg.request_id) {
        const pending = pendingRequests.get(msg.request_id);
        if (pending) {
          pendingRequests.delete(msg.request_id);
          pending.resolve(msg);
        }
      }
      break;
  }
}

// ---- Request/response helper (for non-tool round-trips) -------------------
const pendingRequests = new Map();
const REQUEST_TIMEOUT_MS = 10_000;

function sendRequest(type, payload = {}) {
  return new Promise((resolve, reject) => {
    if (!wsReady) {
      reject(new Error("Not connected to daemon"));
      return;
    }
    const request_id = uuid();
    const pending = { resolve, reject, type, timer: null };
    pendingRequests.set(request_id, pending);
    wsSend({ type, request_id, ...payload });
    pending.timer = setTimeout(
      () => {
        if (pendingRequests.has(request_id)) {
          pendingRequests.delete(request_id);
          reject(new Error(`Request "${type}" timed out`));
        }
      },
      type === "release_session" ? 60_000 : type === "list_models" ? 30000 : REQUEST_TIMEOUT_MS,
    );
  });
}

// ---------------------------------------------------------------------------
// Tool dispatcher
// ---------------------------------------------------------------------------
// Apply the user's track-changes mode setting on top of whatever the agent
// passed. The agent's schema-default value can't be trusted as the final
// arbiter: the user's mode setting is the policy.
//   - "always": force true regardless of what the agent passed.
//   - "never":  do not turn tracking on. Word's existing Track Changes state
//               is left alone, so an already-tracked document still records
//               the edit. This does not switch tracking off for the write.
//               office_replace_text does not use this flag: a word or phrase
//               replacement always leaves revision marks.
//   - "modifications": honor the agent's flag — the system prompt tells the
//                      agent to pass false only for fresh-section drafts.
function effectiveTrackChanges(provided) {
  const mode = (typeof settings !== "undefined" && settings?.trackChangesMode) || "always";
  if (mode === "never") return false;
  if (mode === "always") return true;
  return provided !== false;
}

async function executeOfficeTool(msg, control) {
  const { id, name, args = {} } = msg;
  try {
    // Host guard: refuse wrong-host tools with a clear message so the
    // agent can self-correct on its next turn. The daemon registers both
    // tool families on every session, so this is the only place we can
    // catch mismatches.
    if (HOST === "excel" && name.startsWith("office_")) {
      throw new Error(
        `Tool ${name} is Word-only; the active host is Excel. Use excel_* tools instead.`,
      );
    }
    if (HOST === "word" && name.startsWith("excel_")) {
      throw new Error(
        `Tool ${name} is Excel-only; the active host is Word. Use office_* tools instead.`,
      );
    }
    let result;
    // Apply the user's track-changes mode to every write-tool call before
    // dispatching, so user preference always wins over the agent's value.
    const writeArgs = () => ({
      ...args,
      track_changes: effectiveTrackChanges(args && args.track_changes),
    });
    switch (name) {
      case "office_document_map":
        result = await toolDocumentMap();
        break;
      case "office_read_story":
        result = await toolReadStory(args);
        break;
      case "office_read_formatting":
        result = await toolReadFormatting(args);
        break;
      case "office_read_revisions":
        result = await toolReadRevisions(args);
        break;
      case "office_audit_redline":
        result = await toolAuditRedline(args);
        break;
      case "office_get_selection":
        result = await toolGetSelection();
        break;
      case "office_read_paragraphs":
        result = await toolReadParagraphs(args);
        break;
      case "office_insert_paragraphs":
        result = await toolInsertParagraphs(writeArgs(), control);
        break;
      case "office_replace_paragraphs":
        result = await toolReplaceParagraphs(writeArgs(), control);
        break;
      case "office_replace_section":
        result = await toolReplaceSection(writeArgs(), control);
        break;
      case "office_delete_paragraphs":
        result = await toolDeleteParagraphs(writeArgs(), control);
        break;
      case "office_replace_text":
        // Only the confirmed prepareTextReplacement/toolApplyTextReplacement path may write.
        throw new Error("office_replace_text must use the confirmed surgical replacement path.");
      case "office_highlight":
        result = await toolHighlight(args, control);
        break;
      case "office_clear_highlights":
        result = await toolClearHighlights(args, control);
        break;
      case "office_add_comment":
        result = await toolAddComment(args);
        break;
      case "office_clear_comments":
        result = await toolClearComments(args, control);
        break;
      case "office_apply_style":
        result = await toolApplyStyle(writeArgs(), control);
        break;
      case "office_set_font":
        result = await toolSetFont(writeArgs(), control);
        break;
      case "office_set_paragraph_formatting":
        result = await toolSetParagraphFormatting(writeArgs(), control);
        break;
      case "office_insert_table":
        result = await toolInsertTable(writeArgs(), control);
        break;
      case "office_set_table_cell":
        result = await toolSetTableCell(writeArgs(), control);
        break;
      case "office_get_document_text":
        result = await toolGetDocumentText();
        break;
      case "office_get_outline":
        result = await toolGetOutline();
        break;
      case "office_set_list":
        result = await toolSetList(writeArgs(), control);
        break;
      case "office_insert_image":
        result = await toolInsertImage(writeArgs(), control);
        break;
      case "office_insert_hyperlink":
        result = await toolInsertHyperlink(writeArgs(), control);
        break;
      case "office_insert_bookmark":
        result = await toolInsertBookmark(args, control);
        break;
      case "office_find":
        result = await toolFind(args);
        break;
      case "office_list_comments":
        result = await toolListComments(args);
        break;
      case "office_reply_to_comment":
        result = await toolReplyToComment(args);
        break;
      case "office_resolve_comment":
        result = await toolResolveComment(args);
        break;
      case "office_header_footer":
        result = await toolHeaderFooter(args, control);
        break;
      // ---- Excel ----
      case "excel_get_selected_range":
        result = await toolExcelGetSelectedRange();
        break;
      case "excel_list_sheets":
        result = await toolExcelListSheets();
        break;
      case "excel_read_range":
        result = await toolExcelReadRange(args);
        break;
      case "excel_write_range":
        result = await toolExcelWriteRange(args, control);
        break;
      case "excel_find_value":
        result = await toolExcelFindValue(args);
        break;
      case "excel_insert_rows":
        result = await toolExcelInsertRows(args, control);
        break;
      case "excel_delete_rows":
        result = await toolExcelDeleteRows(args, control);
        break;
      case "excel_select_range":
        result = await toolExcelSelectRange(args);
        break;
      case "excel_write_formula":
        result = await toolExcelWriteFormula(args, control);
        break;
      case "excel_set_format":
        result = await toolExcelSetFormat(args, control);
        break;
      case "excel_insert_columns":
        result = await toolExcelInsertColumns(args, control);
        break;
      case "excel_delete_columns":
        result = await toolExcelDeleteColumns(args, control);
        break;
      case "excel_add_sheet":
        result = await toolExcelAddSheet(args, control);
        break;
      case "excel_delete_sheet":
        result = await toolExcelDeleteSheet(args, control);
        break;
      case "excel_rename_sheet":
        result = await toolExcelRenameSheet(args, control);
        break;
      case "excel_clear_range":
        result = await toolExcelClearRange(args, control);
        break;
      case "excel_sort_range":
        result = await toolExcelSortRange(args, control);
        break;
      case "excel_autofilter":
        result = await toolExcelAutoFilter(args, control);
        break;
      case "excel_create_table":
        result = await toolExcelCreateTable(args, control);
        break;
      case "excel_add_table_rows":
        result = await toolExcelAddTableRows(args, control);
        break;
      case "excel_create_chart":
        result = await toolExcelCreateChart(args, control);
        break;
      case "excel_set_column_width":
        result = await toolExcelSetColumnWidth(args, control);
        break;
      case "excel_set_row_height":
        result = await toolExcelSetRowHeight(args, control);
        break;
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
    return result;
  } catch (err) {
    console.error(`[tool ${name}] failed:`, err);
    throw err;
  }
}

const readOnlyOfficeTools = new Set([
  "office_document_map",
  "office_read_story",
  "office_read_formatting",
  "office_read_revisions",
  "office_audit_redline",
  "office_get_selection",
  "office_read_paragraphs",
  "office_get_document_text",
  "office_get_outline",
  "office_find",
  "office_list_comments",
  "excel_get_selected_range",
  "excel_list_sheets",
  "excel_read_range",
  "excel_find_value",
  "excel_select_range",
]);
let operationStorage;
try {
  operationStorage = sessionStorage;
} catch {
  operationStorage = null;
}
const officeQueue = createOperationQueue({ storage: operationStorage });
const confirmations = new Map();
const inFlightWrites = new Set();
const inspectionCards = new Map();
const inspectionKey = (url) => "draftspect-inspect-before-write:" + (url || panePersistId);
function offerInspectionAcknowledgement(url) {
  if (inspectionCards.get(url)?.isConnected) return;
  const card = document.createElement("div");
  card.className = "msg error";
  card.textContent =
    "A prior edit has an uncertain outcome. Inspect and reread the live document and its revisions before authorizing more writes. This action does not undo anything. ";
  const button = document.createElement("button");
  button.textContent = `I inspected and reread ${hostLabel()} — allow new proposals`;
  button.addEventListener("click", () => {
    if (inFlightWrites.size) {
      appendNotice(
        "An Office operation is still returning. Wait for it or reopen the panel after inspecting the document.",
      );
      return;
    }
    operationStorage?.setItem(inspectionKey(url), "");
    card.remove();
    inspectionCards.delete(url);
  });
  card.appendChild(button);
  $messages.appendChild(card);
  inspectionCards.set(url, card);
}
function cancelOfficeOperations() {
  officeQueue.cancel();
  if (inFlightWrites.size) {
    // A final Office sync may never return after the connection is lost. Mark
    // uncertainty now; do not wait for a late success or fabricate a rollback.
    try {
      operationStorage?.setItem(inspectionKey(activeDocUrl), "unknown");
    } catch {}
    offerInspectionAcknowledgement(activeDocUrl);
  }
  for (const cancel of confirmations.values()) cancel();
  confirmations.clear();
}

function confirmOfficeChange(msg, preview, guard) {
  const card = toolCards.get(msg.id);
  // A confirm_window_v1 daemon re-arms the tool deadline when the card opens
  // (45s to decide plus one minute for commit/readback). Older daemons keep a
  // fixed deadline, so reserve that minute out of what remains.
  const windowMs = confirmWindowSupported
    ? 45000
    : Math.max(0, Math.min(45000, msg.deadline - Date.now() - 60000));
  if (!windowMs)
    throw new Error("CONFIRMATION_EXPIRED: insufficient time remains; no write authorized.");
  if (confirmWindowSupported && msg.deadlineRef)
    msg.deadlineRef.at = Math.max(msg.deadlineRef.at, Date.now() + windowMs + 60000);
  card.el.open = true;
  updateToolResult(
    msg.id,
    preview,
    `awaiting your confirmation (${Math.ceil(windowMs / 1000)} seconds)`,
  );
  setAgentStatus("working", `等待你确认修改 · 最多 ${Math.ceil(windowMs / 1000)} 秒`);
  wsSend({ type: "tool_phase", id: msg.id, turn_id: msg.turn_id, phase: "awaiting_confirmation" });
  return new Promise((resolve, reject) => {
    const approve = document.createElement("button");
    approve.textContent = "Apply this change";
    const deny = document.createElement("button");
    deny.textContent = "Cancel";
    let settled = false;
    let timer;
    const finish = (ok) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      confirmations.delete(msg.id);
      approve.remove();
      deny.remove();
      try {
        guard();
        if (!ok) throw new Error("USER_DECLINED_OR_EXPIRED: no edit was authorized.");
        resolve();
      } catch (error) {
        reject(error);
      }
    };
    approve.addEventListener("click", () => finish(true));
    deny.addEventListener("click", () => finish(false));
    confirmations.set(msg.id, () => finish(false));
    timer = setTimeout(() => finish(false), windowMs);
    card.el.append(approve, deny);
    maybeScrollToBottom();
  });
}

function excelRecheckWarning(preview) {
  if (preview?.recheck === "formulas" || preview?.recheck === "formulas_chunked") {
    const scope =
      preview.recheck === "formulas_chunked"
        ? "Values and formulas across the whole used range are compared again in parts after you approve."
        : "Values and formulas in the target are compared again after you approve.";
    return `${scope} Formatting is not compared, and the result still needs review.`;
  }
  return "The sheet identity is rechecked after you approve. This operation does not compare every cell, and the result needs review.";
}

async function runOfficeTool(msg) {
  const owner = ws;
  const documentUrl = activeDocUrl;
  const mutates = !readOnlyOfficeTools.has(msg.name);
  appendToolUse("Office execution: " + msg.name, msg.args, msg.id);
  let writeStarted = false;
  try {
    if (currentTurnId !== msg.turn_id || !currentTurnId)
      throw new Error("TURN_MISMATCH: this tool does not belong to the current request.");
    if (!wsReady || healthBlocked)
      throw new Error("CONNECTION_UNAVAILABLE: document operation refused.");
    if ((Office.context.document.url || null) !== documentUrl)
      throw new Error("DOCUMENT_CHANGED: reconnect and reread before requesting an operation.");
    const deadline = { at: msg.deadline || Date.now() };
    msg = { ...msg, deadlineRef: deadline };
    const result = await officeQueue.run(
      msg.id,
      {
        mutates,
        deadline: () => deadline.at,
        valid: () =>
          ws === owner &&
          wsReady &&
          !healthBlocked &&
          !!currentTurnId &&
          currentTurnId === msg.turn_id &&
          (Office.context.document.url || null) === documentUrl,
      },
      async ({ guard, beforeWrite: markWrite }) => {
        const beforeWrite = () => {
          markWrite();
          operationStorage.setItem(inspectionKey(documentUrl), "pending");
          writeStarted = true;
          inFlightWrites.add(msg.id);
          setAgentStatus("working", `正在写入 ${hostLabel()} · 等待读回验证`);
          wsSend({ type: "tool_phase", id: msg.id, turn_id: msg.turn_id, phase: "office_write" });
        };
        if (!mutates) {
          wsSend({ type: "tool_phase", id: msg.id, turn_id: msg.turn_id, phase: "office_read" });
          return executeOfficeTool(msg);
        }
        if (HOST === "word")
          msg = { ...msg, args: await resolveWordHeadingArgs(msg.name, msg.args || {}) };
        if (!operationStorage)
          throw new Error("Write receipt storage is unavailable; reopen the panel before editing.");
        if (operationStorage.getItem(inspectionKey(documentUrl))) {
          offerInspectionAcknowledgement(documentUrl);
          throw new Error(
            "RESULT_UNKNOWN: inspect the prior edit and acknowledge it in the pane before requesting another write.",
          );
        }
        let reviewName = msg.name,
          reviewArgs = msg.args || {};
        if (msg.name === "office_add_comment") {
          reviewName = "office_manage_comment";
          reviewArgs = {
            action: "add",
            target: {
              story: "body",
              paragraph_id: msg.args.paragraph_id,
              ...(msg.args.query ? { query: msg.args.query } : {}),
            },
            text: msg.args.text,
          };
        } else if (
          msg.name === "office_reply_to_comment" ||
          msg.name === "office_resolve_comment"
        ) {
          reviewName = "office_manage_comment";
          reviewArgs = {
            ...msg.args,
            action:
              msg.name === "office_reply_to_comment"
                ? "reply"
                : msg.args.resolved === false
                  ? "reopen"
                  : "resolve",
          };
        }
        if (
          ["office_edit_range", "office_manage_comment", "office_manage_revisions"].includes(
            reviewName,
          )
        ) {
          const prepared = await prepareReviewWrite(reviewName, {
            ...reviewArgs,
            track_changes: effectiveTrackChanges(reviewArgs.track_changes),
          });
          guard();
          await confirmOfficeChange(msg, { document: documentUrl, ...prepared.preview }, guard);
          guard();
          return prepared.apply({ guard, beforeWrite });
        }
        const safe =
          HOST === "word"
            ? await prepareSafeWordWrite(msg.name, {
                ...msg.args,
                track_changes: effectiveTrackChanges(msg.args?.track_changes),
              })
            : null;
        if (safe) {
          guard();
          await confirmOfficeChange(msg, { document: documentUrl, ...safe.preview }, guard);
          guard();
          return safe.apply({ guard, beforeWrite });
        }
        const minimal = HOST === "word" && msg.name === "office_replace_text";
        const before = HOST === "word" ? await toolReadParagraphs({ preview: false }) : null;
        const formatBefore =
          HOST === "word"
            ? await captureWordFormattingPrecondition(msg.name, msg.args || {})
            : null;
        const excelBefore =
          HOST === "excel" ? await captureExcelPrecondition(msg.name, msg.args || {}) : null;
        if (excelBefore) msg = { ...msg, args: excelBefore.args };
        const commentBefore =
          msg.name === "office_clear_comments" ? await toolReadComments() : null;
        const plan = minimal ? await prepareTextReplacement(msg.args) : null;
        guard();
        await confirmOfficeChange(
          msg,
          minimal
            ? { document: documentUrl, tracked_changes: true, changes: plan }
            : {
                document: documentUrl,
                operation: msg.name,
                args: msg.args,
                ...(commentBefore ? { affected_comment_inventory: commentBefore.comments } : {}),
                ...(excelBefore
                  ? {
                      ...excelBefore.preview,
                      warning: excelRecheckWarning(excelBefore.preview),
                    }
                  : {
                      tracked_changes: effectiveTrackChanges(msg.args?.track_changes),
                      warning:
                        (settings?.trackChangesMode === "never"
                          ? "Track Changes will not be turned on. If this document already has it on, this edit is still recorded. "
                          : "") +
                        (formatBefore
                          ? "Target paragraph style and OOXML formatting will be rechecked before writing. Native visual review is still required."
                          : "Broad operation: body text/style preconditions only. Other formatting, headers and non-body changes require manual review."),
                    }),
              },
          guard,
        );
        guard();
        if (minimal) return toolApplyTextReplacement(msg.args, plan, { guard, beforeWrite });
        if (excelBefore) {
          const latest = await captureExcelPrecondition(msg.name, msg.args);
          if (JSON.stringify(latest.state) !== JSON.stringify(excelBefore.state))
            throw new Error(
              "STALE_PROPOSAL: target cells or sheets changed after confirmation preview; propose again.",
            );
        }
        if (before) {
          const latest = await toolReadParagraphs({ preview: false });
          if (JSON.stringify(latest.paragraphs) !== JSON.stringify(before.paragraphs))
            throw new Error(
              "STALE_PROPOSAL: body changed after confirmation preview; propose again.",
            );
        }
        if (formatBefore) {
          const latest = await captureWordFormattingPrecondition(msg.name, msg.args || {});
          if (JSON.stringify(latest) !== JSON.stringify(formatBefore))
            throw new Error(
              "STALE_PROPOSAL: target formatting changed after confirmation preview; propose again.",
            );
        }
        if (
          commentBefore &&
          JSON.stringify(await toolReadComments()) !== JSON.stringify(commentBefore)
        )
          throw new Error(
            "STALE_PROPOSAL: comment threads or anchors changed; read and propose again.",
          );
        const applied = await executeOfficeTool(msg, {
          guard,
          beforeWrite,
          formatExpected: formatBefore,
        });
        if (commentBefore) {
          const after = await toolReadComments();
          if (after.comment_count !== commentBefore.comment_count - applied.cleared_comments)
            throw new Error("RESULT_UNKNOWN: comment cleanup readback differs");
        }
        const readback = HOST === "word" ? await toolGetDocumentText() : (applied.readback ?? null);
        return {
          ...applied,
          verification:
            applied.verification ||
            "operation_returned; manual structural/visual verification required",
          readback,
        };
      },
    );
    if (writeStarted && currentTurnId === msg.turn_id && owner === ws)
      operationStorage.setItem(inspectionKey(documentUrl), "");
    updateToolResult(msg.id, result, officeResultState(result));
    if (owner === ws && wsReady) wsSend({ type: "tool_result", id: msg.id, ok: true, result });
  } catch (error) {
    const detail =
      (writeStarted && !String(error.message).includes("RESULT_UNKNOWN")
        ? "RESULT_UNKNOWN: "
        : "") + error.message;
    const uncertain = writeStarted || detail.includes("RESULT_UNKNOWN");
    const hostName = hostLabel();
    updateToolResult(
      msg.id,
      detail,
      uncertain ? `outcome unknown — inspect ${hostName}` : "not applied / error",
    );
    showHealthNotice(
      uncertain
        ? `${hostName} 修改结果待核实。请先检查文档；系统不会自动重试写入。`
        : `${hostName} 操作未完成：${detail}`,
      uncertain ? "RESULT_UNKNOWN" : detail.split(":", 1)[0],
    );
    if (uncertain) {
      try {
        operationStorage?.setItem(inspectionKey(documentUrl), "unknown");
      } catch {}
      offerInspectionAcknowledgement(documentUrl);
    }
    if (owner === ws && wsReady)
      wsSend({ type: "tool_result", id: msg.id, ok: false, error: detail });
  } finally {
    inFlightWrites.delete(msg.id);
  }
}

// ---------------------------------------------------------------------------
// Selection tracking — push context_update on changes (debounced).
// ---------------------------------------------------------------------------
let selectionDebounce = null;

async function captureSelection() {
  try {
    let r;
    if (HOST === "excel") {
      r = await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address, values, rowCount, columnCount");
        await context.sync();
        const cellCount = (range.rowCount || 0) * (range.columnCount || 0);
        const text =
          cellCount === 1
            ? String(range.values?.[0]?.[0] ?? "")
            : `${range.address} (${range.rowCount}×${range.columnCount})`;
        return { text, address: range.address };
      });
    } else {
      r = await Word.run(async (context) => {
        const sel = context.document.getSelection();
        const current = sel.getReviewedText(Word.ChangeTrackingVersion.current);
        await context.sync();
        return { text: current.value, para_id: null, para_count: null };
      });
    }
    lastSelection = r;
    refreshSelectionChip();
    if (wsReady) {
      wsSend({
        type: "context_update",
        active_doc: Office.context.document.url || null,
        selection: attachSelection ? lastSelection : null,
      });
    }
  } catch (e) {
    // Selection may be transient; ignore.
  }
}

function onSelectionChanged() {
  if (selectionDebounce) clearTimeout(selectionDebounce);
  selectionDebounce = setTimeout(captureSelection, 100);
}

// ---------------------------------------------------------------------------
// Composer
// ---------------------------------------------------------------------------
function sendUserMessage(rawText) {
  // Gate: refuse new submissions while a turn is still streaming. Without
  // this, appendUserMessage resets the assistant bubble pointer and the
  // remaining deltas of the prior turn would mis-order into a new bubble.
  if (
    turnInFlight ||
    !modelConfirmed ||
    !replayReady ||
    !wsReady ||
    sessionReleased ||
    handoffBusy ||
    healthBlocked
  ) {
    showHealthNotice(
      "当前尚未就绪，或上一请求／会话交接仍未结束。请先检查状态。",
      "SEND_NOT_READY",
    );
    return false;
  }
  if ((Office.context.document.url || null) !== activeDocUrl) {
    activeDocUrl = Office.context.document.url || null;
    ws.close();
    appendNotice("Document identity changed. Reconnecting before accepting another request.");
    return false;
  }
  const text = String(rawText || "").trim();
  if (!text) return false;
  appendUserMessage(text);
  currentTurnId = uuid();
  const turnId = currentTurnId,
    owner = ws,
    epoch = connectionEpoch;
  beginTurn();
  showHealthNotice(
    persistenceWarning ? "会话持久化仍需检查，不能保证后续恢复。" : null,
    persistenceWarning || "",
  );
  setAgentStatus("working", "正在提交请求…");
  if (healthSupported) {
    sendRequest("user_message", { text, turn_id: turnId })
      .then((reply) => {
        if (ws !== owner || currentTurnId !== turnId) return;
        if (reply.connection_epoch !== epoch || reply.turn_id !== turnId) {
          healthBlocked = true;
          showHealthNotice("请求接收确认的标识不一致，正在核对状态。", "ACK_IDENTITY_MISMATCH");
          checkTurnHealth();
          return;
        }
        if (!reply.ok) {
          cancelOfficeOperations();
          endTurn();
          showHealthNotice(`请求未被接收：${reply.error || "未知原因"}`, "REQUEST_REJECTED");
          setAgentStatus("err", "请求未被接收");
        } else setAgentStatus("working", "已接收 · 等待模型回复");
      })
      .catch(() => {
        if (ws !== owner || currentTurnId !== turnId) return;
        showHealthNotice("尚未收到接收确认，正在查询实际状态；不会重复提交请求。", "ACK_TIMEOUT");
        checkTurnHealth();
      });
  } else wsSend({ type: "user_message", request_id: uuid(), text, turn_id: turnId });
  // After sending, the chip resets to "attached" for the next turn.
  attachSelection = true;
  refreshSelectionChip();
  return true;
}
$composer.addEventListener("submit", (e) => {
  e.preventDefault();
  if (sendUserMessage($input.value)) $input.value = "";
});

$input.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    $composer.dispatchEvent(new Event("submit"));
  }
});

// ---------------------------------------------------------------------------
// Theme — match Word's theme (light/dark) via Office.context.officeTheme.
// If Office doesn't expose a theme (older build), the CSS
// `prefers-color-scheme: dark` media query already handles the OS-level
// preference, so this is purely an override for when Word's theme differs
// from the OS.
// ---------------------------------------------------------------------------
function applyOfficeTheme() {
  try {
    const t = Office.context && Office.context.officeTheme;
    const bg = t && t.bodyBackgroundColor;
    if (!bg) return;
    const m = /^#?([0-9a-f]{6})$/i.exec(bg);
    if (!m) return;
    const hex = m[1];
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    document.documentElement.dataset.theme = luminance < 0.5 ? "dark" : "light";
  } catch {
    /* fall through; CSS media query handles OS-level preference. */
  }
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  // Reconcile HOST with what Office reports — body data-host should already
  // match, but Office is authoritative if they disagree.
  if (info.host === Office.HostType.Excel) HOST = "excel";
  else if (info.host === Office.HostType.Word) HOST = "word";
  else {
    setConnectionStatus("err", `Unsupported host: ${info.host}`);
    return;
  }
  document.body.dataset.host = HOST;

  applyOfficeTheme();

  try {
    activeDocUrl = Office.context.document.url || null;
  } catch {
    /* ignore */
  }
  refreshMismatchIndicator();

  // Hide host-irrelevant settings.
  if (HOST === "excel") {
    const tcRow = document.getElementById("setting-track-changes-mode")?.closest(".setting-row");
    if (tcRow) tcRow.hidden = true;
  }

  // Wire selection-change event.
  if (HOST === "excel") {
    Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(onSelectionChanged);
      await context.sync();
    }).catch((err) => console.warn("Could not attach Excel selection handler:", err));
  } else {
    Word.run(async (context) => {
      context.document.onSelectionChanged.add(onSelectionChanged);
      await context.sync();
    }).catch((err) => console.warn("Could not attach Word selection handler:", err));
  }

  // Capture once on boot.
  captureSelection();

  // Initialize presets UI.
  initPresets();

  // Show the welcome card on first launch (per browser-storage, so it
  // re-shows in a fresh profile or if the user clears storage).
  maybeShowOnboarding();

  // Fetch the bridge token before opening the WS — the bridge will close any
  // connection that doesn't present it.
  fetchBridgeToken().then(() => wsConnect());
});

// ---------------------------------------------------------------------------
// First-run onboarding card
// ---------------------------------------------------------------------------
const ONBOARDING_KEY = "claude-code-office-onboarding-seen-v1";
function maybeShowOnboarding() {
  try {
    if (localStorage.getItem(ONBOARDING_KEY)) return;
  } catch {
    /* ignore */
  }
  const card = document.createElement("div");
  card.className = "onboarding-card";
  card.innerHTML = `
    <div class="onboarding-head">
      <strong>Welcome to Draftspect for ${hostLabel()}</strong>
      <button type="button" class="onboarding-dismiss" title="Dismiss">×</button>
    </div>
    <div class="onboarding-body">
      <p>The agent reads and edits this ${hostLabel()} document, plus any folders or files you add as context.</p>
      <ol>
        <li><strong>Pick a workspace folder</strong> in the <a href="#" data-onboarding-jump="setup">Setup tab</a> — that's the folder the agent treats as its working directory.</li>
        <li><strong>Try a preset</strong> — the chips above the chat input are one-click prompts. "Summarize this document" is a good first try.</li>
        <li><strong>Add context files</strong> in Setup if you want the agent to consider background material (notes, prior drafts, references).</li>
      </ol>
    </div>
  `;
  const messagesEl = document.getElementById("messages");
  if (messagesEl) messagesEl.insertBefore(card, messagesEl.firstChild);
  else document.body.insertBefore(card, document.body.firstChild);
  card.querySelector(".onboarding-dismiss").addEventListener("click", () => {
    try {
      localStorage.setItem(ONBOARDING_KEY, "1");
    } catch {
      /* ignore */
    }
    card.remove();
  });
  card.querySelector("[data-onboarding-jump='setup']").addEventListener("click", (e) => {
    e.preventDefault();
    setActiveTab("setup");
  });
}

// ===========================================================================
// Tabs
// ===========================================================================
function setActiveTab(tabName) {
  document.body.dataset.activeTab = tabName;
  document.querySelectorAll(".tab-content").forEach((el) => {
    el.hidden = el.dataset.tab !== tabName;
  });
  document.querySelectorAll(".tab").forEach((btn) => {
    btn.setAttribute("aria-current", btn.dataset.tab === tabName ? "page" : "false");
  });
  if (tabName === "setup") {
    loadContext();
    loadWorkspaceSection();
    refreshSessionInfo().catch((error) => appendError(error.message));
  }
}

document.querySelectorAll(".tab").forEach((btn) => {
  btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
});

// ===========================================================================
// Presets — saved prompts that the user can pin to quick-chips or browse
// in the Library tab.
// ===========================================================================
// Host-scoped: localStorage is per-origin and BOTH task panes are served
// from the same origin (127.0.0.1), so a single key would let Word's
// presets show up in Excel (and vice-versa). Keying by host keeps each
// app's library separate.
const PRESETS_KEY = "claude-code-office-presets-v1:" + HOST;

function defaultPresets() {
  return HOST === "excel" ? defaultExcelPresets() : defaultWordPresets();
}

function defaultWordPresets() {
  return [
    {
      id: uuid(),
      title: "Summarize this document",
      category: "Summarize",
      prompt:
        "Read the whole document and give me a tight summary — main argument, key points, anything notable. Don't edit the document.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Outline this document",
      category: "Summarize",
      prompt:
        "Show me the heading outline of this document with paragraph counts per section. Don't edit anything.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Improve writing in selection",
      category: "Edit",
      prompt:
        "Improve the writing in my current selection — clearer, tighter, no redundancy, preserve meaning. Use track changes.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Fix typos and inconsistencies",
      category: "Edit",
      prompt:
        "Scan the whole document for typos, grammar errors, and inconsistencies (terminology, capitalization, punctuation). Use office_highlight with severity 'warning' for each issue and summarize them in chat.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Simplify the selection",
      category: "Edit",
      prompt:
        "Simplify the selected paragraph for clarity without losing meaning. Use track changes.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Add comments on this section",
      category: "Review",
      prompt:
        "Review the section my selection is in. Add Word comments on each paragraph that has a problem (unclear phrasing, weak argument, missing detail). Don't edit the text itself.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Answer using my context files",
      category: "Research",
      prompt: "Use the context files I've added to this workspace to answer: ",
      pinned: false,
      auto_send: false,
    },
    ...defaultEditingPresets(),
  ];
}

function defaultExcelPresets() {
  return [
    {
      id: uuid(),
      title: "Summarize this sheet",
      category: "Summarize",
      prompt:
        "List the worksheets, then read the active sheet's used range and give me a tight summary — what the data is, columns, row count, anything notable. Don't change anything.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Explain the selected range",
      category: "Summarize",
      prompt:
        "Read my current selection and explain what it contains — the columns, the values, and any pattern or total worth noting. Don't change anything.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Add a totals row",
      category: "Edit",
      prompt:
        "Add a labelled Total row beneath the data on the active sheet, using SUM formulas (not pre-computed numbers) for each numeric column. Read the range first; don't overwrite existing formulas.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Check the data for problems",
      category: "Review",
      prompt:
        "Scan the active sheet for data problems — blank cells in a filled column, inconsistent formatting/casing, likely typos, duplicates. List what you find in chat with cell addresses; don't change anything yet.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Find a value",
      category: "Edit",
      prompt: "Find every cell containing: ",
      pinned: false,
      auto_send: false,
    },
    {
      id: uuid(),
      title: "Answer using my context files",
      category: "Research",
      prompt: "Use the context files I've added to this workspace to answer: ",
      pinned: false,
      auto_send: false,
    },
  ];
}

// Editing presets are factored out so the migration in initPresets can
// append them to existing users' lists without duplicating the constants.
function defaultEditingPresets() {
  return [
    {
      id: uuid(),
      title: "Clear highlighting",
      category: "Editing",
      prompt:
        'Call office_clear_highlights with arguments {"all": true} to remove every highlight from the document.',
      pinned: true,
      auto_send: true,
    },
  ];
}

function uuid() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return "p_" + Math.random().toString(36).slice(2, 10);
}

let presets = [];

function loadPresets() {
  try {
    const raw = localStorage.getItem(PRESETS_KEY);
    if (raw) return JSON.parse(raw);
  } catch {
    /* ignore */
  }
  return null;
}

function savePresets() {
  try {
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
  } catch {
    /* ignore */
  }
}

function initPresets() {
  const existing = loadPresets();
  if (existing === null) {
    presets = defaultPresets();
    savePresets();
  } else {
    presets = existing;
    // Migration: append the Word-only Editing category (clear highlights)
    // for Word users seeded before it existed. Excel has no highlight
    // tool, so it never gets this.
    if (HOST === "word" && !presets.some((p) => p.category === "Editing")) {
      presets.push(...defaultEditingPresets());
      savePresets();
    }
  }
  renderLibrary();
  renderQuickChips();
}

// ---- Quick chips (pinned presets) -----------------------------------------
const $quickChips = document.getElementById("quick-chips");

function renderQuickChips() {
  const pinned = presets.filter((p) => p.pinned);
  if (pinned.length === 0) {
    $quickChips.hidden = true;
    $quickChips.innerHTML = "";
    return;
  }
  $quickChips.hidden = false;
  $quickChips.innerHTML = "";
  for (const p of pinned) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "quick-chip";
    chip.textContent = p.title;
    chip.title = p.prompt.length > 200 ? p.prompt.slice(0, 197) + "…" : p.prompt;
    chip.addEventListener("click", () => usePreset(p));
    $quickChips.appendChild(chip);
  }
}

// ---- Library tab list -----------------------------------------------------
const $libraryList = document.getElementById("library-list");

function renderLibrary() {
  $libraryList.innerHTML = "";
  if (presets.length === 0) {
    const empty = document.createElement("div");
    empty.className = "library-empty";
    empty.textContent = 'No presets yet. Click "+ New preset" to add one.';
    $libraryList.appendChild(empty);
    return;
  }

  // Group by category. Uncategorized go under "Other".
  const groups = new Map();
  for (const p of presets) {
    const k = p.category && p.category.trim() ? p.category.trim() : "Other";
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(p);
  }

  for (const [category, items] of groups) {
    const heading = document.createElement("div");
    heading.className = "library-category";
    heading.textContent = category;
    $libraryList.appendChild(heading);
    for (const p of items) {
      $libraryList.appendChild(renderPresetRow(p));
    }
  }
}

function renderPresetRow(p) {
  const row = document.createElement("div");
  row.className = "preset-row";
  row.title = p.prompt.length > 300 ? p.prompt.slice(0, 297) + "…" : p.prompt;

  // Click row → use preset
  row.addEventListener("click", (e) => {
    if (e.target.closest(".preset-actions") || e.target.closest(".preset-pin")) return;
    usePreset(p);
  });

  const title = document.createElement("div");
  title.className = "preset-title";
  title.textContent = p.title;
  row.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "preset-meta";
  if (p.auto_send) meta.textContent = "auto-send";
  row.appendChild(meta);

  const pin = document.createElement("button");
  pin.type = "button";
  pin.className = "preset-pin" + (p.pinned ? " pinned" : "");
  pin.textContent = p.pinned ? "📌" : "📍";
  pin.title = p.pinned ? "Pinned (click to unpin)" : "Pin to quick chips";
  pin.addEventListener("click", (e) => {
    e.stopPropagation();
    p.pinned = !p.pinned;
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  row.appendChild(pin);

  const actions = document.createElement("div");
  actions.className = "preset-actions";
  const editBtn = document.createElement("button");
  editBtn.className = "icon-btn";
  editBtn.type = "button";
  editBtn.title = "Edit";
  editBtn.textContent = "✎";
  editBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    openPresetModal(p);
  });
  actions.appendChild(editBtn);

  const delBtn = document.createElement("button");
  delBtn.className = "icon-btn";
  delBtn.type = "button";
  delBtn.title = "Delete";
  delBtn.textContent = "✕";
  delBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    // No confirm dialog (Office.js taskpanes block it on some builds); the
    // user can re-add a deleted preset if needed.
    presets = presets.filter((x) => x.id !== p.id);
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  actions.appendChild(delBtn);

  row.appendChild(actions);
  return row;
}

// ---- Use preset (click handler) -------------------------------------------
function usePreset(p) {
  if (p.auto_send) {
    sendUserMessage(p.prompt);
  } else {
    $input.value = p.prompt;
    $input.focus();
    // Cursor at end so user can extend the prompt
    $input.setSelectionRange($input.value.length, $input.value.length);
  }
  setActiveTab("chat");
}

// ---- Preset editor modal --------------------------------------------------
const $presetModal = document.getElementById("preset-modal");
const $presetModalTitle = document.getElementById("preset-modal-title");
const $presetTitle = document.getElementById("preset-title");
const $presetPrompt = document.getElementById("preset-prompt");
const $presetCategory = document.getElementById("preset-category");
const $presetAutoSend = document.getElementById("preset-auto-send");
const $presetPinned = document.getElementById("preset-pinned");
const $presetSave = document.getElementById("preset-save");
const $presetCancel = document.getElementById("preset-cancel");
const $presetModalClose = document.getElementById("preset-modal-close");

let editingPresetId = null;

function openPresetModal(p) {
  if (p) {
    editingPresetId = p.id;
    $presetModalTitle.textContent = "Edit preset";
    $presetTitle.value = p.title;
    $presetPrompt.value = p.prompt;
    $presetCategory.value = p.category || "";
    $presetAutoSend.checked = !!p.auto_send;
    $presetPinned.checked = !!p.pinned;
  } else {
    editingPresetId = null;
    $presetModalTitle.textContent = "New preset";
    $presetTitle.value = "";
    $presetPrompt.value = "";
    $presetCategory.value = "";
    $presetAutoSend.checked = false;
    $presetPinned.checked = false;
  }
  $presetModal.hidden = false;
  setTimeout(() => $presetTitle.focus(), 0);
}

function closePresetModal() {
  $presetModal.hidden = true;
  editingPresetId = null;
}

$presetCancel.addEventListener("click", closePresetModal);
$presetModalClose.addEventListener("click", closePresetModal);
$presetModal.addEventListener("click", (e) => {
  if (e.target === $presetModal) closePresetModal();
});

$presetSave.addEventListener("click", () => {
  const title = $presetTitle.value.trim();
  const prompt = $presetPrompt.value;
  if (!title) {
    $presetTitle.focus();
    return;
  }
  if (!prompt.trim()) {
    $presetPrompt.focus();
    return;
  }

  const data = {
    title,
    prompt,
    category: $presetCategory.value.trim(),
    auto_send: $presetAutoSend.checked,
    pinned: $presetPinned.checked,
  };

  if (editingPresetId) {
    const idx = presets.findIndex((p) => p.id === editingPresetId);
    if (idx !== -1) presets[idx] = { ...presets[idx], ...data };
  } else {
    presets.push({ id: uuid(), ...data });
  }

  savePresets();
  renderLibrary();
  renderQuickChips();
  closePresetModal();
});

document.getElementById("add-preset").addEventListener("click", () => openPresetModal(null));

// ===========================================================================
// Setup tab — Context files (folders or individual files saved to the
// workspace's CLAUDE.md; file tools require explicit extended mode).
// ===========================================================================
let contextCache = null;
let contextLoadingPromise = null;
const $contextList = document.getElementById("context-list");

async function removeContextEntryAt(idx) {
  if (!contextCache) return;
  contextCache = contextCache.filter((_, i) => i !== idx);
  await saveContext();
}

async function loadContext(force = false) {
  if (contextCache && !force) {
    renderContext();
    return;
  }
  if (contextLoadingPromise) return contextLoadingPromise;
  contextLoadingPromise = (async () => {
    try {
      $contextList.innerHTML = '<div class="references-loading">Loading…</div>';
      const r = await sendRequest("get_context");
      contextCache = Array.isArray(r.entries) ? r.entries : [];
      renderContext();
    } catch (e) {
      showListMessage($contextList, "references-empty", `Could not load: ${e.message}`);
    } finally {
      contextLoadingPromise = null;
    }
  })();
  return contextLoadingPromise;
}

function renderContext() {
  $contextList.innerHTML = "";
  if (!contextCache || contextCache.length === 0) {
    const empty = document.createElement("div");
    empty.className = "references-empty";
    empty.textContent = "No context files yet — add a folder or file to give the agent background.";
    $contextList.appendChild(empty);
    return;
  }
  contextCache.forEach((e, i) => {
    const row = document.createElement("div");
    row.className = "reference-row";
    row.title = e.path;

    const info = document.createElement("div");
    info.className = "reference-info";
    const pathEl = document.createElement("div");
    pathEl.className = "reference-path";
    const tag = document.createElement("span");
    tag.className = "kind-tag";
    tag.textContent = e.kind || "?";
    pathEl.appendChild(tag);
    pathEl.appendChild(document.createTextNode(e.path));
    info.appendChild(pathEl);
    if (e.description) {
      const desc = document.createElement("div");
      desc.className = "reference-description";
      desc.textContent = e.description;
      info.appendChild(desc);
    }
    row.appendChild(info);

    const remove = document.createElement("button");
    remove.className = "reference-remove";
    remove.type = "button";
    remove.title = "Remove";
    remove.textContent = "✕";
    remove.addEventListener("click", () => removeContextEntryAt(i));
    row.appendChild(remove);

    $contextList.appendChild(row);
  });
}

// In-pane error notice for the Context-files section. window.alert() works
// on Mac Office but is flaky/blocked in web Office; show a dismissible
// inline message in the Setup tab instead. textContent — never innerHTML —
// so a daemon-supplied path/error string can't inject markup.
const $contextError = document.getElementById("context-error");
function showContextError(message) {
  if (!$contextError) {
    console.warn("[context]", message);
    return;
  }
  $contextError.textContent = message;
  $contextError.hidden = false;
}
function clearContextError() {
  if ($contextError) {
    $contextError.textContent = "";
    $contextError.hidden = true;
  }
}

async function saveContext() {
  if (!contextCache) return false;
  clearContextError();
  try {
    const r = await sendRequest("set_context", { entries: contextCache });
    if (r.errors && r.errors.length > 0) {
      const lines = r.errors.map((e) => `${e.path} — ${e.error}`).join("; ");
      showContextError(`Some entries could not be saved: ${lines}`);
    }
    contextCache = Array.isArray(r.saved) ? r.saved : contextCache;
    renderContext();
    return true;
  } catch (e) {
    showContextError(`Could not save: ${e.message}`);
    return false;
  }
}

// ---- Add-folder modal (shared between guidelines + samples) ----------------
const $addFolderModal = document.getElementById("add-folder-modal");
const $addFolderModalTitle = document.getElementById("add-folder-modal-title");
const $addFolderPath = document.getElementById("add-folder-path");
const $addFolderDescription = document.getElementById("add-folder-description");
const $addFolderError = document.getElementById("add-folder-error");
const $addFolderSave = document.getElementById("add-folder-save");
const $addFolderCancel = document.getElementById("add-folder-cancel");
const $addFolderModalClose = document.getElementById("add-folder-modal-close");
const $addFolderBrowseFile = document.getElementById("add-folder-browse-file");
const $addFolderBrowseFolder = document.getElementById("add-folder-browse-folder");

function openAddFolderModal(prefillPath = "", kind = null) {
  $addFolderModalTitle.textContent =
    kind === "file" ? "Add file" : kind === "folder" ? "Add folder" : "Add folder or file";
  $addFolderPath.value = prefillPath;
  $addFolderDescription.value = "";
  $addFolderError.hidden = true;
  $addFolderModal.hidden = false;
  // If the path is already chosen (the common flow — user picked first),
  // jump straight to the description field.
  setTimeout(() => (prefillPath ? $addFolderDescription : $addFolderPath).focus(), 0);
}
function closeAddFolderModal() {
  $addFolderModal.hidden = true;
}

// Two explicit entry points: pick first (single-mode dialog, reliable on
// every OS), then the modal just collects an optional description.
async function addContextEntry(includeFiles) {
  let picked;
  try {
    picked = await pickPathNative({
      start_path: currentWorkspaceCwd || null,
      include_files: includeFiles,
    });
  } catch (e) {
    console.error("[picker]", e);
    return;
  }
  if (!picked) return; // canceled — don't open an empty modal
  openAddFolderModal(picked.path, includeFiles ? "file" : "folder");
}
document
  .querySelectorAll(".add-context-folder")
  .forEach((b) => b.addEventListener("click", () => addContextEntry(false)));
document
  .querySelectorAll(".add-context-file")
  .forEach((b) => b.addEventListener("click", () => addContextEntry(true)));
$addFolderCancel.addEventListener("click", closeAddFolderModal);
$addFolderModalClose.addEventListener("click", closeAddFolderModal);
$addFolderModal.addEventListener("click", (e) => {
  if (e.target === $addFolderModal) closeAddFolderModal();
});

$addFolderSave.addEventListener("click", async () => {
  const path = $addFolderPath.value.trim();
  const description = $addFolderDescription.value.trim();
  $addFolderError.hidden = true;
  if (!path) {
    $addFolderError.textContent = "Path is required.";
    $addFolderError.hidden = false;
    return;
  }
  if (!contextCache) await loadContext();
  if (!contextCache) contextCache = [];
  contextCache = [...contextCache, { path, description }];
  const ok = await saveContext();
  if (ok) closeAddFolderModal();
});

// ---- Native folder/file picker ---------------------------------------------
// Forwards the pick request through the daemon to the Electron main process,
// which shows a real macOS NSOpenPanel. This is the same panel every native
// Mac app uses, so it can navigate to Google Drive, iCloud, "Shared with me",
// recent items, sidebar shortcuts — none of which a synthetic in-page browser
// can reach. Resolves to `{ path, kind }` or `null` if the user cancelled.
async function pickPathNative({ start_path = null, include_files = false, title = null } = {}) {
  const r = await sendRequest("pick_path", {
    default_path: start_path,
    include_files,
    title,
  });
  if (!r.ok) throw new Error(r.error || "Picker failed");
  if (r.canceled) return null;
  return { path: r.path, kind: r.kind };
}

// Two single-mode pickers. A combined file+folder native dialog can't
// exist on Windows (it degrades to folder-only, hiding files), so the
// user explicitly chooses which kind to browse for.
async function browseInto(includeFiles) {
  const startPath = $addFolderPath.value.trim() || currentWorkspaceCwd || null;
  try {
    const picked = await pickPathNative({ start_path: startPath, include_files: includeFiles });
    if (!picked) return;
    $addFolderPath.value = picked.path;
    $addFolderDescription.focus();
  } catch (e) {
    console.error("[picker]", e);
  }
}
$addFolderBrowseFile.addEventListener("click", () => browseInto(true));
$addFolderBrowseFolder.addEventListener("click", () => browseInto(false));

$addFolderPath.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    $addFolderSave.click();
  }
});
$addFolderDescription.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    $addFolderSave.click();
  }
});

// ===========================================================================
// Workspace selection — managed in the Setup tab's Workspace section.
// The topbar chip is read-only status (current workspace name + mismatch warning);
// clicking it navigates to the Setup tab where the actual switching UI lives.
// ===========================================================================
let currentWorkspaceCwd = null;

// Replace a list container with a single-line empty-state message. Uses
// textContent so an attacker-controlled error string can't inject markup.
function showListMessage(container, klass, message) {
  container.innerHTML = "";
  const div = document.createElement("div");
  div.className = klass;
  div.textContent = message;
  container.appendChild(div);
}

const $workspaceChip = document.getElementById("workspace-chip");
const $workspaceFolder = document.getElementById("workspace-folder");
const $workspaceCwd = document.getElementById("workspace-cwd");
const $addWorkspace = document.getElementById("add-workspace");
const $workspaceError = document.getElementById("workspace-error");
const $workspaceWarning = document.getElementById("workspace-warning");

// The workspace is simply the folder the open document lives in. We follow
// it automatically. An explicit "Change workspace" pick overrides that and
// sticks until the active document moves to a *different* folder — tracked
// via lastFollowedDocDir so a re-render/reconnect doesn't yank the user's
// deliberate choice back.
let lastFollowedDocDir = null;

function setWorkspaceDisplay(cwd) {
  currentWorkspaceCwd = cwd;
  const name = cwd ? cwd.split(/[\\/]/).filter(Boolean).pop() : "(no workspace)";
  $workspaceFolder.textContent = name;
  $workspaceFolder.title = cwd || "";
  if ($workspaceCwd) {
    $workspaceCwd.textContent = cwd || "(no workspace)";
    $workspaceCwd.title = cwd || "";
  }
  refreshMismatchIndicator();
}

// Decide whether the workspace chip should show a warning. True when an active
// doc exists and its filesystem location is NOT inside the current workspace
// folder — the agent's filesystem tools won't see the doc's siblings.
function refreshMismatchIndicator() {
  let mismatch = false;
  if (activeDocUrl && currentWorkspaceCwd) {
    const docDir = docDirFromActiveUrl(activeDocUrl);
    // Mismatch if docDir isn't underneath the workspace cwd. Cloud-hosted
    // docs (docDir === "") fall through as not-a-mismatch — there's no
    // filesystem location to compare, so the chip stays neutral.
    mismatch = !!docDir && !isInOrUnder(docDir, currentWorkspaceCwd);
  }
  const baseTitle =
    "Workspace controls session context. Reading source files requires explicit extended mode; Office-only is the default.";
  if (mismatch) {
    $workspaceChip.classList.add("mismatch");
    $workspaceWarning.hidden = false;
    $workspaceChip.title =
      baseTitle + " (⚠ The current workspace doesn't match the doc you're editing.)";
  } else {
    $workspaceChip.classList.remove("mismatch");
    $workspaceWarning.hidden = true;
    $workspaceChip.title = baseTitle;
  }
}

async function refreshWorkspaceFromDaemon() {
  try {
    const r = await sendRequest("get_cwd_state");
    if (r.current_cwd) setWorkspaceDisplay(r.current_cwd);
    await autoFollowDocWorkspace();
  } catch {
    /* ignore on initial boot */
  }
}

async function loadWorkspaceSection() {
  $workspaceError.hidden = true;
  await autoFollowDocWorkspace();
  try {
    const r = await sendRequest("get_cwd_state");
    if (r.current_cwd) setWorkspaceDisplay(r.current_cwd);
  } catch (e) {
    $workspaceError.textContent = `Could not load workspace: ${e.message}`;
    $workspaceError.hidden = false;
  }
}

// The workspace follows the open document's folder automatically. Workspace
// detection is deterministic now (the doc's own folder), so there's nothing
// to confirm — no banner, no setting. We switch when the doc's folder isn't
// the current workspace, UNLESS the user made an explicit pick for this same
// doc-folder (lastFollowedDocDir), in which case their choice stands until
// they open a document in a different folder.
async function autoFollowDocWorkspace() {
  if (!activeDocUrl) return;
  const docDir = docDirFromActiveUrl(activeDocUrl);
  if (!docDir) return; // cloud doc (no filesystem path) — leave workspace as-is
  if (docDir === lastFollowedDocDir) return; // already handled (incl. explicit override)
  if (currentWorkspaceCwd && isInOrUnder(docDir, currentWorkspaceCwd)) {
    lastFollowedDocDir = docDir; // doc already inside the workspace — fine
    return;
  }
  lastFollowedDocDir = docDir;
  await doSwitch(null, { autodetectFromDoc: true });
}

async function doSwitch(cwd, { autodetectFromDoc = false } = {}) {
  $workspaceError.hidden = true;
  try {
    const payload = autodetectFromDoc ? { autodetect_from_doc: activeDocUrl } : { cwd };
    const r = await sendRequest("set_cwd", payload);
    if (!r.ok) throw new Error(r.error || "switch failed");
    // An explicit pick (not the auto-follow) is a deliberate override: pin
    // it to the current doc's folder so autoFollowDocWorkspace won't yank
    // it back until the user opens a document in a different folder.
    if (!autodetectFromDoc) lastFollowedDocDir = docDirFromActiveUrl(activeDocUrl) || null;
    // The daemon emits cwd_changed which updates the chip via assistant_event.
    // Clear chat — visually distinguishing the new session from the old.
    // The daemon's scoped transcript_replay is authoritative; do not clear it here.
    // Refresh the workspace section so the displayed cwd updates.
    loadWorkspaceSection();
  } catch (e) {
    $workspaceError.textContent = e.message;
    $workspaceError.hidden = false;
  }
}

// Clicking the topbar chip jumps to the Setup tab where the workspace UI lives.
$workspaceChip.addEventListener("click", () => setActiveTab("setup"));

// "Change workspace" — opens the native folder picker; on pick, switch to
// that folder (a deliberate override of the auto-followed doc folder).
$addWorkspace.addEventListener("click", async () => {
  try {
    const picked = await pickPathNative({ title: "Choose a workspace folder" });
    if (picked) doSwitch(picked.path);
  } catch (e) {
    console.error("[picker]", e);
  }
});
