// Opt-in Electron/Chromium DOM acceptance; excluded from npm test's *.test.mjs.
// Run: <installed Electron executable> tests/pane-disconnect-electron-e2e.mjs <absolute evidence directory>
// Office and transport are controlled mocks. No daemon, SDK, model or remote
// service is loaded. This proves native Chromium presentation, not Office.js.
import { app, BrowserWindow } from "electron";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { mkdtempSync } from "node:fs";
import { tmpdir, release } from "node:os";
import { join, isAbsolute } from "node:path";
import { pathToFileURL } from "node:url";
import { createHash, randomUUID } from "node:crypto";

const output = process.argv[2];
if (!output || !isAbsolute(output)) throw new Error("Supply an absolute evidence directory");
const temporary = mkdtempSync(join(tmpdir(), "draftspect-pane-disconnect-"));
app.setPath("userData", join(temporary, "user-data"));
app.setPath("crashDumps", join(temporary, "crash-dumps"));
const startedAt = new Date().toISOString();
const network = [],
  processFaults = [],
  checks = [],
  windows = new Set();
const ownedProcesses = new Map([[process.pid, { pid: process.pid, type: "Browser" }]]);
const hash = (value) => createHash("sha256").update(value).digest("hex");
const preferences = {
  contextIsolation: true,
  sandbox: true,
  webSecurity: true,
  nodeIntegration: false,
};
// There are separate isolated host windows. Keep this test app alive between
// their destruction; the bounded run's finally block is its sole quit owner.
app.on("window-all-closed", () => {});
const identity = {
  started_at: startedAt,
  versions: {
    electron: process.versions.electron,
    chromium: process.versions.chrome,
    node: process.versions.node,
  },
  platform: process.platform,
  architecture: process.arch,
  os_release: release(),
  web_preferences: preferences,
  hardware_acceleration: "DEFAULT_UNCHANGED",
  native_office: "NOT_RUN",
  office_operations: "CONTROLLED_MOCK_ONLY",
  rpc_timeouts: "CONTROLLED_TIMER_INJECTION_IN_NATIVE_RENDERER",
  capture_pipeline: "TWO_RENDERER_ANIMATION_FRAMES_BEFORE_SNAPSHOT",
  model_requests: 0,
  crash_reporter_uploads: 0,
};
app.on("child-process-gone", (_event, details) => {
  if (!["clean-exit", "killed"].includes(details.reason))
    processFaults.push({ type: details.type, reason: details.reason, exit_code: details.exitCode });
});

async function nativePane(host, source, queue, css) {
  const htmlSource = await readFile(
    new URL(`../taskpane/${host}/index.html`, import.meta.url),
    "utf8",
  );
  const html = htmlSource
    .replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, "")
    .replace('<link rel="stylesheet" href="/shared/styles.css" />', `<style>${css}</style>`);
  const fixture = join(temporary, `${host}.html`);
  await writeFile(fixture, html, { flag: "wx", mode: 0o600 });
  const window = new BrowserWindow({
    show: false,
    width: 430,
    height: 820,
    webPreferences: {
      ...preferences,
      partition: `native-pane-disconnect-${randomUUID()}`,
      // Keep only this isolated hidden renderer painting for current snapshots.
      backgroundThrottling: false,
    },
  });
  windows.add(window);
  const effectivePreferences = window.webContents.getLastWebPreferences();
  if (Object.entries(preferences).some(([key, value]) => effectivePreferences[key] !== value))
    throw new Error("Native renderer security preferences differ from the required test profile");
  if (window.webContents.getBackgroundThrottling())
    throw new Error("Native capture test background throttling must be disabled");
  window.webContents.setWindowOpenHandler(() => ({ action: "deny" }));
  window.webContents.on("will-navigate", (event, url) => {
    if (url !== pathToFileURL(fixture).href) event.preventDefault();
  });
  window.webContents.on("render-process-gone", (_event, details) => {
    processFaults.push({
      host,
      type: "renderer",
      reason: details.reason,
      exit_code: details.exitCode,
    });
  });
  window.webContents.session.webRequest.onBeforeRequest(
    { urls: ["http://*/*", "https://*/*", "ws://*/*", "wss://*/*"] },
    (_details, done) => {
      network.push({ host, blocked: true });
      done({ cancel: true });
    },
  );
  await window.loadFile(fixture);
  await window.webContents.executeJavaScript(`
    globalThis.nativePaneHarness = { sent: [], sockets: [], timers: new Map(), timerId: 0,
      checks: [], writes: 0, fetches: [], releaseCommit: null };
    globalThis.Office = { context: { document: { url: ${JSON.stringify(`synthetic.${host === "word" ? "docx" : "xlsx"}`)} } }, onReady() {} };
    localStorage.setItem("draftspect-preview-settings-v1", JSON.stringify({ model: "synthetic-model-a", modelsByRuntime: { codebuddy: "synthetic-model-a" } }));
    globalThis.fetch = async url => {
      nativePaneHarness.fetches.push(String(url));
      if (url !== "/bridge-token") throw new Error("Unexpected mock fetch");
      return { ok: true, text: async () => "synthetic-token" };
    };
    globalThis.setTimeout = (fn, ms) => {
      const id = ++nativePaneHarness.timerId;
      nativePaneHarness.timers.set(id, { fn, ms });
      return id;
    };
    globalThis.clearTimeout = id => nativePaneHarness.timers.delete(id);
    globalThis.WebSocket = function () {
      const socket = { send: raw => nativePaneHarness.sent.push(JSON.parse(raw)), close: () => socket.onclose?.() };
      nativePaneHarness.sockets.push(socket);
      return socket;
    };
    ${queue}
    const prepareReviewWrite = async () => ({
      preview: { action: "Synthetic replacement" },
      apply: async ({ beforeWrite }) => {
        beforeWrite(); nativePaneHarness.writes++;
        await new Promise(resolve => { nativePaneHarness.releaseCommit = resolve; });
        return { verified: true };
      }
    });
    const resolveWordHeadingArgs = async (_name, args) => args;
    ${source}
    nativePaneHarness.check = (name, ok) => {
      nativePaneHarness.checks.push({ host: ${JSON.stringify(host)}, name, status: ok ? "PASS" : "FAIL" });
      if (!ok) throw new Error(name);
    };
    nativePaneHarness.visible = el => !!el && el.checkVisibility({ checkVisibilityCSS: true });
    nativePaneHarness.flush = async () => { for (let i = 0; i < 20; i++) await Promise.resolve(); };
    nativePaneHarness.expire = async type => {
      const request = [...pendingRequests.values()].find(p => p.type === type);
      nativePaneHarness.check("controlled RPC timeout has an owned pending request", !!request);
      const timer = nativePaneHarness.timers.get(request.timer);
      nativePaneHarness.timers.delete(request.timer); timer.fn();
      await nativePaneHarness.flush();
    };
    nativePaneHarness.reconnect = async () => {
      const [id, timer] = [...nativePaneHarness.timers].find(([, timer]) => timer.ms === 1500);
      nativePaneHarness.timers.delete(id); timer.fn();
      await nativePaneHarness.flush();
      nativePaneHarness.sockets.at(-1).onopen();
    };
    nativePaneHarness.welcome = async epoch => {
      await handleServerMessage({ type: "welcome", protocol_version: 2, runtime: "codebuddy",
        connection_epoch: epoch, capabilities: ["turn_health_v1", "confirm_window_v1"] });
      const catalog = nativePaneHarness.sent.findLast(m => m.type === "list_models");
      await handleServerMessage({ type: "list_models_result", request_id: catalog.request_id,
        ok: true, runtime: "codebuddy", models: [{ value: "synthetic-model-a" }] });
      await nativePaneHarness.flush();
    };
    nativePaneHarness.model = async () => {
      const model = nativePaneHarness.sent.findLast(m => m.type === "set_model");
      await handleServerMessage({ type: "set_model_result", request_id: model.request_id, ok: true, model: model.model });
      await nativePaneHarness.flush();
    };
    nativePaneHarness.history = () => handleServerMessage({ type: "transcript_replay", events: [] });
    nativePaneHarness.api = { wsConnect, sendUserMessage, handleServerMessage, runOfficeTool,
      checkTurnHealth, appendToolUse, updateToolResult, card: id => toolCards.get(id),
      state: () => ({ inFlight: turnInFlight, turn: currentTurnId, healthBlocked, wsReady }),
      turn: id => { currentTurnId = id; beginTurn(); } };
    void 0;
  `);
  const evaluate = (code) => window.webContents.executeJavaScript(`(async () => { ${code} })()`);
  const screenshot = async (name) => {
    // DOM checks can finish before Chromium composites a new frame. Keep the
    // real frame scheduler (only RPC timers are mocked) and capture afterwards.
    await window.webContents.executeJavaScript(
      "new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))",
    );
    await writeFile(
      join(output, `${host}-${name}.png`),
      (await window.webContents.capturePage(undefined, { stayHidden: true })).toPNG(),
      { flag: "wx", mode: 0o600 },
    );
  };
  try {
    await evaluate(`
      const h = nativePaneHarness;
      h.api.wsConnect(); h.sockets[0].onopen();
      await h.welcome("epoch-A"); await h.history(); await h.model();
      h.check("normal readiness is visible after all admission acknowledgements", /可发送/.test(document.getElementById("agent-status").textContent) && !document.getElementById("send").disabled);
      h.check("production fault detail starts closed", !document.querySelector("#health-notice details").open);
      h.check("renderer has no Node integration", typeof globalThis.require === "undefined" && typeof globalThis.process === "undefined");
    `);
    await screenshot("ready");
    await evaluate(`
      const h = nativePaneHarness;
      h.check("only an explicit user submission creates a request", h.api.sendUserMessage("Synthetic summary; do not edit"));
      const request = h.sent.findLast(m => m.type === "user_message");
      await h.api.handleServerMessage({ type: "user_message_result", request_id: request.request_id, turn_id: request.turn_id, connection_epoch: "epoch-A", ok: true });
      await h.flush();
      h.api.appendToolUse("read complete", {}, "completed"); h.api.updateToolResult("completed", "Synthetic verified read", "completed");
      h.api.appendToolUse("read pending", {}, "pending");
      h.sockets[0].onclose(); await h.flush();
      const card = h.api.card("pending");
      h.check("disconnect fault card visible in native DOM", h.visible(document.getElementById("health-notice")) && /^CONNECTION_UNAVAILABLE/.test(document.getElementById("health-notice-detail").textContent));
      h.check("loss of connection is not displayed as a confirmed crash", /原因尚未确认/.test(document.getElementById("health-notice-text").textContent) && !/已崩溃|确认崩溃/.test(document.getElementById("health-notice-text").textContent));
      h.check("pending record becomes visibly unconfirmed outside closed details", card.state === "interrupted / outcome unknown" && !card.group.details.open && h.visible(card.group.alert) && !h.visible(card.el));
      h.check("confirmed prior result is preserved", h.api.card("completed").state === "completed");
      h.check("disconnect has no permanent Working or false readiness", !h.api.state().inFlight && document.getElementById("send").disabled && !/可发送|Working/.test(document.getElementById("agent-status").textContent));
      h.check("disconnected footer uses the error state rather than a green idle indicator", document.getElementById("agent-status").className === "agent-status err");
      h.check("fault diagnostic details remain collapsed", !document.querySelector("#health-notice details").open);
      h.check("existing recovery control is visible", h.visible(document.getElementById("health-reconnect")));
    `);
    await screenshot("disconnected");
    await evaluate(`
      const h = nativePaneHarness;
      await h.reconnect(); await h.welcome("epoch-B");
      h.check("fresh transport welcome alone cannot enable sending", document.getElementById("send").disabled);
      await h.model();
      h.check("model acknowledgement still waits for history", document.getElementById("send").disabled);
      await h.history();
      h.check("all fresh acknowledgements recover the composer", !document.getElementById("send").disabled && /可发送/.test(document.getElementById("agent-status").textContent));
      h.check("recovery guidance explicitly requires a manual new request and reread", /^CONNECTION_RECOVERED/.test(document.getElementById("health-notice-detail").textContent) && /未被自动重发/.test(document.getElementById("health-notice-text").textContent) && /手动/.test(document.getElementById("health-notice-text").textContent));
      h.check("reconnect never replays a model request or an Office mutation", h.sent.filter(m => m.type === "user_message").length === 1 && h.writes === 0 && !h.sent.some(m => m.type === "tool_result"));
    `);
    await screenshot("recovered");
    await evaluate(`
      const h = nativePaneHarness;
      h.api.turn("unresponsive-synthetic");
      for (let i = 0; i < 2; i++) { const pending = h.api.checkTurnHealth(); await h.expire("get_turn_state"); await pending; }
      h.check("unresponsive service displays error while keeping its unknown request locked", document.getElementById("agent-status").className === "agent-status err" && h.api.state().inFlight && document.getElementById("send").disabled);
      h.check("unresponsive Stop and recovery controls remain visible", h.visible(document.getElementById("stop-agent")) && h.visible(document.getElementById("health-reconnect")));
      h.check("unresponsive service classification is not confirmed process death", /^SERVICE_UNRESPONSIVE/.test(document.getElementById("health-notice-detail").textContent) && !/已崩溃|确认崩溃/.test(document.getElementById("health-notice-text").textContent));
      h.check("timeout detection does not create extra user or write requests", h.sent.filter(m => m.type === "user_message").length === 1 && h.writes === 0);
    `);
    await screenshot("unresponsive");
    if (host === "word") {
      await evaluate(`
        const h = nativePaneHarness;
        h.sockets.at(-1).onclose(); await h.reconnect();
        await h.welcome("epoch-C"); await h.model(); await h.history();
        h.api.turn("in-mock-write");
        const message = { id: "native-mock-write", name: "office_edit_range", args: { target: { query: "Synthetic" }, action: "replace", text: "Changed" }, turn_id: "in-mock-write", deadline: Date.now() + 120000 };
        const pending = h.api.runOfficeTool(message); await h.flush();
        const card = h.api.card(message.id);
        h.check("a write requires a real DOM approval button", /awaiting your confirmation/.test(card.state) && h.visible(card.el.querySelector("button")));
        card.el.querySelector("button").click(); await h.flush();
        h.check("approved controlled Office commit is in flight exactly once", h.writes === 1 && typeof h.releaseCommit === "function");
        h.sockets.at(-1).onclose(); await h.flush();
        h.check("disconnect during a controlled commit immediately persists unknown outcome", sessionStorage.getItem("draftspect-inspect-before-write:synthetic.docx") === "unknown" && /修改结果待核实/.test(document.getElementById("agent-status").textContent));
        h.releaseCommit(); await pending;
        h.check("late mock commit return is not promoted to success", card.state === "outcome unknown — inspect Word" && !h.sent.some(m => m.type === "tool_result" && m.id === message.id && m.ok));
        await h.reconnect(); await h.welcome("epoch-D"); await h.model(); await h.history();
        h.check("transport recovery preserves the uncertain document warning", /修改结果待核实/.test(document.getElementById("agent-status").textContent) && sessionStorage.getItem("draftspect-inspect-before-write:synthetic.docx") === "unknown");
        h.api.turn("new-mock-proposal");
        await h.api.runOfficeTool({ ...message, turn_id: "new-mock-proposal" });
        await h.api.runOfficeTool({ ...message, id: "different-id", turn_id: "new-mock-proposal" });
        h.check("neither original nor fresh ID can replay an uninspected mutation", h.writes === 1 && sessionStorage.getItem("draftspect-operation:native-mock-write") === "attempted");
        h.check("reread and inspection acknowledgement remains explicit", [...document.querySelectorAll(".msg.error")].some(el => /uncertain outcome/.test(el.textContent) && /reread/.test(el.textContent)));
      `);
      await screenshot("unknown-write");
    }
    const result = await evaluate(`return { checks: nativePaneHarness.checks,
      model_request_envelopes: nativePaneHarness.sent.filter(m => m.type === "user_message").length,
      controlled_writes: nativePaneHarness.writes, fetches: nativePaneHarness.fetches };`);
    result.effective_web_preferences = Object.fromEntries(
      Object.keys(preferences).map((key) => [key, effectivePreferences[key]]),
    );
    result.background_throttling = window.webContents.getBackgroundThrottling();
    for (const metric of app.getAppMetrics())
      ownedProcesses.set(metric.pid, { pid: metric.pid, type: metric.type });
    checks.push(...result.checks);
    return result;
  } catch (error) {
    const partial = await evaluate("return nativePaneHarness.checks;").catch(() => []);
    checks.push(...partial);
    throw error;
  } finally {
    window.destroy();
    windows.delete(window);
  }
}

async function run() {
  await mkdir(output, { recursive: true });
  let sourceHash = null;
  const timeout = setTimeout(() => {
    for (const window of windows) if (!window.isDestroyed()) window.destroy();
    void writeFile(
      join(output, "timeout.json"),
      JSON.stringify(
        { ...identity, status: "FAIL", reason: "BOUNDED_NATIVE_TEST_TIMEOUT" },
        null,
        2,
      ),
      { flag: "wx", mode: 0o600 },
    ).finally(() => app.exit(90));
  }, 45000);
  try {
    await app.whenReady();
    const original = await readFile(
      new URL("../taskpane/shared/taskpane.js", import.meta.url),
      "utf8",
    );
    sourceHash = hash(original);
    const source = original.replace(/^import\s+[\s\S]*?;\s*$/gm, "");
    const queue = (
      await readFile(new URL("../taskpane/shared/operation-queue.js", import.meta.url), "utf8")
    ).replace("export function", "function");
    const css = await readFile(new URL("../taskpane/shared/styles.css", import.meta.url), "utf8");
    const hosts = {};
    for (const host of ["word", "excel"]) hosts[host] = await nativePane(host, source, queue, css);
    if (network.length || processFaults.length)
      throw new Error("Unexpected blocked network request or native process fault");
    const report = {
      ...identity,
      ended_at: new Date().toISOString(),
      status: "PASS",
      source_sha256: sourceHash,
      test_sha256: hash(await readFile(new URL(import.meta.url))),
      checks,
      check_count: checks.length,
      network_attempts: network.length,
      process_faults: processFaults,
      owned_processes: [...ownedProcesses.values()],
      hosts,
      temporary_profile: temporary,
      own_windows_remaining: windows.size,
    };
    await writeFile(join(output, "native-report.json"), JSON.stringify(report, null, 2) + "\n", {
      flag: "wx",
      mode: 0o600,
    });
    console.log(
      JSON.stringify({
        status: report.status,
        checks: checks.length,
        source_sha256: sourceHash,
        network_attempts: 0,
        model_requests: 0,
      }),
    );
  } catch (error) {
    await writeFile(
      join(output, "native-failure.json"),
      JSON.stringify(
        {
          ...identity,
          ended_at: new Date().toISOString(),
          status: "FAIL",
          source_sha256: sourceHash,
          checks,
          network_attempts: network.length,
          process_faults: processFaults,
          error: String(error.stack || error),
        },
        null,
        2,
      ) + "\n",
      { flag: "wx", mode: 0o600 },
    );
    console.error(error);
    process.exitCode = 1;
  } finally {
    clearTimeout(timeout);
    for (const window of windows) if (!window.isDestroyed()) window.destroy();
    app.exit(process.exitCode || 0);
  }
}
void run();
