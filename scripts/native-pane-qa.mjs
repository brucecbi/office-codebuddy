// Actual production pane + model-free bridge. No daemon/runtime/user-config imports.
import { readFile, writeFile, mkdir, stat } from "node:fs/promises";
import { resolve, dirname, extname, sep, isAbsolute, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";
import { createInterface } from "node:readline";
import { createBridge } from "../daemon/bridge.mjs";
import {
  loadLocalTransport,
  createLocalServer,
  TASKPANE_PORT,
  BRIDGE_PORT,
} from "../daemon/local-transport.mjs";

const assets = fileURLToPath(new URL("../taskpane/", import.meta.url));
const MIME = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".mjs": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
};
const CONTROL = "NATIVE QA";
const WORD_SUMMARY =
  "Read the whole document and give me a tight summary — main argument, key points, anything notable. Don't edit the document.";
const EXCEL_SUMMARY =
  "List the worksheets, then read the active sheet's used range and give me a tight summary — what the data is, columns, row count, anything notable. Don't change anything.";
const MODEL = "synthetic-qa-no-model";
const MARKER = "DRAFTSPECT SYNTHETIC QA 20260912";
function documentPath(value) {
  try {
    return value?.startsWith("file:")
      ? resolve(fileURLToPath(value))
      : isAbsolute(value || "")
        ? resolve(value)
        : null;
  } catch {
    return null;
  }
}

export async function createNativePaneQa({
  env = process.env,
  httpPort = TASKPANE_PORT,
  wsPort = BRIDGE_PORT,
} = {}) {
  // A QA-only UI slot can reuse the observed selected key, avoiding changes to
  // the user's sticky model preference. No provider consumes this value.
  const modelSlot = env.DRAFTSPECT_NATIVE_QA_MODEL_KEY || MODEL;
  if (!/^[a-zA-Z0-9_.\[\]-]+$/.test(modelSlot)) throw new Error("Invalid synthetic UI model slot");
  const document = documentPath(env.DRAFTSPECT_NATIVE_QA_DOCUMENT);
  const host = document && { ".docx": "word", ".xlsx": "excel" }[extname(document).toLowerCase()];
  if (!host || !(await stat(document)).isFile())
    throw new Error(
      "DRAFTSPECT_NATIVE_QA_DOCUMENT must name one existing synthetic .docx or .xlsx file",
    );
  const summary = host === "excel" ? EXCEL_SUMMARY : WORD_SUMMARY;
  const reportDir = env.DRAFTSPECT_QA_REPORT_DIR;
  if (!reportDir || !isAbsolute(reportDir))
    throw new Error("Set an absolute DRAFTSPECT_QA_REPORT_DIR");
  const transport = await loadLocalTransport(env),
    token = randomUUID();
  const evidence = { synthetic_only: true, model_requests: 0, document, host, events: [] };
  let paneKey = null,
    armed = null,
    running = false;
  const save = async (event) => {
    evidence.events.push({ at: new Date().toISOString(), ...event });
    await mkdir(reportDir, { recursive: true });
    const path = join(reportDir, `native-pane-${Date.now()}-${randomUUID()}.json`);
    await writeFile(path, JSON.stringify(evidence, null, 2) + "\n", { flag: "wx", mode: 0o600 });
    console.log(`QA report: ${path}`);
  };
  const server = createLocalServer({
    port: httpPort,
    tls: transport.tls,
    handler: async (req, res) => {
      res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0");
      try {
        if (req.method !== "GET") return res.writeHead(405).end();
        const route = new URL(req.url, "http://127.0.0.1").pathname;
        if (route === "/bridge-token")
          return res.writeHead(200, { "Content-Type": "text/plain" }).end(token);
        const path = resolve(assets, "." + route),
          ext = extname(path);
        if (!path.startsWith(assets.endsWith(sep) ? assets : assets + sep) || !MIME[ext])
          return res.writeHead(403).end();
        let body = await readFile(path);
        if (ext === ".html")
          body = body
            .toString()
            .replace(/(\/shared\/(?:taskpane\.js|styles\.css))"/g, `$1?v=${token}"`);
        if (ext === ".js" || ext === ".mjs")
          body = body
            .toString()
            .replace(/(from\s*["'])(\.[^"'?]*\.(?:js|mjs))(["'])/g, `$1$2?v=${token}$3`);
        if (route === `/${host}/index.html`) await save({ event: "pane_asset_request", route });
        res.writeHead(200, { "Content-Type": MIME[ext] }).end(body);
      } catch (error) {
        res.writeHead(error.code === "ENOENT" ? 404 : 500).end("QA asset error");
      }
    },
  });
  await server.ready;
  const origin = `${transport.secure ? "https" : "http"}://127.0.0.1:${server.address().port}`;
  const info = { provider: "synthetic-qa", state: host, session_id: null };
  const rpc = (result) => async (msg, reply) => {
    reply({ type: msg.type + "_result", request_id: msg.request_id, ...result });
    await save({ event: "rpc", type: msg.type, ok: result.ok });
  };
  const bridge = createBridge({
    port: wsPort,
    tls: transport.tls,
    token,
    allowedOrigins: [origin],
    runtimeId: "synthetic-qa",
    allowHello: (candidateHost, doc) => candidateHost === host && documentPath(doc) === document,
    onHello: (key) => {
      paneKey = key;
      bridge.sendToTaskpane({ type: "transcript_replay", events: [], session_info: info }, key);
      bridge.sendAssistantEvent(
        {
          event: "info",
          message:
            "Synthetic native QA only. No model, provider session, history or user MCP is connected.",
        },
        key,
      );
      return save({ event: "allowed_hello", protocol: transport.secure ? "wss" : "ws" });
    },
    onClose: (key) => {
      if (paneKey === key) {
        paneKey = null;
        armed = null;
      }
    },
    onUserMessage: (key, _host, _doc, text) => {
      if (
        !armed ||
        Date.now() > armed.expires ||
        running ||
        (text !== CONTROL && !(armed.command === "read" && text === summary))
      )
        throw new Error(
          "No model is connected. Arm an allowed command in the QA terminal, then send exactly NATIVE QA within two minutes.",
        );
      const { command } = armed;
      armed = null;
      running = true;
      setImmediate(() =>
        run(key, command).catch((error) =>
          console.error("QA command/report failed:", error.message),
        ),
      );
    },
    extraHandlers: {
      list_models: rpc({
        ok: true,
        runtime: "synthetic-qa",
        models: [{ value: modelSlot, label: "Synthetic QA — no model" }],
      }),
      set_model: async (msg, reply) =>
        rpc(
          msg.model === modelSlot
            ? { ok: true, model: modelSlot }
            : { ok: false, error: "Only synthetic QA is available" },
        )(msg, reply),
      get_cwd_state: rpc({ ok: true, current_cwd: dirname(document) }),
      get_context: rpc({ ok: true, entries: [] }),
      get_session_info: rpc({ ok: true, info }),
      set_context: rpc({ ok: false, error: "QA does not change context files" }),
      set_cwd: rpc({ ok: false, error: "QA workspace is fixed; no workspace files are written" }),
      pick_path: rpc({ ok: false, error: "QA does not expose the filesystem picker" }),
      release_session: rpc({ ok: false, error: "QA has no provider session to archive" }),
      return_session: rpc({ ok: false, error: "QA has no provider session to resume" }),
      stop_agent: async (msg, reply, key) => {
        armed = null;
        bridge.clearUserMessages(key);
        bridge.cancelTools(key);
        bridge.sendAssistantEvent(
          { event: "turn_complete", turn_id: bridge.activeTurn(key), interrupted: true },
          key,
        );
        reply({ type: "stop_agent_result", request_id: msg.request_id, ok: true });
      },
    },
  });
  async function run(key, command) {
    let turnId;
    const calls = [];
    try {
      const turn = await bridge.nextUserMessage(key);
      turnId = turn.turn_id;
      const call = async (name, args) => {
        try {
          const result = await bridge.callTaskpaneTool(name, args, key);
          calls.push({ name, args, result });
          return result;
        } catch (error) {
          calls.push({ name, args, error: error.message });
          throw error;
        }
      };
      if (host === "excel") {
        const sheets = await call("excel_list_sheets", {});
        if (sheets.sheets?.length !== 1 || sheets.sheets[0].name !== "Sheet1")
          throw new Error("REFUSED: one synthetic Sheet1 is required");
        const read = await call("excel_read_range", { sheet: "Sheet1" });
        if (read.values?.[0]?.[0] !== MARKER)
          throw new Error("REFUSED: first cell is not the synthetic QA marker");
        if (command === "sort") {
          if (JSON.stringify(read.values) !== JSON.stringify([[MARKER]]))
            throw new Error("REFUSED: fresh marker-only workbook required before any write");
          await call("excel_write_range", {
            sheet: "Sheet1",
            address: "A3:B5",
            values: [
              ["key", "value"],
              [2, 20],
              [1, 10],
            ],
          });
          await call("excel_sort_range", {
            sheet: "Sheet1",
            address: "A3:B5",
            key: 0,
            has_headers: true,
          });
          const sorted = await call("excel_read_range", { sheet: "Sheet1", address: "A3:B5" });
          if (
            JSON.stringify(sorted.values) !==
            JSON.stringify([
              ["key", "value"],
              [1, 10],
              [2, 20],
            ])
          )
            throw new Error("RESULT_UNKNOWN: sorted header/row values did not match");
          await call("excel_write_range", {
            sheet: "Sheet1",
            address: "D3:D4",
            values: [[2], [1]],
          });
          await call("excel_write_formula", {
            sheet: "Sheet1",
            address: "E3:E4",
            formulas: [["=D3*10"], ["=D4*10"]],
          });
          await call("excel_sort_range", { sheet: "Sheet1", address: "D3:E4", key: 0 });
          const relative = await call("excel_read_range", { sheet: "Sheet1", address: "D3:E4" });
          if (
            JSON.stringify(relative.values) !==
            JSON.stringify([
              [1, 10],
              [2, 20],
            ])
          )
            throw new Error("RESULT_UNKNOWN: sorted relative-formula values did not match");
        }
      } else {
        const read = await call("office_read_story", {
          target: { story: "body" },
          version: "Current",
          limit: 500,
        });
        if (read.paragraphs?.[0]?.text?.trim() !== MARKER)
          throw new Error("REFUSED: first paragraph is not the synthetic QA marker");
        if (command === "format") {
          const hits =
            read.paragraphs?.filter((p) => p.text.replace(/\r$/, "") === "Formatting target.") ||
            [];
          if (read.next_offset !== null || hits.length !== 1 || !hits[0].id)
            throw new Error(
              "Expected a complete synthetic body with one stable-ID Formatting target. paragraph",
            );
          const target = { story: "body", paragraph_id: hits[0].id, query: "Formatting target." };
          const before = await call("office_read_formatting", { target });
          const size = before.font?.size === 13 ? 14 : 13;
          await call("office_edit_range", {
            target,
            action: "format",
            font: { size },
            track_changes: true,
          });
          const after = await call("office_read_formatting", { target });
          if (after.font?.size !== size || after.current !== "Formatting target.")
            throw new Error("RESULT_UNKNOWN: formatting/text readback did not match");
        }
      }
      await save({ event: "native_command", command, ok: true, calls });
      bridge.sendAssistantEvent(
        {
          event: "info",
          message: `Synthetic ${command} returned; native results saved locally. No model was called.`,
        },
        key,
      );
      bridge.sendAssistantEvent(
        { event: "turn_complete", turn_id: turnId, subtype: "synthetic_qa_complete" },
        key,
      );
    } catch (error) {
      await save({
        event: "native_command",
        command,
        ok: false,
        calls,
        error: error.message,
      }).catch((failure) => console.error("QA report save failed:", failure.message));
      bridge.sendAssistantEvent(
        { event: "turn_complete", turn_id: turnId || bridge.activeTurn(key), error: error.message },
        key,
      );
    } finally {
      running = false;
    }
  }
  try {
    await bridge.ready;
  } catch (error) {
    await Promise.allSettled([server.close(), bridge.close()]);
    throw error;
  }
  return {
    origin,
    host,
    diagnosticSummary: bridge.diagnosticSummary,
    bridgeAddress: bridge.address,
    evidence,
    arm(command) {
      if (!(host === "excel" ? ["read", "sort"] : ["read", "format"]).includes(command))
        throw new Error(`Commands: read, ${host === "excel" ? "sort" : "format"}, quit`);
      if (!paneKey || running || bridge.activeTurn(paneKey))
        throw new Error("Wait for the allowed pane and an idle turn");
      armed = { command, expires: Date.now() + 120000 };
      return `Armed ${command} for two minutes. Send ${CONTROL}${command === "read" ? ` or click Summarize this ${host === "excel" ? "sheet" : "document"}` : ""}. No model is connected.`;
    },
    close: () => Promise.all([bridge.close(), server.close()]),
  };
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const qa = await createNativePaneQa();
  console.log(
    `Model-free production-pane QA: ${qa.origin}/${qa.host}/index.html\nTerminal commands: read, ${qa.host === "excel" ? "sort" : "format"}, quit`,
  );
  const input = createInterface({ input: process.stdin });
  let closing = false;
  const close = async () => {
    if (closing) return;
    closing = true;
    input.close();
    await qa.close();
  };
  input.on("line", (line) => {
    if (line.trim() === "quit") return void close();
    try {
      console.log(qa.arm(line.trim()));
    } catch (error) {
      console.error(error.message);
    }
  });
  process.on("SIGINT", close);
  process.on("SIGTERM", close);
}
