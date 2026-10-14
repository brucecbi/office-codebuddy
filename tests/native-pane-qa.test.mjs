import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, writeFile, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { once } from "node:events";
import http from "node:http";
import WebSocket from "ws";
import { createNativePaneQa } from "../scripts/native-pane-qa.mjs";

function get(url) {
  return new Promise((resolve, reject) => {
    http
      .get(url, { agent: false }, (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () =>
          resolve({ status: res.statusCode, text: Buffer.concat(chunks).toString() }),
        );
      })
      .on("error", reject);
  });
}

test("model-free native pane QA serves actual assets, rejects other docs/messages and returns real turn-tagged tool results", async () => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-pane-qa-test-"));
  const document = join(dir, "synthetic fixture.docx"),
    reportDir = join(dir, "reports");
  // This test exercises the wire contract, not a native DOCX parser.
  await writeFile(document, "synthetic wire fixture");
  let qa;
  const clients = [];
  try {
    qa = await createNativePaneQa({
      env: {
        DRAFTSPECT_NATIVE_QA_DOCUMENT: document,
        DRAFTSPECT_QA_REPORT_DIR: reportDir,
        DRAFTSPECT_STATE_DIR: join(dir, "isolated-state"),
      },
      httpPort: 0,
      wsPort: 0,
    });
    const token = (await get(qa.origin + "/bridge-token")).text;
    const html = await get(qa.origin + "/word/index.html");
    assert.equal(html.status, 200);
    assert.ok(html.text.includes(`/shared/taskpane.js?v=${token}`));
    const js = await get(qa.origin + "/shared/taskpane.js");
    assert.ok(js.text.includes(`./word-review.js?v=${token}`));
    assert.equal(
      (await get(qa.origin + "/report")).status,
      403,
      "No HTTP command/report upload endpoint",
    );
    const messages = [];
    const connect = async (activeDoc) => {
      const ws = new WebSocket(`ws://127.0.0.1:${qa.bridgeAddress().port}`, { origin: qa.origin });
      clients.push(ws);
      ws.on("message", (data) => messages.push({ ws, ...JSON.parse(String(data)) }));
      await once(ws, "open");
      const close = once(ws, "close");
      ws.send(JSON.stringify({ type: "hello", host: "word", active_doc: activeDoc, token }));
      return { ws, close };
    };
    const forbidden = await connect(join(dir, "not-authorized.docx"));
    assert.equal((await forbidden.close)[0], 4007);
    assert.ok(!messages.some((m) => m.ws === forbidden.ws && m.type === "welcome"));
    const { ws } = await connect(pathToFileURL(document).href);
    const waitFor = async (predicate) => {
      const deadline = Date.now() + 2500;
      while (Date.now() < deadline) {
        const found = messages.find((m) => m.ws === ws && predicate(m));
        if (found) return found;
        await new Promise((resolve) => setTimeout(resolve, 5));
      }
      throw new Error("Synthetic QA response timed out");
    };
    assert.equal((await waitFor((m) => m.type === "welcome")).runtime, "synthetic-qa");
    await waitFor((m) => m.type === "transcript_replay");
    for (const type of ["list_models", "get_context", "get_session_info", "get_cwd_state"]) {
      ws.send(JSON.stringify({ type, request_id: type }));
      assert.equal((await waitFor((m) => m.request_id === type)).ok, true);
    }
    ws.send(
      JSON.stringify({ type: "user_message", text: "ordinary user request", turn_id: "denied" }),
    );
    const denied = await waitFor((m) => m.event === "turn_complete" && m.turn_id === "denied");
    assert.match(denied.error, /No model is connected/);
    assert.ok(!messages.some((m) => m.type === "tool_call"));
    assert.match(qa.arm("read"), /NATIVE QA/);
    ws.send(JSON.stringify({ type: "user_message", text: "NATIVE QA", turn_id: "read-turn" }));
    const read = await waitFor((m) => m.type === "tool_call" && m.turn_id === "read-turn");
    assert.equal(read.name, "office_read_story");
    const result = {
      next_offset: null,
      paragraphs: [
        { id: "P0", text: "DRAFTSPECT SYNTHETIC QA 20260912" },
        { id: "P1", text: "Formatting target.\r" },
      ],
    };
    ws.send(JSON.stringify({ type: "tool_result", id: read.id, ok: true, result }));
    await waitFor((m) => m.event === "turn_complete" && m.turn_id === "read-turn");
    assert.equal(
      qa.evidence.events.find((e) => e.command === "read").calls[0].result.paragraphs[1].text,
      "Formatting target.\r",
    );
    assert.match(qa.arm("format"), /NATIVE QA/);
    ws.send(JSON.stringify({ type: "user_message", text: "NATIVE QA", turn_id: "format-turn" }));
    const answered = new Set();
    let formatReads = 0;
    const deadline = Date.now() + 2500;
    while (
      !messages.some((m) => m.event === "turn_complete" && m.turn_id === "format-turn") &&
      Date.now() < deadline
    ) {
      for (const call of messages.filter(
        (m) => m.type === "tool_call" && m.turn_id === "format-turn" && !answered.has(m.id),
      )) {
        answered.add(call.id);
        const output =
          call.name === "office_read_story"
            ? result
            : call.name === "office_edit_range"
              ? { verified: true }
              : { current: "Formatting target.", font: { size: ++formatReads === 1 ? 11 : 13 } };
        ws.send(JSON.stringify({ type: "tool_result", id: call.id, ok: true, result: output }));
      }
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    const formatted = qa.evidence.events.find((e) => e.command === "format");
    assert.equal(formatted.ok, true);
    const write = formatted.calls.find((c) => c.name === "office_edit_range");
    assert.equal(write.args.font.size, 13);
    assert.equal(write.args.target.paragraph_id, "P1");
    assert.equal(qa.evidence.model_requests, 0);
    qa.arm("format");
    ws.send(JSON.stringify({ type: "user_message", text: "NATIVE QA", turn_id: "missing-marker" }));
    const markerRead = await waitFor(
      (m) => m.type === "tool_call" && m.turn_id === "missing-marker",
    );
    ws.send(
      JSON.stringify({
        type: "tool_result",
        id: markerRead.id,
        ok: true,
        result: { next_offset: null, paragraphs: [{ id: "P1", text: "Formatting target." }] },
      }),
    );
    const refused = await waitFor(
      (m) => m.event === "turn_complete" && m.turn_id === "missing-marker",
    );
    assert.match(refused.error, /synthetic QA marker/);
    assert.equal(
      messages.filter((m) => m.type === "tool_call" && m.turn_id === "missing-marker").length,
      1,
    );
    const reports = await readdir(reportDir);
    assert.ok(reports.length >= 3);
    assert.ok(reports.every((name) => name.endsWith(".json")));
    for (const file of reports)
      assert.equal(JSON.parse(await readFile(join(reportDir, file), "utf8")).synthetic_only, true);
  } finally {
    // qa.close must release native WebSocket clients without asking Word to quit.
    await qa?.close();
    clients.forEach((ws) => ws.terminate());
    await rm(dir, { recursive: true });
  }
});

test("native pane QA refuses missing document/report scope before opening a listener", async () => {
  await assert.rejects(createNativePaneQa({ env: {}, httpPort: 0, wsPort: 0 }), /synthetic/);
});

test("Excel production-pane QA enforces host and fresh marker-only write scope without any provider", async (t) => {
  const dir = await mkdtemp(join(tmpdir(), "draftspect-excel-pane-qa-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const document = join(dir, "synthetic.xlsx");
  await writeFile(document, "synthetic wire fixture");
  const qa = await createNativePaneQa({
    env: {
      DRAFTSPECT_NATIVE_QA_DOCUMENT: document,
      DRAFTSPECT_QA_REPORT_DIR: join(dir, "reports"),
      DRAFTSPECT_STATE_DIR: join(dir, "state"),
    },
    httpPort: 0,
    wsPort: 0,
  });
  t.after(() => qa.close());
  assert.equal(qa.host, "excel");
  assert.throws(() => qa.arm("format"), /Commands/);
  const token = (await get(qa.origin + "/bridge-token")).text;
  assert.equal((await get(qa.origin + "/excel/index.html")).status, 200);
  const rejected = new WebSocket(`ws://127.0.0.1:${qa.bridgeAddress().port}`, {
    origin: qa.origin,
  });
  await once(rejected, "open");
  const denied = once(rejected, "close");
  rejected.send(JSON.stringify({ type: "hello", host: "word", active_doc: document, token }));
  assert.equal((await denied)[0], 4007);
  const ws = new WebSocket(`ws://127.0.0.1:${qa.bridgeAddress().port}`, { origin: qa.origin });
  const messages = [];
  ws.on("message", (bytes) => messages.push(JSON.parse(String(bytes))));
  await once(ws, "open");
  const wait = async (predicate) => {
    const until = Date.now() + 2500;
    while (Date.now() < until) {
      const hit = messages.find(predicate);
      if (hit) return hit;
      await new Promise((resolve) => setTimeout(resolve, 5));
    }
    throw new Error("Excel wire fixture timeout");
  };
  ws.send(
    JSON.stringify({
      type: "hello",
      host: "excel",
      active_doc: pathToFileURL(document).href,
      token,
    }),
  );
  await wait((m) => m.type === "welcome");
  assert.deepEqual(qa.diagnosticSummary().connected_panes_by_host, { word: 0, excel: 1 });
  for (const [turn, values, ok] of [
    ["read", [["DRAFTSPECT SYNTHETIC QA 20260912"]], true],
    ["no-marker", [["business content"]], false],
    ["not-fresh", [["DRAFTSPECT SYNTHETIC QA 20260912", "other content"]], false],
  ]) {
    qa.arm(turn === "read" ? "read" : "sort");
    ws.send(JSON.stringify({ type: "user_message", text: "NATIVE QA", turn_id: turn }));
    const sheets = await wait((m) => m.type === "tool_call" && m.turn_id === turn);
    assert.equal(sheets.name, "excel_list_sheets");
    ws.send(
      JSON.stringify({
        type: "tool_result",
        id: sheets.id,
        ok: true,
        result: { sheets: [{ name: "Sheet1" }] },
      }),
    );
    const read = await wait(
      (m) => m.type === "tool_call" && m.turn_id === turn && m.name === "excel_read_range",
    );
    ws.send(JSON.stringify({ type: "tool_result", id: read.id, ok: true, result: { values } }));
    const end = await wait((m) => m.event === "turn_complete" && m.turn_id === turn);
    assert.equal(!end.error, ok);
    assert.ok(
      !messages.some(
        (m) =>
          m.turn_id === turn &&
          m.type === "tool_call" &&
          !["excel_list_sheets", "excel_read_range"].includes(m.name),
      ),
    );
  }
  assert.equal(qa.evidence.model_requests, 0);
  ws.terminate();
});
