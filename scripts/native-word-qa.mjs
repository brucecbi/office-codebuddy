// Explicit, model-free native QA. Never boot the normal daemon/user MCPs here.
// Start manually, open the installed pane in a NEW document containing only
// the marker printed below, then click Run. Quit this server after testing.
import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { fileURLToPath } from "node:url";
import { join, resolve } from "node:path";

const port = Number(process.env.DRAFTSPECT_QA_PORT || 47834);
const tlsKeyPath = process.env.DRAFTSPECT_QA_TLS_KEY;
const tlsCertPath = process.env.DRAFTSPECT_QA_TLS_CERT;
if (!!tlsKeyPath !== !!tlsCertPath)
  throw new Error("Set both DRAFTSPECT_QA_TLS_KEY and DRAFTSPECT_QA_TLS_CERT, or neither");
const tls = tlsKeyPath
  ? { key: await readFile(tlsKeyPath), cert: await readFile(tlsCertPath), minVersion: "TLSv1.2" }
  : null;
const origin = `${tls ? "https" : "http"}://127.0.0.1:${port}`;
const token = randomUUID();
const repo = fileURLToPath(new URL("../", import.meta.url));
const marker = "DRAFTSPECT SYNTHETIC QA 20260912";
const reportDir = process.env.DRAFTSPECT_QA_REPORT_DIR;
if (!reportDir) throw new Error("Set DRAFTSPECT_QA_REPORT_DIR to a task-local output directory");
const boundaryDocument = process.env.DRAFTSPECT_NATIVE_QA_DOCUMENT || null;
if (
  boundaryDocument &&
  (!boundaryDocument.startsWith("/") || !/\.(?:docx|xlsx)$/i.test(boundaryDocument))
)
  throw new Error("Native boundary QA requires an absolute disposable Office file path");
const html = `<!doctype html><html><head><meta charset="utf-8"><title>Draftspect native QA</title>
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
<style>body{font:14px system-ui;margin:18px}button{font:inherit;padding:10px}pre{white-space:pre-wrap;overflow-wrap:anywhere;font-size:11px}#state{font-weight:600}</style></head>
<body><h2>Model-free Office QA</h2><p>Only a NEW test document containing exactly this marker can be initialized:</p><pre>${marker}</pre>
<button id="run" disabled>Run synthetic Word checks</button><button id="reliability" disabled>Run reliability checks (new marker-only document)</button><button id="boundaries" disabled>Run write-boundary checks (new marker-only file)</button><button id="targets" disabled>Inspect reliability targets (read-only)</button><button id="snapshots" disabled>Inspect snapshot stability (read-only)</button><button id="revisions" disabled>Diagnose revisions (writes synthetic test text)</button><p id="state">Loading Office…</p><pre id="out"></pre>
<script type="module">
import { runNativeChecks, inspectNativeSnapshots, inspectNativeRevisionSemantics } from '/native-checks.js?v=${token}';
import { runNativeReliabilityChecks, inspectNativeReliabilityTargets } from '/native-reliability-checks.js?v=${token}';
import { runNativeWordBoundaryChecks, runNativeExcelBoundaryChecks } from '/native-write-boundary-checks.js?v=${token}';
Office.onReady(info => {
  document.querySelector('#state').textContent = 'Ready: ' + info.host + ' / ' + info.platform;
  document.querySelector('#run').disabled = info.host !== 'Word';
  document.querySelector('#snapshots').disabled = info.host !== 'Word';
  document.querySelector('#revisions').disabled = info.host !== 'Word';
  document.querySelector('#reliability').disabled = info.host !== 'Word';
  document.querySelector('#targets').disabled = info.host !== 'Word';
  const boundDocument = ${JSON.stringify(boundaryDocument)};
  const exactDocument = () => {
    const url = Office.context.document.url;
    try { return boundDocument && (url.startsWith('file:') ? decodeURIComponent(new URL(url).pathname) : url) === boundDocument; }
    catch { return false; }
  };
  document.querySelector('#boundaries').disabled = !['Word', 'Excel'].includes(info.host) || !exactDocument();
  document.querySelector('#boundaries').onclick = async () => {
    if (!exactDocument()) throw new Error('REFUSED: synthetic document identity changed');
    for (const button of document.querySelectorAll('button')) button.disabled = true;
    try {
      const run = info.host === 'Excel' ? runNativeExcelBoundaryChecks : runNativeWordBoundaryChecks;
      const checks = await run(${JSON.stringify(marker)}, row => {
        document.querySelector('#state').textContent = row.name + ': ' + row.status;
        document.querySelector('#out').textContent += JSON.stringify(row) + String.fromCharCode(10);
      });
      const response = await fetch('/report', {method:'POST',headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,suite:'write-boundaries-20261008',synthetic_document:boundDocument,checks})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Saved: ' + checks.filter(x => x.status === 'PASS').length + '/' + checks.length + ' passed';
    } catch (error) { document.querySelector('#state').textContent = 'Boundary QA failed: ' + error.message; }
  };
  document.querySelector('#targets').onclick = async () => {
    document.querySelector('#targets').disabled = true;
    try {
      const detail = await inspectNativeReliabilityTargets(${JSON.stringify(marker)});
      const response = await fetch('/report', {method:'POST',headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,checks:[{name:'reliability_target_diagnostics',status:'PASS',detail}]})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Read-only target diagnostics saved';
    } catch (error) { document.querySelector('#state').textContent = 'Diagnostics failed: ' + error.message; }
  };
  document.querySelector('#reliability').onclick = async () => {
    for (const button of document.querySelectorAll('button')) button.disabled = true;
    try {
      const checks = await runNativeReliabilityChecks(${JSON.stringify(marker)}, row => {
        document.querySelector('#state').textContent = row.name + ': ' + row.status;
        document.querySelector('#out').textContent += JSON.stringify(row) + String.fromCharCode(10);
      });
      const response = await fetch('/report', {method:'POST',headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,suite:'reliability-20260917',checks})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Saved: ' + checks.filter(x => x.status === 'PASS').length + '/' + checks.length + ' passed';
    } catch (error) { document.querySelector('#state').textContent = 'Reliability QA failed: ' + error.message; }
  };
  document.querySelector('#revisions').onclick = async () => {
    document.querySelector('#revisions').disabled = true;
    document.querySelector('#state').textContent = 'Collecting synthetic revision semantics…';
    try {
      const detail = await inspectNativeRevisionSemantics(${JSON.stringify(marker)});
      const response = await fetch('/report', {method:'POST', headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,checks:[{name:'revision_diagnostics_collected',status:'PASS',detail}]})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Revision diagnostic evidence saved; inspect outcome in report';
    } catch (error) { document.querySelector('#state').textContent = 'Revision diagnostics failed: ' + error.message; }
  };
  document.querySelector('#snapshots').onclick = async () => {
    document.querySelector('#snapshots').disabled = true;
    document.querySelector('#state').textContent = 'Reading consecutive synthetic snapshots…';
    try {
      const detail = await inspectNativeSnapshots(${JSON.stringify(marker)});
      const response = await fetch('/report', {method:'POST', headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,checks:[{name:'snapshot_diagnostics',status:'PASS',detail}]})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Snapshot diagnostics saved (read-only)';
    } catch (error) {
      document.querySelector('#state').textContent = 'Snapshot diagnostics failed: ' + error.message;
    }
  };
  document.querySelector('#run').onclick = async () => {
    document.querySelector('#run').disabled = true;
    const result = await runNativeChecks(${JSON.stringify(marker)}, row => {
      document.querySelector('#state').textContent = row.name + ': ' + row.status;
      document.querySelector('#out').textContent += JSON.stringify(row) + String.fromCharCode(10);
    });
    const summary = result.filter(x => x.status === 'PASS').length + '/' + result.length + ' passed';
    document.querySelector('#state').textContent = 'Checks finished: ' + summary + '; saving report…';
    try {
      const response = await fetch('/report', {method:'POST', headers:{'Content-Type':'application/json','X-QA-Token':${JSON.stringify(token)}},body:JSON.stringify({at:new Date().toISOString(),host:info,checks:result})});
      if (!response.ok) throw new Error('HTTP ' + response.status);
      document.querySelector('#state').textContent = 'Saved: ' + summary;
    } catch (error) {
      document.querySelector('#state').textContent = 'Report save failed: ' + error.message + '; checks: ' + summary;
    }
  };
});</script></body></html>`;
const handler = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (req.headers.host !== `127.0.0.1:${port}`) {
    res.writeHead(403).end();
    return;
  }
  try {
    const path = new URL(req.url, origin).pathname;
    // Log only a bounded route name, never query strings, headers, or payloads.
    console.log(
      `QA ${req.method} ${["/word/index.html", "/excel/index.html", "/icon-32.png", "/icon-64.png", "/native-checks.js", "/report"].includes(path) ? path : "[other route]"}`,
    );
    if (path === "/report" && req.method === "POST") {
      if (req.headers.origin !== origin || req.headers["x-qa-token"] !== token) {
        res.writeHead(403).end();
        return;
      }
      const parts = [];
      let size = 0;
      for await (const part of req) {
        size += part.length;
        if (size > 2_000_000) throw new Error("Report too large");
        parts.push(part);
      }
      const report = JSON.parse(Buffer.concat(parts).toString("utf8"));
      await mkdir(reportDir, { recursive: true });
      const prefix = report.host?.host === "Excel" ? "native-excel" : "native-word";
      const file = join(reportDir, `${prefix}-${Date.now()}.json`);
      await writeFile(file, JSON.stringify(report, null, 2) + "\n", { mode: 0o600, flag: "wx" });
      console.log(
        JSON.stringify({
          report: file,
          checks: report.checks.map(({ name, status, error }) => ({ name, status, error })),
        }),
      );
      res.writeHead(200).end("Saved");
      return;
    }
    if (req.method !== "GET") {
      res.writeHead(405).end();
      return;
    }
    if (path === "/word/index.html" || path === "/excel/index.html") {
      res.setHeader("Content-Type", "text/html");
      res.end(html);
      return;
    }
    if (path === "/icon-32.png" || path === "/icon-64.png") {
      res.setHeader("Content-Type", "image/png");
      res.end(await readFile(join(repo, "taskpane", path.slice(1))));
      return;
    }
    const allowed = [
      "/shared/word-review.js",
      "/shared/word-ooxml-state.js",
      "/shared/tools-word.js",
      "/shared/word-edit-plan.js",
      "/shared/word-write-safety.js",
      "/shared/tools-excel.js",
      "/shared/operation-queue.js",
    ];
    const file =
      path === "/native-write-boundary-checks.js"
        ? join(repo, "scripts/native-write-boundary-checks.js")
        : path === "/native-reliability-checks.js"
          ? join(repo, "scripts/native-reliability-checks.js")
          : path === "/native-checks.js"
            ? join(repo, "scripts/native-word-checks.js")
            : allowed.includes(path)
              ? resolve(repo, "taskpane", "." + path)
              : null;
    if (!file) {
      res.writeHead(404).end("QA server only; no daemon or model endpoint");
      return;
    }
    const js = (await readFile(file, "utf8")).replace(
      /(from\s*["'])([./][^"'?]*\.(?:js|mjs))(["'])/g,
      `$1$2?v=${token}$3`,
    );
    res.setHeader("Content-Type", "text/javascript");
    res.end(js);
  } catch (error) {
    res.writeHead(500).end(error.message);
  }
};
// Certificate trust is deliberately outside this script. Never disable TLS
// verification or install a root automatically as part of starting QA.
const server = tls ? createHttpsServer(tls, handler) : createHttpServer(handler);
server.listen(port, "127.0.0.1", () =>
  console.log(`Synthetic-only QA at ${origin}/word/index.html\nMarker: ${marker}`),
);
