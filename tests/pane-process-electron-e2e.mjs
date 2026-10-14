// Opt-in real Chromium DOM/CSS check. No Word, daemon, model or network access.
// Run with the installed Electron executable, followed by an absolute report directory.
import { app, BrowserWindow } from "electron";
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, isAbsolute } from "node:path";

const output = process.argv[2];
if (!output || !isAbsolute(output)) throw new Error("Supply an absolute report directory");
const temporary = mkdtempSync(join(tmpdir(), "draftspect-process-render-"));
app.setPath("userData", join(temporary, "user-data"));
app.disableHardwareAcceleration();
async function run() {
  await app.whenReady();
  app.dock?.hide();
  let window;
  const timeout = setTimeout(() => app.exit(90), 30000);
  try {
    await mkdir(output, { recursive: true });
    const source = (
      await readFile(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8")
    ).replace(/^import\s+[\s\S]*?;\s*$/gm, "");
    const queue = (
      await readFile(new URL("../taskpane/shared/operation-queue.js", import.meta.url), "utf8")
    ).replace("export function", "function");
    const css = await readFile(new URL("../taskpane/shared/styles.css", import.meta.url), "utf8");
    const html = (await readFile(new URL("../taskpane/word/index.html", import.meta.url), "utf8"))
      .replace(/<script\b[^>]*>[\s\S]*?<\/script>/g, "")
      .replace('<link rel="stylesheet" href="/shared/styles.css" />', `<style>${css}</style>`);
    const fixture = join(temporary, "pane.html");
    await writeFile(fixture, html);
    window = new BrowserWindow({
      show: false,
      width: 400,
      height: 800,
      webPreferences: {
        offscreen: true,
        contextIsolation: true,
        sandbox: true,
        partition: "process-render-test",
      },
    });
    const network = [];
    window.webContents.session.webRequest.onBeforeRequest(
      { urls: ["http://*/*", "https://*/*", "ws://*/*", "wss://*/*"] },
      (details, done) => {
        network.push(details.url);
        done({ cancel: true });
      },
    );
    await window.loadFile(fixture);
    await window.webContents.executeJavaScript(`
    globalThis.Office = { context: { document: { url: "synthetic.docx" } }, onReady() {} };
    ${queue}
    const prepareReviewWrite = async () => ({ preview: { action: "Synthetic format" }, apply: async () => ({ verified: true }) });
    const resolveWordHeadingArgs = async (_name, args) => args;
    ${source}
    globalThis.paneTest = { appendUserMessage, appendToolUse, updateToolResult, upsertAssistantMessage, runOfficeTool,
      card: id => toolCards.get(id),
      turn: id => { ws = { send() {} }; wsReady = true; activeDocUrl = "synthetic.docx"; currentTurnId = id; beginTurn(); }
    };
    void 0;
  `);
    const checks = await window.webContents.executeJavaScript(`(async () => {
    const checks = [];
    const check = (name, ok) => { if (!ok) throw new Error(name); checks.push({ name, status: "PASS" }); };
    const visible = el => el.checkVisibility({ checkVisibilityCSS: true });
    paneTest.appendUserMessage("请总结这个合成测试文档，不修改内容。");
    for (let i = 0; i < 24; i++) {
      paneTest.appendToolUse(i % 2 ? "Office execution: office_find" : "mcp__office__office_find", { query: "synthetic" }, "r" + i);
      paneTest.updateToolResult("r" + i, "Synthetic result " + i, "completed");
    }
    paneTest.upsertAssistantMessage("answer", "这是合成测试的最终回答。工具记录已收进工作过程，可按需展开查看。");
    const group = paneTest.card("r0").group;
    check("24 records produce one closed disclosure", document.querySelectorAll(".process-log").length === 1 && !group.details.open);
    check("native details hides tools but keeps answer visible", !visible(paneTest.card("r0").el) && visible(document.querySelector(".msg.assistant")));
    check("completed log has no visible error alert", !visible(group.alert));
    return checks;
  })()`);
    await writeFile(
      join(output, "collapsed.png"),
      (await window.webContents.capturePage()).toPNG(),
    );
    checks.push(
      ...(await window.webContents.executeJavaScript(`(async () => {
    const checks = [];
    const check = (name, ok) => { if (!ok) throw new Error(name); checks.push({ name, status: "PASS" }); };
    const visible = el => el.checkVisibility({ checkVisibilityCSS: true });
    const group = paneTest.card("r0").group;
    group.summary.click();
    check("summary click reveals tool cards", group.details.open && visible(paneTest.card("r0").el));
    paneTest.updateToolResult("r0", "Updated result", "completed");
    check("live updates preserve expanded state", group.details.open);
    group.summary.click();
    paneTest.updateToolResult("r1", "Synthetic tool failure", "error");
    check("provider failure visible outside closed log", !group.details.open && visible(group.alert));
    group.alert.querySelector("button").click();
    check("error detail button opens and focuses disclosure", group.details.open && document.activeElement === group.summary);
    group.summary.click();
    paneTest.appendUserMessage("请对合成内容应用格式修改。");
    paneTest.turn("approve");
    paneTest.appendToolUse("earlier read", {}, "earlier");
    paneTest.updateToolResult("earlier", "ok", "completed");
    paneTest.upsertAssistantMessage("analysis", "Long synthetic analysis.\\n".repeat(100));
    const pending = paneTest.runOfficeTool({ id: "format", name: "office_edit_range", args: { target: { query: "synthetic" }, action: "format", font: { size: 13 } }, turn_id: "approve", deadline: Date.now() + 120000 });
    for (let i = 0; i < 20; i++) await Promise.resolve();
    const card = paneTest.card("format");
    check("approval and preview visible without opening log", !card.group.details.open && card.el.open && visible(card.el) && visible(card.output) && visible(card.el.querySelector("button")));
    const viewport = document.getElementById("messages").getBoundingClientRect();
    const approve = card.el.querySelector("button").getBoundingClientRect();
    check("approval follows long analysis at the visible chat tail", document.getElementById("messages").lastElementChild === card.el && approve.top >= viewport.top && approve.bottom <= viewport.bottom);
    card.el.querySelector("button").click();
    await pending;
    check("approved receipt returns to closed log", card.state === "completed" && !visible(card.el) && !card.group.details.open && !card.el.querySelector("button"));
    return checks;
  })()`)),
    );
    await writeFile(
      join(output, "attention.png"),
      (await window.webContents.capturePage()).toPNG(),
    );
    if (network.length) throw new Error("Unexpected network attempts: " + network.length);
    const report = {
      status: "PASS",
      checks,
      network_attempts: 0,
      model_requests: 0,
      native_word: "NOT_RUN",
      fixture: temporary,
    };
    await writeFile(join(output, "render-report.json"), JSON.stringify(report, null, 2) + "\n");
    console.log(JSON.stringify(report));
  } catch (error) {
    await writeFile(join(output, "render-error.txt"), String(error.stack || error));
    console.error(error);
    process.exitCode = 1;
  } finally {
    clearTimeout(timeout);
    window?.destroy();
    app.exit(process.exitCode || 0);
  }
}
void run();
