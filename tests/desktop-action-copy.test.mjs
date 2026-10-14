import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
const source = await readFile(new URL("../app/desktop-status.js", import.meta.url), "utf8");

test("actual renderer action copy distinguishes requested startup/stop from confirmed completion", async () => {
  for (const [action, phase, expected] of [
    ["start", "Starting…", "启动请求已提交，等待本地服务确认就绪。"],
    ["restart", "Starting…", "启动请求已提交，等待本地服务确认就绪。"],
    ["stop", "Stopping — waiting for owned service exit", "已请求停止，等待自有服务退出。"],
    ["stop", "Stopped", "操作已完成。"],
    ["initialize-word", "Stopped", "操作已完成。"],
  ]) {
    const fields = new Map();
    let click;
    const button = {
      dataset: { action },
      addEventListener: (_event, fn) => {
        click = fn;
      },
    };
    const context = {
      document: {
        querySelectorAll: () => [button],
        getElementById: (id) => {
          if (!fields.has(id)) fields.set(id, { textContent: "", hidden: false });
          return fields.get(id);
        },
      },
      window: {
        draftspectDesktop: {
          getStatus: async () => ({ status: phase }),
          act: async () => ({ ok: true, status: { status: phase } }),
        },
        setInterval: () => 1,
        clearInterval() {},
        addEventListener() {},
      },
    };
    vm.runInNewContext(source, context);
    await click();
    assert.equal(fields.get("action-result").textContent, expected, `${action}: ${phase}`);
  }
});
