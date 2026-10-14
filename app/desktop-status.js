/* global window, document */
"use strict";
(() => {
  const api = window.draftspectDesktop;
  const buttons = [...document.querySelectorAll("button[data-action]")];
  const fields = Object.fromEntries(
    ["status", "runtime", "detail", "error", "action-result"].map((id) => [
      id,
      document.getElementById(id),
    ]),
  );
  let busy = false;
  let reading = false;
  const text = (value, limit) => (typeof value === "string" ? value.slice(0, limit) : "");
  function render(value) {
    const status = value && typeof value === "object" ? value : {};
    fields.status.textContent = text(status.status, 120) || "状态暂不可用";
    fields.runtime.textContent = text(status.runtime, 120)
      ? "运行引擎：" + text(status.runtime, 120)
      : "";
    fields.detail.textContent = text(status.detail, 4000);
    fields.error.textContent = text(status.error, 2000);
    fields.error.hidden = !fields.error.textContent;
  }
  async function refresh() {
    if (reading) return;
    reading = true;
    try {
      render(await api.getStatus());
    } catch {
      render({
        status: "状态暂不可用",
        error: "无法连接本地控制窗口。请重新打开 Draftspect 或查看日志。",
      });
    } finally {
      reading = false;
    }
  }
  for (const button of buttons)
    button.addEventListener("click", async () => {
      if (busy) return;
      busy = true;
      buttons.forEach((item) => {
        item.disabled = true;
      });
      fields["action-result"].textContent = "正在处理…";
      try {
        const result = await api.act(button.dataset.action);
        const phase = text(result?.status?.status, 120);
        const requestedStart =
          ["start", "restart"].includes(button.dataset.action) && /Starting|正在启动/i.test(phase);
        const requestedStop = button.dataset.action === "stop" && /Stopping|正在停止/i.test(phase);
        fields["action-result"].textContent =
          result?.ok === true
            ? requestedStart
              ? "启动请求已提交，等待本地服务确认就绪。"
              : requestedStop
                ? "已请求停止，等待自有服务退出。"
                : "操作已完成。"
            : text(result?.error, 2000) || "操作未完成，请查看状态。";
      } catch {
        fields["action-result"].textContent = "操作未完成，请检查服务状态或查看日志。";
      } finally {
        busy = false;
        buttons.forEach((item) => {
          item.disabled = false;
        });
        await refresh();
      }
    });
  void refresh();
  const timer = window.setInterval(() => {
    void refresh();
  }, 1000);
  window.addEventListener("beforeunload", () => window.clearInterval(timer), { once: true });
})();
