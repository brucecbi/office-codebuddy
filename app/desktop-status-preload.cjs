const { contextBridge, ipcRenderer } = require("electron");
const actions = new Set([
  "start",
  "restart",
  "stop",
  "initialize-word",
  "initialize-excel",
  "open-logs",
  "quit",
]);
contextBridge.exposeInMainWorld(
  "draftspectDesktop",
  Object.freeze({
    getStatus: (...args) => {
      if (args.length !== 0) return Promise.reject(new Error("DESKTOP_IPC_DENIED"));
      return ipcRenderer.invoke("draftspect:desktop-status:get");
    },
    act: (...args) => {
      if (args.length !== 1 || typeof args[0] !== "string" || !actions.has(args[0]))
        return Promise.reject(new Error("DESKTOP_IPC_DENIED"));
      return ipcRenderer.invoke("draftspect:desktop-status:act", args[0]);
    },
  }),
);
