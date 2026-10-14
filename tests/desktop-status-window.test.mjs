import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter } from "node:events";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { createDesktopStatusWindow } from "../app/desktop-status-window.mjs";

const page = new URL("../app/desktop-status.html", import.meta.url).href;
const GET = "draftspect:desktop-status:get";
const ACT = "draftspect:desktop-status:act";
function fixture({
  readStatus = () => ({ status: "Ready", detail: "Local only", runtime: "codebuddy", error: "" }),
  onAction = async () => {},
  autoReady = true,
  ...options
} = {}) {
  const windows = [],
    handlers = new Map(),
    app = new EventEmitter();
  class BrowserWindow extends EventEmitter {
    constructor(options) {
      super();
      this.options = options;
      this.destroyed = false;
      this.shows = 0;
      this.hides = 0;
      this.focuses = 0;
      this.webContents = Object.assign(new EventEmitter(), {
        mainFrame: { url: page },
        getURL: () => page,
        setWindowOpenHandler: (handler) => {
          this.openHandler = handler;
        },
        session: {
          setPermissionRequestHandler: (handler) => {
            this.permissionRequest = handler;
          },
          setPermissionCheckHandler: (handler) => {
            this.permissionCheck = handler;
          },
        },
      });
      windows.push(this);
    }
    isDestroyed() {
      return this.destroyed;
    }
    setTitle(value) {
      this.title = value;
    }
    setMenu(value) {
      this.menu = value;
    }
    loadFile(value) {
      this.loaded = value;
      if (autoReady) queueMicrotask(() => this.emit("ready-to-show"));
      return Promise.resolve();
    }
    isMinimized() {
      return this.minimized === true;
    }
    restore() {
      this.minimized = false;
    }
    show() {
      this.shows++;
    }
    hide() {
      this.hides++;
    }
    focus() {
      this.focuses++;
    }
    destroy() {
      this.destroyed = true;
      this.emit("closed");
    }
  }
  const ipcMain = {
    handle: (name, fn) => handlers.set(name, fn),
    removeHandler: (name) => handlers.delete(name),
  };
  const api = createDesktopStatusWindow({
    BrowserWindow,
    ipcMain,
    app,
    readStatus,
    onAction,
    ...options,
  });
  const event = () => ({
    sender: windows[0].webContents,
    senderFrame: windows[0].webContents.mainFrame,
  });
  return { api, windows, handlers, app, event };
}

test("desktop status window uses strict local BrowserWindow and shows only on readiness", async () => {
  const f = fixture({ autoReady: false });
  assert.equal(f.windows.length, 0);
  const showing = f.api.show();
  const win = f.windows[0];
  assert.equal(win.shows, 0);
  assert.equal(win.options.show, false);
  assert.equal(win.options.webPreferences.contextIsolation, true);
  assert.equal(win.options.webPreferences.sandbox, true);
  assert.equal(win.options.webPreferences.nodeIntegration, false);
  assert.equal(win.options.webPreferences.webviewTag, false);
  assert.equal(win.options.webPreferences.webSecurity, true);
  assert.match(win.options.webPreferences.preload, /desktop-status-preload\.cjs$/);
  assert.equal(win.title, "Draftspect — Ready");
  win.emit("ready-to-show");
  await showing;
  assert.equal(win.loaded, new URL(page).pathname);
  assert.equal(win.shows, 1);
  assert.equal(win.focuses, 1);
  win.minimized = true;
  await f.api.show();
  assert.equal(f.windows.length, 1);
  assert.equal(win.minimized, false);
  assert.equal(win.shows, 2);
  f.api.dispose();
});

test("desktop navigation, child frames, windows and permission requests are denied", async () => {
  const f = fixture();
  await f.api.show();
  const win = f.windows[0];
  assert.deepEqual(win.openHandler({ url: "https://example.com" }), { action: "deny" });
  for (const name of ["will-navigate", "will-frame-navigate", "will-attach-webview"]) {
    let denied = false;
    win.webContents.emit(
      name,
      {
        preventDefault() {
          denied = true;
        },
      },
      "https://example.com",
    );
    assert.equal(denied, true, name);
  }
  let granted;
  win.permissionRequest(win.webContents, "media", (value) => {
    granted = value;
  });
  assert.equal(granted, false);
  assert.equal(win.permissionCheck(), false);
  f.api.dispose();
});

test("desktop IPC admits only its loaded window main frame and zero get arguments", async () => {
  const f = fixture();
  await f.api.show();
  const get = f.handlers.get(GET),
    good = f.event();
  assert.equal(get(good).runtime, "codebuddy");
  for (const bad of [
    undefined,
    {},
    { sender: {}, senderFrame: good.senderFrame },
    { sender: good.sender, senderFrame: { url: page } },
    { sender: good.sender },
  ])
    assert.throws(() => get(bad), /DESKTOP_IPC_DENIED/);
  assert.throws(() => get(good, "/arbitrary/path"), /DESKTOP_IPC_DENIED/);
  good.senderFrame.url = "https://example.com";
  assert.throws(() => get(good), /DESKTOP_IPC_DENIED/);
  good.senderFrame.url = page;
  good.sender.getURL = () => "file:///other.html";
  assert.throws(() => get(good), /DESKTOP_IPC_DENIED/);
  f.api.dispose();
  assert.throws(() => get(good), /DESKTOP_IPC_DENIED/);
});

test("desktop actions accept only the exact seven action names without extra arguments", async () => {
  const calls = [],
    f = fixture({
      onAction: async (action) => {
        calls.push(action);
        return { secret: "NEVER_RETURN" };
      },
    });
  await f.api.show();
  const act = f.handlers.get(ACT);
  for (const action of [
    "start",
    "restart",
    "stop",
    "initialize-word",
    "initialize-excel",
    "open-logs",
    "quit",
  ]) {
    const result = await act(f.event(), action);
    assert.equal(result.ok, true);
    assert.doesNotMatch(JSON.stringify(result), /NEVER_RETURN/);
  }
  assert.equal(calls.length, 7);
  for (const args of [
    [],
    ["run"],
    [{ action: "start" }],
    ["open-logs", "/secret"],
    [" start"],
    ["__proto__"],
  ])
    await assert.rejects(act(f.event(), ...args), /DESKTOP_IPC_DENIED/);
  await assert.rejects(act({ ...f.event(), senderFrame: {} }, "start"), /DESKTOP_IPC_DENIED/);
  assert.equal(calls.length, 7);
  f.api.dispose();
});

test("desktop actions are single-flight and failures do not leak arbitrary exception data", async () => {
  let finish,
    calls = 0;
  const f = fixture({
    onAction: () => {
      calls++;
      return new Promise((resolve) => {
        finish = resolve;
      });
    },
  });
  await f.api.show();
  const act = f.handlers.get(ACT);
  const first = act(f.event(), "restart");
  assert.equal((await act(f.event(), "stop")).ok, false);
  assert.equal(calls, 1);
  finish();
  assert.equal((await first).ok, true);
  f.api.dispose();
  const broken = fixture({
    onAction: async () => {
      throw new Error("SECRET_TEST_EXCEPTION");
    },
  });
  await broken.api.show();
  const result = await broken.handlers.get(ACT)(broken.event(), "start");
  assert.equal(result.ok, false);
  assert.doesNotMatch(JSON.stringify(result), /SECRET_TEST_EXCEPTION/);
  broken.api.dispose();
});

test("close hides and prevents late readiness reopening; will-quit disposes exactly once", async () => {
  const f = fixture({ autoReady: false });
  const showing = f.api.show();
  const win = f.windows[0];
  let prevented = 0;
  win.emit("close", {
    preventDefault() {
      prevented++;
    },
  });
  assert.equal(prevented, 1);
  assert.equal(win.hides, 1);
  assert.equal(win.destroyed, false);
  win.emit("ready-to-show");
  await showing;
  assert.equal(win.shows, 0);
  await f.api.show();
  assert.equal(win.shows, 1);
  f.app.emit("before-quit", {});
  assert.equal(win.destroyed, false);
  f.app.emit("will-quit");
  assert.equal(win.destroyed, true);
  assert.equal(f.handlers.size, 0);
  assert.equal(f.app.listenerCount("will-quit"), 0);
  f.api.dispose();
  await f.api.show();
  assert.equal(f.windows.length, 1);
});

test("status snapshots allow only bounded primitive fields and sanitize read failures", async () => {
  const f = fixture({
    readStatus: () => ({
      status: "X".repeat(200),
      detail: "Y".repeat(5000),
      runtime: "codebuddy\0",
      error: { secret: "NO" },
      secret: "NO",
    }),
  });
  const snapshot = f.api.refresh();
  assert.deepEqual(Object.keys(snapshot), ["status", "detail", "runtime", "error"]);
  assert.equal(snapshot.status.length, 120);
  assert.equal(snapshot.detail.length, 4000);
  assert.equal(snapshot.runtime, "codebuddy");
  assert.equal(snapshot.error, "");
  assert.doesNotMatch(JSON.stringify(snapshot), /secret/);
  f.api.dispose();
  const broken = fixture({
    readStatus: () => {
      throw new Error("SECRET_TEST_EXCEPTION");
    },
  });
  assert.equal(broken.api.refresh().status, "Unavailable");
  assert.doesNotMatch(JSON.stringify(broken.api.refresh()), /SECRET_TEST_EXCEPTION/);
  broken.api.dispose();
});

test("every visible desktop action crosses the production preload and main-process allowlists", async () => {
  const [source, html] = await Promise.all([
    readFile(new URL("../app/desktop-status-preload.cjs", import.meta.url), "utf8"),
    readFile(new URL("../app/desktop-status.html", import.meta.url), "utf8"),
  ]);
  const actions = [...html.matchAll(/data-action="([^"]+)"/g)].map((match) => match[1]);
  assert.deepEqual(actions, [
    "start",
    "restart",
    "stop",
    "initialize-word",
    "initialize-excel",
    "open-logs",
    "quit",
  ]);
  const handled = [];
  const f = fixture({ onAction: async (action) => handled.push(action) });
  await f.api.show();
  let api;
  vm.runInNewContext(source, {
    require: (name) => {
      assert.equal(name, "electron");
      return {
        contextBridge: {
          exposeInMainWorld: (_name, value) => {
            api = value;
          },
        },
        ipcRenderer: {
          invoke: (channel, ...args) => f.handlers.get(channel)(f.event(), ...args),
        },
      };
    },
  });
  try {
    for (const action of actions) assert.equal((await api.act(action)).ok, true, action);
    assert.deepEqual(handled, actions);
    for (const args of [["initialize-excel", "/user/workbook.xlsx"], ["initialize-excel "]])
      await assert.rejects(api.act(...args), /DESKTOP_IPC_DENIED/);
    assert.deepEqual(handled, actions);
  } finally {
    f.api.dispose();
  }
});

test("preload exposes only two narrow methods and refuses paths or generic IPC", async () => {
  const source = await readFile(
    new URL("../app/desktop-status-preload.cjs", import.meta.url),
    "utf8",
  );
  let exposed;
  const calls = [];
  vm.runInNewContext(source, {
    require: (name) => {
      assert.equal(name, "electron");
      return {
        contextBridge: {
          exposeInMainWorld: (name, api) => {
            assert.equal(name, "draftspectDesktop");
            exposed = api;
          },
        },
        ipcRenderer: {
          invoke: (...args) => {
            calls.push(args);
            return Promise.resolve({});
          },
        },
      };
    },
  });
  assert.deepEqual(Object.keys(exposed), ["getStatus", "act"]);
  assert.equal(Object.isFrozen(exposed), true);
  await exposed.getStatus();
  await exposed.act("start");
  await assert.rejects(exposed.act("open-logs", "/path"), /DESKTOP_IPC_DENIED/);
  await assert.rejects(exposed.act("execute"), /DESKTOP_IPC_DENIED/);
  await assert.rejects(exposed.getStatus("/path"), /DESKTOP_IPC_DENIED/);
  assert.deepEqual(calls, [[GET], [ACT, "start"]]);
});

test("local page has strict CSP and renderer displays untrusted values only as literal text", async () => {
  const html = await readFile(new URL("../app/desktop-status.html", import.meta.url), "utf8");
  assert.match(html, /default-src 'none'/);
  assert.match(html, /connect-src 'none'/);
  assert.match(html, /script-src 'self'/);
  assert.match(html, /object-src 'none'/);
  assert.doesNotMatch(html, /https?:\/\/|unsafe-inline|unsafe-eval|onclick=/);
  const source = await readFile(new URL("../app/desktop-status.js", import.meta.url), "utf8");
  assert.doesNotMatch(source, /innerHTML|outerHTML|eval\(|fetch\(|XMLHttpRequest|WebSocket/);
  const fields = new Map();
  const element = () => ({
    textContent: "",
    hidden: false,
    set innerHTML(_value) {
      throw new Error("HTML injection");
    },
  });
  const marker = '<img src=x onerror="attack()">';
  let interval,
    cleared = false,
    beforeUnload;
  vm.runInNewContext(source, {
    document: {
      querySelectorAll: () => [],
      getElementById: (id) => {
        if (!fields.has(id)) fields.set(id, element());
        return fields.get(id);
      },
    },
    window: {
      draftspectDesktop: {
        getStatus: async () => ({ status: marker, runtime: marker, detail: marker, error: marker }),
      },
      setInterval: (fn, ms) => {
        assert.equal(ms, 1000);
        interval = fn;
        return 42;
      },
      clearInterval: (id) => {
        assert.equal(id, 42);
        cleared = true;
      },
      addEventListener: (name, fn) => {
        assert.equal(name, "beforeunload");
        beforeUnload = fn;
      },
    },
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(fields.get("status").textContent, marker);
  assert.equal(fields.get("detail").textContent, marker);
  assert.equal(fields.get("error").hidden, false);
  assert.equal(typeof interval, "function");
  beforeUnload();
  assert.equal(cleared, true);
});

function virtualTimers() {
  const values = [];
  return {
    values,
    setTimeout(fn, ms) {
      const timer = { fn, ms, cancelled: false };
      values.push(timer);
      return timer;
    },
    clearTimeout(timer) {
      if (timer) timer.cancelled = true;
    },
    fire(ms) {
      const timer = values.find((item) => item.ms === ms && !item.cancelled);
      assert.ok(timer, `missing live ${ms}ms timer`);
      timer.cancelled = true;
      timer.fn();
      return timer;
    },
  };
}
const turn = () => new Promise((resolve) => setImmediate(resolve));

test("loaded HTML without native readiness times out visibly instead of resolving show", async () => {
  const timers = virtualTimers(),
    f = fixture({ autoReady: false, timers });
  try {
    const showing = f.api.show();
    await turn();
    assert.equal(f.windows[0].shows, 0);
    timers.fire(15_000);
    await assert.rejects(showing, { code: "DESKTOP_STATUS_TIMEOUT" });
  } finally {
    f.api.dispose();
  }
});

test("renderer crash recovery is 1/2/4s bounded, metadata-only and never calls any daemon action", async () => {
  const timers = virtualTimers(),
    incidents = [],
    actions = [],
    f = fixture({
      timers,
      onIncident: (value) => incidents.push(value),
      onAction: async (value) => actions.push(value),
    });
  try {
    await f.api.show();
    for (const delay of [1000, 2000, 4000]) {
      f.windows
        .at(-1)
        .webContents.emit(
          "render-process-gone",
          {},
          { reason: "crashed", exitCode: 9, message: "SECRET" },
        );
      timers.fire(delay);
      await turn();
    }
    assert.equal(f.windows.length, 4);
    f.windows.at(-1).webContents.emit("render-process-gone", {}, { reason: "oom", exitCode: 9 });
    assert.equal(incidents.at(-1).code, "DESKTOP_RENDERER_RECOVERY_EXHAUSTED");
    assert.equal(timers.values.filter((item) => !item.cancelled && item.ms <= 8000).length, 0);
    assert.deepEqual(actions, []);
    assert.doesNotMatch(JSON.stringify(incidents), /SECRET/);
  } finally {
    f.api.dispose();
  }
});

test("stable visible renderer refills recovery budget; hide and Quit revoke already queued recovery", async () => {
  for (const quit of [false, true]) {
    const timers = virtualTimers(),
      f = fixture({ timers });
    await f.api.show();
    f.windows[0].webContents.emit("render-process-gone", {}, { reason: "killed", exitCode: 9 });
    timers.fire(1000);
    await turn();
    timers.fire(45_000);
    const owned = f.windows.at(-1);
    owned.webContents.emit("render-process-gone", {}, { reason: "crashed", exitCode: 9 });
    const restart = timers.values.find((item) => !item.cancelled && item.ms === 1000);
    assert.ok(restart, "stable period must refill to first backoff");
    if (quit) f.api.dispose();
    else owned.emit("close", { preventDefault() {} });
    restart.fn(); // queued callback after cancellation must not create a window
    await turn();
    assert.equal(f.windows.length, 2);
    f.api.dispose();
  }
});

test("desktop hang is reported distinctly, no request replay, and teardown destroys before IPC unregister", async () => {
  const incidents = [],
    f = fixture({ onIncident: (value) => incidents.push(value) });
  await f.api.show();
  const win = f.windows[0],
    order = [],
    originalDestroy = win.destroy.bind(win);
  win.destroy = () => {
    order.push("destroy");
    originalDestroy();
  };
  f.handlers.delete = new Proxy(f.handlers.delete, {
    apply(target, receiver, args) {
      order.push("remove");
      return Reflect.apply(target, receiver, args);
    },
  });
  win.emit("unresponsive");
  win.emit("responsive");
  assert.equal(incidents[0].code, "DESKTOP_RENDERER_UNRESPONSIVE");
  assert.equal(incidents[1].reason, "responsive");
  f.api.dispose();
  assert.equal(order[0], "destroy");
  assert.deepEqual(order, ["destroy", "remove", "remove"]);
});

test("explicit reopen replaces a hidden dead renderer without reopening automatically", async () => {
  const timers = virtualTimers(),
    f = fixture({ timers });
  try {
    await f.api.show();
    const owned = f.windows[0];
    owned.emit("close", { preventDefault() {} });
    owned.webContents.emit("render-process-gone", {}, { reason: "crashed", exitCode: 9 });
    assert.equal(f.windows.length, 1);
    assert.equal(timers.values.filter((item) => !item.cancelled && item.ms <= 8000).length, 0);
    await f.api.show();
    assert.equal(f.windows.length, 2, "explicit show must replace the exact dead renderer");
    assert.equal(owned.isDestroyed(), true);
    assert.equal(f.windows[1].shows, 1);
  } finally {
    f.api.dispose();
  }
});
