// Execute only the real parent-disconnect/shutdown code. No runtime import,
// model, business document, production process or native application is used.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { spawn } from "node:child_process";
const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
const parentGuard = source.slice(
  source.indexOf('if (process.send) process.once("disconnect"'),
  source.indexOf("// Draftspect deliberately uses 47833/47834."),
);
const shutdown = source.slice(
  source.indexOf("async function shutdown()"),
  source.indexOf('process.on("SIGINT", shutdown)'),
);
const turn = () => new Promise((resolve) => setImmediate(resolve));

function fixture({ ipc = true, stuck = false } = {}) {
  const exits = [],
    calls = [],
    timers = [],
    handlers = new Map();
  const context = vm.createContext({
    process: {
      send: ipc ? () => {} : undefined,
      once: (event, fn) => handlers.set(event, fn),
      exit: (code) => exits.push(code),
    },
    setTimeout: (fn, ms) => {
      const timer = { fn, ms, cancelled: false };
      timers.push(timer);
      return timer;
    },
    clearTimeout: (timer) => {
      if (timer) timer.cancelled = true;
    },
    shuttingDown: false,
    console: {
      log: () => calls.push("local shutdown notice"),
      error: () => calls.push("shutdown-unconfirmed"),
    },
    sessions: new Map([["own", { key: "own", cancel: () => calls.push("cancel-owned") }]]),
    runtime: { id: "codebuddy" },
    runtimeClosures: {
      closeAll: () => {
        calls.push("all-runtime-barriers");
        return stuck ? new Promise(() => {}) : Promise.resolve();
      },
      has: () => true,
      wait: (key, ms) => {
        calls.push({ key, ms });
        return stuck ? new Promise(() => {}) : Promise.resolve();
      },
    },
    http: {
      close: () => {
        calls.push("close-http");
        return Promise.resolve();
      },
    },
    bridge: {
      close: () => {
        calls.push("close-bridge");
        return Promise.resolve();
      },
    },
  });
  vm.runInContext(parentGuard + "\n" + shutdown, context);
  return { exits, calls, timers, handlers, context };
}
test("IPC parent death cancels owned work, closes listeners and awaits captured SDK exit barrier", async () => {
  const f = fixture();
  f.handlers.get("disconnect")();
  await turn();
  assert.deepEqual(f.exits, [0]);
  assert.ok(f.calls.includes("cancel-owned"));
  assert.ok(f.calls.includes("close-http"));
  assert.ok(f.calls.includes("close-bridge"));
  assert.deepEqual(
    f.calls.find((value) => value?.key),
    { key: "own", ms: 2500 },
  );
  assert.ok(f.timers.every((timer) => timer.cancelled));
});
test("unconfirmed SDK shutdown is bounded by current-daemon self-exit, without any other PID kill", async () => {
  const f = fixture({ stuck: true });
  f.handlers.get("disconnect")();
  f.timers.find((timer) => timer.ms === 3000 && !timer.cancelled).fn();
  assert.deepEqual(f.exits, [1]);
  assert.ok(f.calls.includes("shutdown-unconfirmed"));
  assert.doesNotMatch(JSON.stringify(f.calls), /kill|restart|resume|replay/);
});
test("standalone developer daemon does not acquire parent-disconnect behavior", () => {
  const f = fixture({ ipc: false });
  assert.equal(f.handlers.size, 0);
  assert.deepEqual(f.exits, []);
});
test("disconnect before startup resource initialization still self-exits safely and boundedly", async () => {
  const f = fixture();
  delete f.context.sessions;
  f.handlers.get("disconnect")();
  await turn();
  assert.deepEqual(f.exits, [0]);
  assert.equal(f.timers[0].cancelled, true);
});

test("real isolated Node IPC disconnect runs actual guard/shutdown and exits without an orphan", async () => {
  const entry = `import vm from "node:vm";
const emit = (event) => console.log(JSON.stringify({ event }));
const context = vm.createContext({ process, setTimeout, clearTimeout,
  shuttingDown:false, console:{log:()=>emit("shutdown"),error:()=>emit("shutdown-unconfirmed")},
  sessions:new Map([["synthetic",{key:"synthetic",cancel:()=>emit("cancel-owned")}]]),
  runtime:{id:"codebuddy"},runtimeClosures:{closeAll:async()=>emit("all-runtime-barriers"),has:()=>true,wait:async()=>emit("owned-runtime-barrier")},
  http:{close:async()=>emit("http-closed")},bridge:{close:async()=>emit("bridge-closed")} });
vm.runInContext(${JSON.stringify(parentGuard + "\n" + shutdown)},context);
emit("ready");`;
  const child = spawn(process.execPath, ["--input-type=module", "-e", entry], {
    stdio: ["ignore", "pipe", "pipe", "ipc"],
  });
  const events = [],
    errors = [];
  let pending = "",
    timer;
  const exited = new Promise((resolve, reject) => {
    child.once("error", reject);
    child.once("exit", (code, signal) => resolve({ code, signal }));
  });
  try {
    await new Promise((resolve, reject) => {
      timer = setTimeout(() => reject(new Error("isolated IPC startup timeout")), 3000);
      child.stdout.on("data", (value) => {
        pending += String(value);
        const lines = pending.split("\n");
        pending = lines.pop();
        for (const line of lines) {
          const event = JSON.parse(line);
          events.push(event.event);
          if (event.event === "ready") resolve();
        }
      });
      child.stderr.on("data", (value) => errors.push(String(value)));
      child.once("error", reject);
    });
    clearTimeout(timer);
    child.disconnect();
    const result = await Promise.race([
      exited,
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error("isolated IPC exit timeout")), 4000);
      }),
    ]);
    assert.deepEqual(result, { code: 0, signal: null });
    assert.ok(events.includes("cancel-owned"));
    assert.ok(events.includes("http-closed"));
    assert.ok(events.includes("bridge-closed"));
    assert.ok(events.includes("owned-runtime-barrier"));
    assert.deepEqual(errors, []);
  } finally {
    clearTimeout(timer);
    if (child.exitCode == null && child.signalCode == null) child.kill("SIGTERM");
    await exited;
  }
});
