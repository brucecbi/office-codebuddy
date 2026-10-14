// Called only inside session-replay-offline's proved OS sandbox. No model send.
// Execute exact daemon handlers with real isolated dependencies; importing the
// daemon entrypoint would start unwanted services, so use the existing VM pattern.
import assert from "node:assert/strict";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { createHash } from "node:crypto";
import vm from "node:vm";
import { unstable_v2_resumeSession } from "@tencent-ai/agent-sdk";
import { runSession } from "../daemon/runtime/runtime-codebuddy.mjs";
import { observeCodeBuddyClose } from "../daemon/runtime/codebuddy-process-close.mjs";
import { createRuntimeCloseRegistry } from "../daemon/runtime/close-registry.mjs";
import { createSessionStore } from "../daemon/sessions.mjs";
import { archiveSession, resumeCommands } from "../daemon/session-archive.mjs";
import { readTranscript, locateSessionFile } from "../daemon/transcript.mjs";

const sha = (bytes) => createHash("sha256").update(bytes).digest("hex");
function fragment(source, start, end) {
  const a = source.indexOf(start),
    b = source.indexOf(end, a);
  assert.ok(a >= 0 && b > a, "Production handler extraction boundary changed");
  return source.slice(a, b);
}
export async function checkSerialHandoff([cli, cwd, config, id, transcript, selfPort]) {
  assert.ok(selfPort, "Serial integration requires the proved self-loopback sandbox");
  assert.equal(process.env.CODEBUDDY_CONFIG_DIR, config);
  assert.equal(process.env.DRAFTSPECT_CODEBUDDY_SETTING_SOURCES, "none");
  assert.equal(process.env.CODEBUDDY_INTERNET_ENVIRONMENT, "external");
  const before = await readFile(transcript);
  const options = { file: join(config, "draftspect-serial-state.json"), legacyFile: null };
  const store = createSessionStore(options);
  const key = "word\u0000" + join(cwd, "synthetic.docx");
  await store.saveSessionId("word", cwd, id, "codebuddy", key);
  await store.saveSessionId("word", cwd, "other-document-id", "codebuddy", "other-document");
  await store.saveSessionId("word", cwd, "other-provider-id", "claude", key);
  const attachment = join(dirname(transcript), id, "tool-results", "synthetic.txt");
  await mkdir(dirname(attachment), { recursive: true, mode: 0o700 });
  await writeFile(attachment, "Synthetic attachment BRONZE-TULIP-42", { mode: 0o600 });
  assert.equal(await locateSessionFile(id, { provider: "codebuddy" }), transcript);
  const children = [],
    stages = [],
    replayMessages = [];
  const observe = (session) =>
    observeCodeBuddyClose(session, {
      onProcess(child) {
        const item = { pid: child.pid, exited: false };
        children.push(item);
        process.stdout.write(JSON.stringify({ phase: "cli_child_started", pid: child.pid }) + "\n");
        child.once("exit", (code, signal) => Object.assign(item, { exited: true, code, signal }));
      },
    });
  const settings = {
    cwd,
    environment: "external",
    pathToCodebuddyCode: cli,
    executable: process.execPath,
    persistSession: false,
    settingSources: [],
    strictMcpConfig: true,
    mcpServers: {},
    tools: [],
    extraArgs: { port: selfPort, host: "127.0.0.1" },
    canUseTool: async () => ({ behavior: "deny", message: "Synthetic serial replay only" }),
  };
  const noSend = (session) => {
    session.send = () => {
      throw new Error("Forbidden: serial integration never sends a prompt");
    };
    return session;
  };
  let ready,
    live,
    active = null;
  const connected = new Promise((resolve) => {
    ready = resolve;
  });
  const handle = runSession(
    {
      cwd,
      resumeId: id,
      mcpServers: {},
      systemPromptAppend: "Synthetic serial handoff QA",
      turnSource: { next: () => new Promise(() => {}) },
      onMessage() {
        throw new Error("No model turn may run in the Word-side fixture");
      },
      onEnd() {
        live = null;
      },
    },
    {
      resume(sessionId, passed) {
        assert.equal(sessionId, id);
        assert.deepEqual(passed.settingSources, []);
        const session = noSend(unstable_v2_resumeSession(id, { ...passed, ...settings }));
        const connect = session.connect.bind(session);
        session.connect = async () => {
          await connect();
          ready();
        };
        return session;
      },
      create() {
        throw new Error("Serial handoff must never fork implicitly");
      },
      observeClose: observe,
    },
  );
  const registry = createRuntimeCloseRegistry();
  registry.track(key, handle.closed);
  live = { cancel: handle.cancel, done: handle.done, lifecycle: { active: false } };
  let cliReplay, cliObserver;
  try {
    await Promise.race([
      connected,
      handle.done.then(() => {
        throw new Error("Word-side runtime ended before connect");
      }),
    ]);
    stages.push("word_runtime_connected_without_send");
    const source = await readFile(new URL("../daemon/index.mjs", import.meta.url), "utf8");
    const infoCode = fragment(source, "async function sessionInfo(", "// Static fallback");
    const replayCode = fragment(
      source,
      "async function sendTranscriptReplayTo(",
      "// Start (or resume)",
    );
    const handlersCode = fragment(source, "    get_session_info: async", "    get_context: async");
    const ensureCode = fragment(
      source,
      "async function ensureLoopForMessage(",
      "\nconst bridge = createBridge(",
    );
    const context = vm.createContext({
      console,
      Date,
      Promise,
      setTimeout,
      clearTimeout,
      runtime: { id: "codebuddy" },
      runtimeClosures: registry,
      handoffInProgress: new Set(),
      startQueue: new Map(),
      bridge: {
        activeTurn: () => active,
        sendToTaskpane(message, target) {
          assert.equal(target, key);
          replayMessages.push(message);
        },
      },
      sessionFor: () => live,
      cwdForKey: () => cwd,
      resolveReplaySessionId: () => store.getSessionId("word", cwd, "codebuddy", key),
      getSessionId: store.getSessionId,
      getHandoff: store.getHandoff,
      setHandoff: store.setHandoff,
      resumeCommands,
      readTranscript,
      archiveSession: async (info) => {
        assert.equal(
          children[0].exited,
          true,
          "Archive started before actual Word runtime child exit",
        );
        stages.push("word_child_exited_before_archive");
        return archiveSession(info, { root: join(config, "archives") });
      },
      diag() {},
      scheduleSessionStart() {
        throw new Error("No model session start permitted by this integration");
      },
    });
    vm.runInContext(
      `${infoCode}\n${replayCode}\n${ensureCode}\nglobalThis.handlers = {${handlersCode}}; globalThis.ensure = ensureLoopForMessage;`,
      context,
    );
    const call = async (name, message = {}) => {
      let response;
      await context.handlers[name](
        { request_id: name, ...message },
        (r) => {
          response = r;
        },
        key,
        "word",
      );
      return response;
    };
    await call("release_session");
    const released = await call("get_session_info");
    assert.equal(released.info.state, "released");
    assert.ok(released.info.commands.resume.includes(`--resume '${id}'`));
    assert.ok(released.info.commands.resume.includes("CODEBUDDY_CONFIG_DIR="));
    active = "synthetic-blocked-turn";
    await assert.rejects(context.ensure(key, "word"), /released/);
    active = null;
    const reopened = createSessionStore(options);
    assert.equal((await reopened.getHandoff("word", cwd, "codebuddy", key)).state, "released");
    const archived = JSON.parse(await readFile(join(released.info.archive_path, "manifest.json")));
    assert.equal(archived.transcript_sha256, sha(before));
    assert.equal(archived.session_files.length, 1);
    assert.equal(
      sha(await readFile(join(released.info.archive_path, archived.session_files[0].path))),
      archived.session_files[0].sha256,
    );
    stages.push("release_archive_hash_attachment_and_reopen_pass");
    cliReplay = noSend(unstable_v2_resumeSession(id, settings));
    cliObserver = observe(cliReplay);
    const events = [];
    for await (const event of cliReplay.stream()) {
      events.push(event);
      assert.ok(events.length <= 30);
    }
    const observedIds = [...new Set(events.map((e) => e.session_id).filter(Boolean))];
    assert.deepEqual(observedIds, [id]);
    assert.ok(events.some((e) => e.type === "result" && !e.is_error));
    assert.ok(events.some((e) => e.message?.content?.some?.((b) => b.type === "tool_use")));
    assert.ok(events.some((e) => e.message?.content?.some?.((b) => b.type === "tool_result")));
    assert.ok(JSON.stringify(events).includes("BRONZE-TULIP-42"));
    cliObserver.close();
    await cliObserver.waitForExit();
    assert.equal(children.length, 2);
    assert.ok(children.every((c) => c.exited));
    stages.push("installed_cli_exact_id_history_replay_and_exit_pass");
    await assert.rejects(call("return_session", { session_id: id }), /exited/);
    await assert.rejects(
      call("return_session", { session_id: "wrong", cli_closed: true }),
      /identity/,
    );
    await call("return_session", { session_id: id, cli_closed: true });
    const replay = replayMessages.at(-1);
    assert.equal(replay.session_id, id);
    assert.ok(replay.events.some((e) => e.kind === "tool_result"));
    assert.ok(
      replay.events.some((e) => e.kind === "assistant" && e.text.includes("BRONZE-TULIP-42")),
    );
    assert.equal(replay.session_info.state, "word");
    assert.equal(await reopened.getHandoff("word", cwd, "codebuddy", key), null);
    assert.equal(await reopened.getSessionId("word", cwd, "codebuddy", key), id);
    assert.equal(
      await reopened.getSessionId("word", cwd, "codebuddy", "other-document"),
      "other-document-id",
    );
    assert.equal(await reopened.getSessionId("word", cwd, "claude", key), "other-provider-id");
    assert.equal(sha(await readFile(transcript)), sha(before));
    stages.push("production_return_replays_same_id_and_scoped_state_pass");
    return {
      phase: "serial_handoff",
      ok: true,
      sends: 0,
      session_id: id,
      observed_session_ids: observedIds,
      scope:
        "Synthetic production-handler/store/archive/runtime-close/installed-CLI/replay integration; no native pane, provider inference or CLI TUI",
      stages,
      children,
      archive_path: released.info.archive_path,
      source_sha256_before: sha(before),
      source_sha256_after: sha(await readFile(transcript)),
      source_unchanged: true,
      tool_call_seen: true,
      tool_result_seen: true,
      result_subtype: events.find((e) => e.type === "result")?.subtype,
    };
  } finally {
    handle.close();
    await handle.closed;
    if (cliObserver) {
      cliObserver.close();
      await cliObserver.waitForExit();
    } else cliReplay?.close();
  }
}
