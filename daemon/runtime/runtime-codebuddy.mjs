// CodeBuddy runtime — drives @tencent-ai/agent-sdk's V2 Session
// (unstable_v2_createSession / unstable_v2_resumeSession). This is the
// default runtime. Select the Claude transport with DRAFTSPECT_RUNTIME=claude.
//
// Behavior this implementation relies on:
//   - The account environment must match the user's own CLI login.
//   - pathToCodebuddyCode must be a REALPATH: the SDK's process transport
//     rewrites a path ending in /bin/codebuddy to ../dist/codebuddy-
//     headless.js WITHOUT resolving symlinks first (npm global bins are
//     symlinks), yielding a non-existent file.
//   - turn shape: session.send(text) then iterate session.stream() until
//     a `result`; a fresh stream() per turn works (long-lived session).
//   - interrupt() may end with a NORMAL success result. The current daemon
//     fences Stop immediately, closes the session and lazily resumes on the
//     next user message; it never trusts that result as a completed answer.
//   - init carries session_id; result carries subtype/errors[].
//   - No stderr callback: usage-limit detection must use the structured
//     signals (Phase 3 of the migration plan, Q9 decision).
//   - systemPrompt has no preset form — the Office append is passed as
//     { append } on top of the CodeBuddy default prompt (texts must be
//     re-constrained for that default; Phase 5).

import {
  unstable_v2_createSession,
  unstable_v2_resumeSession,
  createSdkMcpServer,
  tool,
} from "@tencent-ai/agent-sdk";
import { realpathSync, statSync } from "node:fs";
import { join, resolve } from "node:path";
import { observeCodeBuddyClose } from "./codebuddy-process-close.mjs";
import { createModelDiscovery } from "./model-discovery.mjs";
import { acceptanceProfile, acceptanceSdkOptions } from "../acceptance-profile.mjs";
import { codebuddyEnvironment, previewSdkOptions } from "../preview-policy.mjs";

export const id = "codebuddy";
export const supportsStderr = false;

// Use this runtime's own MCP primitives. Cross-SDK instance identity alone
// has not been established as the cause of historical provider 400 errors.
export const mcp = { createSdkMcpServer, tool };

// Discovery-based catalog — the daemon fetches it via listModels() and
// caches; there is no static list to validate against synchronously.
export const knownModels = null;

// Public-preview defaults and explicit extension opt-ins live in preview-policy.mjs.

// Resolve once in the daemon's cwd, matching transcript lookup/resume commands.
// Otherwise a relative config override would resolve differently in each
// document's CLI child cwd. This only overrides that child's matching env key.
export function codebuddyConfigEnvironment(env = process.env) {
  const config = env.CODEBUDDY_CONFIG_DIR?.trim();
  return config ? { env: { CODEBUDDY_CONFIG_DIR: resolve(config) } } : {};
}

// Resolve the CodeBuddy CLI to a REAL path:
//   1. CODEBUDDY_CODE_PATH (explicit override)
//   2. `codebuddy` on PATH (the user's global install, e.g. an nvm bin)
//   3. null → let the SDK use its bundled CLI
// Always realpath() before handing the SDK a path — see header note.
export function resolveCliPath() {
  const explicit = (process.env.CODEBUDDY_CODE_PATH || "").trim();
  if (explicit) return safeRealpath(explicit);
  for (const dir of (process.env.PATH || "").split(":").filter(Boolean)) {
    const candidate = join(dir, "codebuddy");
    try {
      const st = statSync(candidate);
      if (st.isFile() || st.isSymbolicLink()) return realpathSync(candidate);
    } catch {
      /* keep looking */
    }
  }
  return null;
}

function safeRealpath(p) {
  try {
    return realpathSync(p);
  } catch {
    return p;
  }
}

async function withTimeout(promise, ms, label) {
  let timer;
  try {
    return await Promise.race([
      promise,
      new Promise((_, reject) => {
        timer = setTimeout(() => reject(new Error(`TIMEOUT ${label} ${ms}ms`)), ms);
        timer.unref?.();
      }),
    ]);
  } finally {
    clearTimeout(timer);
  }
}

export async function probe() {
  const cli = resolveCliPath();
  try {
    codebuddyEnvironment();
  } catch (error) {
    return { ok: false, detail: error.message };
  }
  return {
    ok: true,
    detail: `environment=${codebuddyEnvironment()} cli=${cli ?? "(sdk bundled)"}`,
  };
}

// Dynamic model discovery. Works on CLI ≥2.146 (verified 44 models); on
// older CLIs or auth problems this resolves to null. Spawns a throwaway
// session — cached for 10 minutes with in-flight dedup.
const modelDiscovery = createModelDiscovery({
  create: unstable_v2_createSession,
  observeClose: observeCodeBuddyClose,
  options: () => ({
    environment: codebuddyEnvironment(),
    ...codebuddyConfigEnvironment(),
    ...(resolveCliPath() ? { pathToCodebuddyCode: resolveCliPath() } : {}),
  }),
});
export const shutdown = (timeoutMs) => modelDiscovery.shutdown(timeoutMs);

export async function listModels({ cwd = process.cwd(), timeoutMs = 25000, force = false } = {}) {
  if (acceptanceProfile) {
    // This is a test pin, not provider discovery evidence. No discovery child,
    // user hooks, MCP servers or extra request can be started through this path.
    console.log("[acceptance] pinned model catalog override; provider discovery NOT_RUN");
    return [
      {
        value: acceptanceProfile.model,
        label: `Acceptance pinned: ${acceptanceProfile.model} (not discovered)`,
      },
    ];
  }
  return modelDiscovery.list({ cwd, timeoutMs, force });
}

export function runSession(
  {
    cwd,
    resumeId = null,
    model,
    systemPromptAppend,
    mcpServers,
    disallowedTools,
    canUseTool,
    includePartialMessages = true,
    turnSource,
    isStale = () => false,
    onMessage,
    onTurnEnd,
    onEnd,
  },
  {
    create = unstable_v2_createSession,
    resume = unstable_v2_resumeSession,
    connectTimeoutMs = 30000,
    observeClose = observeCodeBuddyClose,
  } = {},
) {
  const cli = resolveCliPath();
  const makeOptions = () =>
    acceptanceSdkOptions(
      previewSdkOptions({
        cwd,
        environment: codebuddyEnvironment(),
        ...codebuddyConfigEnvironment(),
        ...(cli ? { pathToCodebuddyCode: cli } : {}),
        mcpServers,
        disallowedTools,
        canUseTool,
        systemPrompt: { append: systemPromptAppend },
        includePartialMessages,
        ...(model ? { model } : {}),
      }),
    );
  let session = null;
  let observer = null;
  let closeError = null;
  let closed = false;
  let resourcesClosed = false;
  const closeWaiters = new Set();
  // Remove the cancellation waiter after every settled read; a long-lived
  // pending Promise.race cancellation promise would retain every text delta.
  const untilClosed = (promise) =>
    new Promise((resolve, reject) => {
      if (closed) {
        resolve(null);
        return;
      }
      const cancel = () => resolve(null);
      closeWaiters.add(cancel);
      Promise.resolve(promise).then(
        (value) => {
          closeWaiters.delete(cancel);
          resolve(value);
        },
        (error) => {
          closeWaiters.delete(cancel);
          reject(error);
        },
      );
    });
  const release = () => {
    if (!session || resourcesClosed) return;
    resourcesClosed = true;
    try {
      if (observer) observer.close();
      else session.close();
    } catch (error) {
      closeError = error;
    }
  };
  const close = () => {
    if (closed) return;
    closed = true;
    for (const cancel of closeWaiters) cancel();
    closeWaiters.clear();
    try {
      Promise.resolve(session?.interrupt?.()).catch(() => {});
    } catch {}
    release();
  };
  const done = (async () => {
    let ending = { reason: "ended" };
    try {
      // Resume/lock failures are explicit; never silently replace a history.
      console.log(
        `[agent] codebuddy session process starting (${resumeId ? `resume ${resumeId}` : "new session"})`,
      );
      const options = makeOptions();
      session = resumeId ? resume(resumeId, options) : create(options);
      observer = observeClose(session);
      await untilClosed(
        withTimeout(Promise.resolve(session.connect?.()), connectTimeoutMs, "connect"),
      );
      if (closed) return;
      while (!closed && !isStale()) {
        const content = await untilClosed(turnSource.next());
        if (content === null || closed || isStale()) break;
        await untilClosed(session.send(content));
        if (closed) break;
        let sawResult = false;
        const iterator = session.stream()[Symbol.asyncIterator]();
        try {
          while (!closed) {
            const next = await untilClosed(iterator.next());
            if (!next || next.done || closed || isStale()) break;
            onMessage(next.value);
            if (next.value.type === "result") {
              sawResult = true;
              break;
            }
          }
        } finally {
          const returning = Promise.resolve(iterator.return?.());
          if (closed) returning.catch(() => {});
          else await untilClosed(withTimeout(returning, 5000, "close stream"));
        }
        if (!sawResult && !closed && !isStale()) {
          onTurnEnd?.({ sawResult: false });
          break;
        }
      }
    } catch (error) {
      ending = { reason: "error", error };
    } finally {
      release();
      try {
        if (observer) await observer.waitForExit();
        else if (session) throw new Error("CodeBuddy process exit could not be observed");
      } catch (error) {
        closeError = error;
      }
      onEnd(closed ? { reason: "aborted" } : ending);
    }
  })();
  const exited = done.then(() => {
    if (closeError) throw closeError;
  });
  exited.catch(() => {});
  return {
    cancel: close,
    close,
    done,
    closed: exited,
    getCloseStatus: () =>
      observer?.getCloseStatus?.() ?? {
        phase: "starting",
        close_requested: closed,
        can_terminate: false,
      },
    terminateOwnedProcess: (options) => {
      if (!observer?.terminateOwnedProcess)
        throw new Error("Runtime process is not observable; cannot terminate safely.");
      return observer.terminateOwnedProcess(options);
    },
  };
}
