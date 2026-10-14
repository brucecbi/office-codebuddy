// Model metadata only. Close means an observed child exit, including during shutdown.
import { createRuntimeCloseRegistry } from "./close-registry.mjs";

export function createModelDiscovery({ create, observeClose, options, now = Date.now }) {
  const children = createRuntimeCloseRegistry();
  let cache = null,
    inflight = null,
    stopping = false;
  return {
    async list({ cwd, timeoutMs = 25000, force = false } = {}) {
      if (stopping) return null;
      if (!force && cache && now() - cache.at < 600000) return cache.list;
      if (inflight) return inflight;
      inflight = (async () => {
        const session = create({
          ...options(),
          cwd,
          settingSources: [],
          hooks: {},
          mcpServers: {},
          tools: [],
        });
        let observer,
          timer,
          closeRequested = false,
          resolveExit,
          rejectExit;
        const closed = new Promise((resolve, reject) => {
          resolveExit = resolve;
          rejectExit = reject;
        });
        closed.catch(() => {});
        const close = () => {
          if (closeRequested) return;
          closeRequested = true;
          try {
            if (!observer) throw new Error("Model discovery process exit is unobservable");
            observer.close();
            observer.waitForExit().then(resolveExit, rejectExit);
          } catch (error) {
            rejectExit(error);
          }
        };
        children.track("catalog", closed, { close });
        try {
          observer = observeClose(session);
          const list = await Promise.race([
            (async () => {
              await session.connect?.();
              return session.getAvailableModels();
            })(),
            new Promise((_, reject) => {
              timer = setTimeout(() => reject(new Error("Model discovery timeout")), timeoutMs);
              timer.unref?.();
            }),
          ]);
          if (!stopping && Array.isArray(list) && list.length) {
            cache = { at: now(), list };
            return list;
          }
          return null;
        } catch {
          return null;
        } finally {
          clearTimeout(timer);
          close();
          // A wedged child remains in the registry; shutdown must report it.
          if (!observer) {
            try {
              session.close();
            } catch {}
          }
        }
      })();
      try {
        return await inflight;
      } finally {
        inflight = null;
      }
    },
    async shutdown(timeoutMs = 2500) {
      stopping = true;
      await children.closeAll(timeoutMs);
    },
  };
}
