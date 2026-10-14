// v4: installation-local, provider/host/document-scoped resume IDs.
// Legacy folder IDs are retained for recovery, never assigned to a document.
import { readFile, writeFile, mkdir, open, rename, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";
import { randomUUID, createHash } from "node:crypto";
import { isSystemHomeChild } from "./system-paths.mjs";
import { stateFile } from "./state-paths.mjs";
import { acceptanceProfile } from "./acceptance-profile.mjs";
import { acquireSessionLock } from "./session-lock.mjs";

const VERSION = 4;
const providerKey = (p) =>
  String(p || "claude")
    .trim()
    .toLowerCase();
const validHost = (h) => h === "word" || h === "excel";
const object = (v) => v !== null && typeof v === "object" && !Array.isArray(v);
const documentKey = (key) => createHash("sha256").update(String(key)).digest("hex");

export function migrateState(parsed) {
  if (!object(parsed) || !object(parsed.folders))
    throw new Error("Invalid session store; refusing to overwrite it");
  if (![1, 2, 3, 4].includes(parsed.version))
    throw new Error("Unsupported session store version: " + parsed.version);
  if (parsed.version === VERSION) {
    for (const info of Object.values(parsed.folders)) {
      if (!object(info) || !object(info.documents) || !object(info.legacySessions))
        throw new Error("Invalid v4 session folder; refusing to overwrite it");
    }
    return parsed;
  }
  const folders = {};
  for (const [cwd, info] of Object.entries(parsed.folders)) {
    if (!object(info)) throw new Error("Invalid session folder record");
    const legacySessions =
      parsed.version === 2
        ? { claude: info.sessions ?? {} }
        : parsed.version === 3
          ? (info.sessions ?? {})
          : {};
    folders[cwd] = {
      last_used: info.last_used ?? new Date(0).toISOString(),
      display_name: info.display_name ?? basename(cwd),
      legacySessions,
      ...(info.session_id ? { legacy_unattributed_id: info.session_id } : {}),
      documents: {},
    };
  }
  return { version: VERSION, folders };
}

export function createSessionStore({
  file = stateFile("sessions.json"),
  legacyFile = null, // Importing history is always an explicit caller decision.
  lockTimeoutMs = 2000,
  lockOptions = {},
} = {}) {
  let queue = Promise.resolve();
  async function readState() {
    let raw;
    try {
      raw = await readFile(file, "utf8");
    } catch (error) {
      if (error.code !== "ENOENT") throw error;
      if (legacyFile) {
        try {
          raw = await readFile(legacyFile, "utf8");
        } catch (e) {
          if (e.code !== "ENOENT") throw e;
        }
      }
    }
    return raw === undefined
      ? { state: { version: VERSION, folders: {} }, raw: null }
      : { state: migrateState(JSON.parse(raw)), raw };
  }
  async function transaction(update) {
    const run = queue
      .catch(() => {})
      .then(async () => {
        await mkdir(dirname(file), { recursive: true });
        const lockPath = file + ".lock";
        const lock = await acquireSessionLock(lockPath, {
          timeoutMs: lockTimeoutMs,
          ...lockOptions,
        });
        const temp = file + ".tmp-" + randomUUID();
        try {
          const { state, raw } = await readState();
          if (raw && JSON.parse(raw).version !== VERSION) {
            const digest = createHash("sha256").update(raw).digest("hex").slice(0, 16);
            try {
              await writeFile(file + ".pre-v4-" + digest + ".bak", raw, {
                flag: "wx",
                mode: 0o600,
              });
            } catch (e) {
              if (e.code !== "EEXIST") throw e;
            }
          }
          update(state);
          const out = await open(temp, "wx", 0o600);
          try {
            await out.writeFile(JSON.stringify(state, null, 2) + "\n");
            await out.sync();
          } finally {
            await out.close();
          }
          // A replaced/tampered lock must not authorize committing state.
          await lock.assertOwned();
          await rename(temp, file);
        } finally {
          try {
            await unlink(temp);
          } catch (e) {
            if (e.code !== "ENOENT") throw e;
          } finally {
            await lock.release();
          }
        }
      });
    queue = run;
    return run;
  }
  function folder(state, cwd) {
    return (state.folders[cwd] ??= {
      last_used: "",
      display_name: basename(cwd),
      legacySessions: {},
      documents: {},
    });
  }
  return {
    async getSessionId(host, cwd, provider = "claude", doc = null) {
      if (!validHost(host)) return null;
      await queue.catch(() => {});
      const { state } = await readState();
      const f = state.folders[cwd];
      // No-doc lookups are for legacy inspection. Daemon always supplies doc.
      const group = doc ? f?.documents?.[documentKey(doc)] : f?.legacySessions;
      return group?.[providerKey(provider)]?.[host] ?? null;
    },
    async saveSessionId(host, cwd, id, provider = "claude", doc = null) {
      if (!validHost(host)) return;
      return transaction((state) => {
        const f = folder(state, cwd);
        const group = doc ? (f.documents[documentKey(doc)] ??= {}) : f.legacySessions;
        (group[providerKey(provider)] ??= {})[host] = id;
        f.last_used = new Date().toISOString();
      });
    },
    async getHandoff(host, cwd, provider, doc) {
      await queue.catch(() => {});
      const { state } = await readState();
      return (
        state.folders[cwd]?.documents?.[documentKey(doc)]?._handoffs?.[providerKey(provider)]?.[
          host
        ] ?? null
      );
    },
    async setHandoff(host, cwd, provider, doc, value) {
      if (!validHost(host) || !doc) throw new Error("A document-scoped handoff is required");
      return transaction((state) => {
        const group = (folder(state, cwd).documents[documentKey(doc)] ??= {});
        ((group._handoffs ??= {})[providerKey(provider)] ??= {})[host] = value;
      });
    },
    async touchFolder(cwd) {
      if (isSystemHomeChild(cwd)) return;
      return transaction((state) => {
        folder(state, cwd).last_used = new Date().toISOString();
      });
    },
  };
}
const store = createSessionStore();
export const { getSessionId, saveSessionId, touchFolder, getHandoff, setHandoff } = store;
