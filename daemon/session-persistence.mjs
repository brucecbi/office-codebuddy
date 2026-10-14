// Persistence barriers outlive the live SDK session. Stop/stream failure may
// remove that session while its latest resume-ID write is still pending or has
// failed; archiving must not silently fall back to an older on-disk ID.
export function createSessionPersistenceRegistry({ save, notify = () => {} }) {
  const records = new Map(); // One current record per pane, not an event history.
  const scopeMatches = (record, scope) =>
    record &&
    record.identity.cwd === scope.cwd &&
    record.identity.provider === scope.provider &&
    (scope.host === undefined || record.identity.host === scope.host);
  const sameIdentity = (a, b) =>
    a.host === b.host &&
    a.cwd === b.cwd &&
    a.provider === b.provider &&
    a.session_id === b.session_id;
  const scopeError = () =>
    Object.assign(
      new Error("The session persistence identity changed; refresh before retrying or archiving."),
      { code: "SESSION_PERSISTENCE_SCOPE_CHANGED" },
    );
  const emit = (event, key, record) => {
    if (records.get(key) !== record) return;
    // Delivery failures do not undo a successful disk write or change the
    // persistence outcome. Status RPC remains the authoritative fallback.
    try {
      notify(event, key);
    } catch {}
  };
  function begin(key, identity) {
    if (
      typeof key !== "string" ||
      !key ||
      !["word", "excel"].includes(identity?.host) ||
      !["cwd", "provider", "session_id"].every(
        (name) => typeof identity[name] === "string" && identity[name],
      )
    )
      throw Object.assign(
        new Error("A complete document-scoped persistence identity is required."),
        { code: "SESSION_PERSISTENCE_IDENTITY_INVALID" },
      );
    const previous = records.get(key);
    const record = { identity: { ...identity }, status: "saving", error: null, promise: null };
    records.set(key, record);
    // Serialize even different identities for this pane: an earlier delayed
    // write must never land after the newer ID and overwrite it on disk.
    record.promise = (previous?.promise || Promise.resolve())
      .catch(() => {})
      .then(() =>
        save(
          record.identity.host,
          record.identity.cwd,
          record.identity.session_id,
          record.identity.provider,
          key,
        ),
      )
      .then(
        () => {
          record.status = "saved";
          emit({ event: "persistence_saved" }, key, record);
        },
        (error) => {
          record.status = "failed";
          record.error =
            typeof error?.code === "string" && /^[A-Z][A-Z0-9_]{0,63}$/.test(error.code)
              ? error.code
              : "SESSION_PERSISTENCE_FAILED";
          emit({ event: "persistence_error", code: record.error }, key, record);
          throw error;
        },
      );
    record.promise.catch(() => {}); // Observe, do not convert rejection to success.
    return record.promise;
  }
  return {
    persist(key, identity) {
      const current = records.get(key);
      return current && sameIdentity(current.identity, identity)
        ? current.promise
        : begin(key, identity);
    },
    status(key, scope) {
      const record = records.get(key);
      return scopeMatches(record, scope)
        ? {
            status: record.status,
            error: record.error,
            session_id: record.identity.session_id,
          }
        : null;
    },
    async awaitSaved(key, scope) {
      let current = records.get(key);
      if (!scopeMatches(current, scope)) return null;
      for (;;) {
        // A replacement attempt may supersede a rejected old promise; the
        // latest attempt is authoritative, but still cannot pass while saving.
        let failure;
        try {
          await current.promise;
        } catch (error) {
          failure = error;
        }
        const latest = records.get(key);
        if (!scopeMatches(latest, scope)) throw scopeError();
        if (latest !== current) {
          current = latest;
          continue;
        }
        if (current.status !== "saved") throw failure || new Error("Session persistence failed");
        return { session_id: current.identity.session_id };
      }
    },
    retry(key, scope) {
      const record = records.get(key);
      if (!scopeMatches(record, scope)) throw scopeError();
      if (record.status === "saving") return record.promise;
      return begin(key, record.identity);
    },
  };
}
