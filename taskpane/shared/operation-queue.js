// Serialize Office operations; pending duplicate IDs share one outcome. Persist a minimal
// receipt before writes so a reload cannot automatically repeat an uncertain edit.
export function createOperationQueue({ storage = null, prefix = "draftspect-operation:" } = {}) {
  let tail = Promise.resolve();
  let generation = 0;
  const records = new Map();
  return {
    cancel() {
      generation++;
    },
    run(id, { mutates, deadline, valid = () => true }, action) {
      if (records.has(id)) return records.get(id);
      const epoch = generation;
      // `deadline` may be a function so a confirmation can extend it.
      const expiresAt = () => (typeof deadline === "function" ? deadline() : deadline);
      const guard = () => {
        if (epoch !== generation || Date.now() >= expiresAt() || !valid())
          throw new Error("CANCELLED_OR_EXPIRED: this operation is no longer authorized.");
      };
      const promise = tail
        .catch(() => {})
        .then(async () => {
          guard();
          const key = prefix + id;
          if (mutates && storage?.getItem(key))
            throw new Error(
              "RESULT_UNKNOWN: this write ID was already attempted. Inspect the document; do not replay it.",
            );
          let marked = false;
          const beforeWrite = () => {
            guard();
            if (mutates && !marked) {
              storage?.setItem(key, "attempted");
              marked = true;
            }
          };
          const result = await action({ guard, beforeWrite });
          // A cancel that arrives after preparation, including during the final
          // Office sync, must not be reported as success. Batches already sent
          // cannot be recalled; the caller treats that as an uncertain result
          // once beforeWrite has run.
          guard();
          if (marked) storage?.setItem(key, "completed");
          return result;
        })
        .finally(() => records.delete(id));
      records.set(id, promise);
      tail = promise;
      return promise;
    },
  };
}
