// Model-catalog helpers shared by the daemon and its runtimes. Pure —
// unit-tested in tests/models.test.mjs.
//
// Two catalog shapes exist:
//   - Claude runtime: standard aliases — runtime-claude.STATIC_MODELS.
//   - CodeBuddy runtime: a DISCOVERED account-specific list
//     (getAvailableModels()), fetched and cached by the daemon.
//
// The pane's stored model may be stale after a runtime switch (legacy provider ids
// are meaningless to CodeBuddy and vice versa). pickModel() resolves that
// explicitly: a stale id is DROPPED (runtime default applies) and the
// caller is told, instead of silently sending an id the model layer would
// reject.

/** Normalize a raw SDK model list into [{value, label?}] (null when empty). */
export function normalizeModelList(list) {
  if (!Array.isArray(list)) return null;
  const out = [];
  for (const m of list) {
    if (typeof m === "string" && m.trim()) {
      out.push({ value: m.trim() });
      continue;
    }
    if (m && typeof m === "object") {
      const value = String(m.value ?? m.modelId ?? m.id ?? "").trim();
      if (!value) continue;
      const label = m.label ?? m.name ?? m.displayName ?? m.shortName;
      out.push(label ? { value, label: String(label) } : { value });
    }
  }
  return out.length ? out : null;
}

/**
 * Resolve which model id to use for a session turn.
 *
 * @param {object} args
 * @param {string|null|undefined} args.requested  pane's stored choice
 * @param {string} args.runtimeId                 "claude" | "codebuddy" | …
 * @param {string[]|null} args.knownModels        static catalog (claude)
 * @param {{value:string}[]|null} args.dynamicModels  fetched catalog (others)
 * @param {string|null} args.fallback             static fallback (claude)
 * @returns {{model: string|undefined, dropped: boolean}}
 *   dropped=true → the requested id is not in this runtime's catalog and was
 *   replaced by the runtime default (undefined = let the runtime/CLI decide).
 */
export function pickModel({
  requested,
  runtimeId,
  knownModels = null,
  dynamicModels = null,
  fallback = null,
}) {
  const req = typeof requested === "string" && requested.trim() ? requested.trim() : null;
  if (runtimeId === "claude") {
    if (req && Array.isArray(knownModels) && knownModels.includes(req)) {
      return { model: req, dropped: false };
    }
    if (req) return { model: fallback, dropped: true };
    return { model: fallback, dropped: false };
  }
  // Dynamic-catalog runtimes (CodeBuddy): no choice → runtime default; a
  // choice must be validated against a current catalog before use.
  if (!req) return { model: undefined, dropped: false };
  if (!Array.isArray(dynamicModels) || dynamicModels.length === 0) {
    return { model: undefined, dropped: true };
  }
  if (dynamicModels.some((m) => m && m.value === req)) {
    return { model: req, dropped: false };
  }
  return { model: undefined, dropped: true };
}
