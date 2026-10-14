// Pure precondition/preview logic, shared with synthetic regression fixtures.
export function planReplacement(snapshot, args) {
  const ids = args.paragraph_ids;
  if (snapshot.addressing !== "uniqueLocalId")
    throw new Error("Stable paragraph IDs are required for a controlled replacement.");
  if (!Array.isArray(ids) || !ids.length || new Set(ids).size !== ids.length)
    throw new Error("Choose distinct target paragraph IDs.");
  if (typeof args.find !== "string" || !args.find || typeof args.replace !== "string")
    throw new Error("A nonempty find string and replacement string are required.");
  if (/[\r\n]/.test(args.find + args.replace))
    throw new Error(
      "Use paragraph tools for structural edits; this path replaces text within paragraphs.",
    );
  return ids.map((id) => {
    const p = snapshot.paragraphs.find((p) => p.id === id);
    if (!p) throw new Error(`Paragraph not found: ${id}`);
    const escaped = args.find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (args.whole_word && !Array.isArray(p.search_matches))
      throw new Error("WORD_SEARCH_REQUIRED: whole-word boundaries must be resolved by Word.");
    const matches = p.search_matches
      ? p.search_matches.map(({ text, index }) => Object.assign([text], { index }))
      : [...p.text.matchAll(new RegExp(escaped, args.match_case ? "gu" : "giu"))];
    if (matches.length !== 1)
      throw new Error(
        `Expected one unambiguous match in ${id}; found ${matches.length}. Narrow the find text.`,
      );
    const m = matches[0];
    if (
      !Number.isInteger(m.index) ||
      m.index < 0 ||
      p.text.slice(m.index, m.index + m[0].length) !== m[0]
    )
      throw new Error("AMBIGUOUS_RANGE: Word match does not address Current text.");
    return {
      id,
      before: p.text,
      after: p.text.slice(0, m.index) + args.replace + p.text.slice(m.index + m[0].length),
      match: m[0],
      match_offset: m.index,
    };
  });
}

export function assertPreconditions(snapshot, plan) {
  if (snapshot.addressing !== "uniqueLocalId")
    throw new Error("Stable IDs are no longer available.");
  for (const p of plan) {
    if (snapshot.paragraphs.find((v) => v.id === p.id)?.text !== p.before) {
      throw new Error(
        `STALE_PROPOSAL: paragraph ${p.id} changed after preview. Read and propose again.`,
      );
    }
  }
}
