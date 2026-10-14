import { targetOoxmlElements } from "./word-ooxml-state.js";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// Conservative admission checks. These do not rewrite OOXML or claim native
// format preservation. Unsupported rich targets require a narrower operation.
export function assertTextTargetSafe(raw, { broad = false, cell = false } = {}) {
  const elements = targetOoxmlElements(raw, { cell });
  if (!elements)
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: target OOXML could not be inspected.");
  const has = (names) =>
    elements.some((n) => n.namespace === W && names.split("|").includes(n.name));
  // Check the containing paragraph/story, not only the search hit: a field's
  // begin/instruction/end markers can sit outside its visible result range.
  if (has("fldSimple|fldChar|instrText"))
    throw new Error(
      "PROTECTED_FIELD: target contains a field, TOC or cross-reference. Choose text outside that paragraph/container.",
    );
  if (!broad) return;
  if (
    has(
      "commentRangeStart|commentRangeEnd|commentReference|ins|del|moveFrom|moveTo|sdt|drawing|pict|object|hyperlink|tbl|footnoteReference|endnoteReference|bookmarkStart|bookmarkEnd|altChunk|subDoc|customXml|smartTag",
    ) ||
    elements.some((n) => n.namespace === W && n.name === "sectPr" && n.parent === "pPr")
  )
    throw new Error(
      "UNSUPPORTED_RICH_TARGET: broad replacement may alter review or document structure; use a narrower range operation.",
    );
  if (elements.some((n) => n.namespace !== W))
    throw new Error(
      "UNSUPPORTED_RICH_TARGET: embedded or alternate-namespace content requires a narrower operation.",
    );
  // A plain replacement cannot promise mixed-run formatting inheritance. Even
  // one styled run amid unstyled text is rich, so refuse any explicit run style.
  if (elements.some((n) => n.namespace === W && n.name === "rPr" && n.hasChildren))
    throw new Error(
      "UNSUPPORTED_RICH_TARGET: explicit run formatting requires a narrower range operation.",
    );
}

export function assertNoCrossParagraphFields(raw, { protectReview = false } = {}) {
  const elements = targetOoxmlElements(raw);
  if (!elements)
    throw new Error("TARGET_INSPECTION_UNAVAILABLE: containing story OOXML is required.");
  let paragraph = 0;
  const fields = [];
  const comments = new Set();
  for (const n of elements) {
    if (n.namespace !== W) continue;
    if (n.name === "p") {
      paragraph++;
      if (protectReview && comments.size)
        throw new Error(
          "UNSUPPORTED_RICH_TARGET: a comment spans paragraphs in this story; exact overlap is not established, so broad replacement is refused.",
        );
      if (fields.length)
        throw new Error(
          "PROTECTED_FIELD: story contains a multi-paragraph field (such as a TOC); exact overlap is not established, so text replacement is refused.",
        );
    }
    if (n.name === "fldChar" && n.fieldType === "begin") fields.push(paragraph);
    if (protectReview && n.name === "commentRangeStart") comments.add(n.referenceId);
    if (protectReview && n.name === "commentRangeEnd") {
      if (!comments.delete(n.referenceId))
        throw new Error("UNSUPPORTED_RICH_TARGET: incomplete comment boundaries in target story.");
    }
    if (n.name === "fldChar" && n.fieldType === "end") {
      if (!fields.length)
        throw new Error("PROTECTED_FIELD: incomplete field boundaries in target story.");
      fields.pop();
    }
  }
  if (fields.length)
    throw new Error("PROTECTED_FIELD: incomplete field boundaries in target story.");
  if (protectReview && comments.size)
    throw new Error("UNSUPPORTED_RICH_TARGET: incomplete comment boundaries in target story.");
}

export function assertDistinctIds(ids, label = "ids") {
  if (
    !Array.isArray(ids) ||
    !ids.length ||
    ids.some((id) => typeof id !== "string" || !id) ||
    new Set(ids).size !== ids.length
  )
    throw new Error(`INVALID_SELECTOR: ${label} must contain distinct nonempty paragraph IDs.`);
}

export function assertOneSelector(values, { optional = false } = {}) {
  const count = values.filter((v) => v !== undefined).length;
  if (count !== 1 && !(optional && count === 0))
    throw new Error(
      "INVALID_SELECTOR: choose exactly one paragraph ID, heading, range or all selector.",
    );
}

// Word search treats "^" as the start of a special code (^p, ^t, ^#, ^?) even
// with wildcards off, and rejects search strings over 255 characters.
export function assertLiteralSearchText(text, label = "find") {
  if (typeof text !== "string" || !text) throw new Error(`\`${label}\` must be a nonempty string.`);
  if (text.length > 255)
    throw new Error(
      `SEARCH_TEXT_TOO_LONG: Word search accepts at most 255 characters; use a shorter unique \`${label}\`.`,
    );
  if (text.includes("^"))
    throw new Error(
      `UNSUPPORTED_SEARCH_TEXT: "^" starts a Word special-character code (^p, ^t, ^#); choose a \`${label}\` without it.`,
    );
}
