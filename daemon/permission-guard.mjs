// Permission handler. Hard rule: filesystem tools must never WRITE to .docx
// files. The .docx the user has open in Word is held with unsaved changes; a
// filesystem write clobbers their work and can corrupt the file (Word holds a
// lock, OOXML cross-file references can break). Reading is allowed — the agent
// occasionally falls back to unzipping .docx XML to extract text, which is
// safe and read-only.

// Path/extension test for the Office-managed file types.
export const OFFICE_FILE_EXT = /\.(docx?|xlsx?|docm|xlsm)\b/i;

// Bash is allowed to *read* Office files (the agent legitimately does
// `unzip -p draft.docx word/document.xml`, `cat`, `git log -- report.docx`,
// `ls *.docx`). We only deny commands that mutate one in place: a shell
// redirection whose target is an Office path, or a destructive verb used
// against one. This is an accident guard, not a security boundary — it
// stops the obvious foot-gun (`cp draft.docx active.docx`,
// `rm -rf workspace/*.docx`, `echo x > active.xlsx`, `sed -i … report.docx`,
// `zip -u draft.docx word/document.xml`, `del report.docx`), not a
// determined evasion (base64 payloads, here-docs, etc.).
// Quoted targets may contain spaces (`> "/tmp/My Draft/file.docx"`). An
// optional opening quote plus a token that stops at whitespace misses those.
function redirectsToOfficeFile(cmd) {
  const re = />>?\s*(?:"([^"]*)"|'([^']*)'|([^\s|;&]+))/gi;
  for (const match of cmd.matchAll(re)) {
    const path = match[1] ?? match[2] ?? match[3] ?? "";
    if (OFFICE_FILE_EXT.test(path)) return true;
  }
  return false;
}
const DESTRUCTIVE_VERB =
  /\b(rm|mv|cp|dd|shred|truncate|install|tee|ln|zip|del|erase|move|copy|ren|remove-item|move-item|copy-item|rename-item|set-content|add-content|clear-content|out-file)\b/i;
const SED_IN_PLACE = /\bsed\b[^|;&]*\s-i\b/i;

export function bashMutatesOfficeFile(cmd) {
  if (redirectsToOfficeFile(cmd)) return true;
  if (!OFFICE_FILE_EXT.test(cmd)) return false;
  return DESTRUCTIVE_VERB.test(cmd) || SED_IN_PLACE.test(cmd);
}

function denyWithOfficeMessage() {
  return {
    behavior: "deny",
    message:
      "Refusing to write/move/delete a Word/Excel file via filesystem tools. These files are " +
      "managed by Office and may have unsaved changes; a filesystem mutation can corrupt the " +
      "active document. Use the host's editing tools (office_* for Word, excel_* for Excel) " +
      "to change document contents.",
  };
}

export function customPermissionHandler(toolName, input) {
  if (toolName === "Write" || toolName === "Edit" || toolName === "MultiEdit") {
    const path = input?.file_path ?? input?.path;
    if (typeof path === "string" && OFFICE_FILE_EXT.test(path)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  if (toolName === "Bash") {
    const cmd = input?.command;
    if (typeof cmd === "string" && bashMutatesOfficeFile(cmd)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  // Everything else: auto-approve.
  return Promise.resolve({ behavior: "allow", updatedInput: input ?? {} });
}
