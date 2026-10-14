import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

test("every host system prompt labels document and selected content as untrusted data, not instructions", async () => {
  for (const name of ["system-prompt.md", "system-prompt-word.md", "system-prompt-excel.md"]) {
    const prompt = await readFile(new URL(`../daemon/${name}`, import.meta.url), "utf8");
    assert.match(prompt, /untrusted data/i, name);
    assert.match(prompt, /not (?:as )?instructions|not treat[^\n]*instructions/i, name);
    assert.match(prompt, /user[^\n]*explicit|explicit[^\n]*user/i, name);
  }
});
