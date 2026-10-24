import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { z } from "zod";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";
import { mcp as claude } from "../daemon/runtime/runtime-claude.mjs";
import { mcp as codebuddy } from "../daemon/runtime/runtime-codebuddy.mjs";
test("Office registration requires explicit provider primitives instead of loading a fallback SDK", () => {
  assert.throws(() => createOfficeBridgeMcp({}, "word", "pane"), /selected runtime/);
});

for (const [provider, sdk] of [
  ["claude", claude],
  ["codebuddy", codebuddy],
]) {
  test(`${provider} registers complete Word schemas without tuple-items and routes to one pane`, async () => {
    const calls = [];
    const server = createOfficeBridgeMcp(
      {
        callTaskpaneTool: async (...args) => {
          calls.push(args);
          return { fixture: true };
        },
      },
      "word",
      "pane-A",
      sdk,
    );
    const tools = server.instance._registeredTools;
    assert.equal(Object.keys(tools).length, 35);
    const walk = (value) => {
      if (!value || typeof value !== "object") return;
      if (value.items)
        assert.equal(Array.isArray(value.items), false, "Provider-incompatible tuple-style items");
      assert.ok(!value.prefixItems, "Provider-incompatible prefixItems");
      Object.values(value).forEach(walk);
    };
    for (const t of Object.values(tools))
      walk(z.toJSONSchema(t.inputSchema, { target: "draft-7" }));
    for (const name of [
      "office_document_map",
      "office_read_story",
      "office_read_formatting",
      "office_list_comments",
      "office_edit_range",
      "office_manage_comment",
      "office_read_revisions",
      "office_audit_redline",
      "office_manage_revisions",
    ]) {
      assert.ok(tools[name]);
      const result = await tools[name].handler({}, {});
      assert.deepEqual(calls.at(-1), [name, {}, "pane-A"]);
      assert.equal(result.isError, undefined);
    }
    assert.equal(
      tools.office_edit_range.inputSchema.safeParse({
        target: {},
        action: "format",
        font: { bold: "false" },
      }).success,
      false,
    );
  });
}
test("Word manifest has Hosts before Requirements as required by the task-pane schema", async () => {
  const xml = await readFile(new URL("../manifests/word.xml", import.meta.url), "utf8");
  assert.ok(xml.indexOf("<Hosts>") < xml.indexOf("<Requirements>"));
  assert.ok(xml.indexOf("</Requirements>") < xml.indexOf("<DefaultSettings>"));
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
});

test("Word manifest version meets Microsoft's minimum 1.0 devx validation", async () => {
  const xml = await readFile(new URL("../manifests/word.xml", import.meta.url), "utf8");
  assert.match(xml, /<Version>[1-9]\d*\.\d+\.\d+\.\d+<\/Version>/);
});

for (const [provider, sdk] of [
  ["claude", claude],
  ["codebuddy", codebuddy],
]) {
  test(`${provider} review results stay inline without changing native snapshots or explicit XML`, async () => {
    const raw = "<pkg:package>" + "x".repeat(120000) + "</pkg:package>";
    const comment = {
      id: "comment-1",
      text: "Exact review text",
      resolved: false,
      anchor: { paragraph_ids: ["native-id"], current: "550 credits", original: "", ooxml: raw },
      replies: [{ id: "reply-1", text: "Exact reply" }],
    };
    const responses = {
      office_list_comments: { comment_count: 1, comments: [comment] },
      office_add_comment: { verified: true, action: "add", comment },
      office_manage_comment: { verified: true, action: "edit", comment },
      office_read_revisions: {
        snapshot_id: "snapshot",
        enumeration_complete: false,
        revisions: [{ index: 0, type: "Added", current: "550 credits", original: "", ooxml: raw }],
      },
      office_read_formatting: {
        font: { bold: true },
        current: "550 credits",
        snapshot_id: "format-snapshot",
        ooxml: raw,
        ooxml_length: raw.length + 10,
        next_ooxml_offset: raw.length,
      },
    };
    const original = structuredClone(responses);
    const server = createOfficeBridgeMcp(
      { callTaskpaneTool: async (name) => responses[name] },
      "word",
      "pane",
      sdk,
    );
    const handlers = server.instance._registeredTools;
    for (const name of Object.keys(responses)) {
      const reply = await handlers[name].handler({}, {});
      assert.ok(reply.content[0].text.length < 2000, name);
      const projected = JSON.parse(reply.content[0].text);
      assert.match(projected.ooxml_access, /omits raw OOXML only/);
      if (name === "office_read_revisions") {
        assert.equal(projected.snapshot_id, "snapshot");
        assert.equal(projected.enumeration_complete, false);
        assert.equal(projected.revisions[0].index, 0);
      } else if (name === "office_read_formatting") {
        assert.equal(projected.font.bold, true);
        assert.equal(projected.ooxml_length, raw.length + 10);
      } else {
        const c = projected.comment ?? projected.comments[0];
        assert.equal(c.text, comment.text);
        assert.equal(c.anchor.current, "550 credits");
        assert.deepEqual(c.replies, comment.replies);
        assert.equal(c.anchor.ooxml_omitted, true);
        if (projected.comment) assert.equal(projected.verified, true);
      }
    }
    for (const name of [
      "office_list_comments",
      "office_read_revisions",
      "office_read_formatting",
    ]) {
      const reply = await handlers[name].handler({ include_ooxml: true }, {});
      assert.deepEqual(JSON.parse(reply.content[0].text), responses[name]);
      assert.equal(handlers[name].inputSchema.safeParse({ include_ooxml: true }).success, true);
    }
    const page = await handlers.office_read_formatting.handler({ ooxml_limit: 8000 }, {});
    assert.deepEqual(JSON.parse(page.content[0].text), responses.office_read_formatting);
    assert.deepEqual(responses, original, "Native response/snapshot was mutated");
  });
}
