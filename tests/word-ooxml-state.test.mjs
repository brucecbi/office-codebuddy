import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { stableOoxml, verifyAcceptedInsertion } from "../taskpane/shared/word-ooxml-state.js";

const native = JSON.parse(
  readFileSync(new URL("./fixtures/word-native-ooxml-snapshots.json", import.meta.url)),
);
const target = { fragment: true, paragraphIdentity: true };
const body = { paragraphIdentity: true };
const reviewed = JSON.parse(
  readFileSync(new URL("./fixtures/word-native-reviewed-ooxml-snapshots.json", import.meta.url)),
);
const accepted = JSON.parse(
  readFileSync(new URL("./fixtures/word-native-accepted-insertion.json", import.meta.url)),
);

test("unedited native Mac target and body exports compare equal despite observed generated metadata", () => {
  for (const [name, options] of [
    ["target", target],
    ["body", body],
  ]) {
    assert.notEqual(native[name][0], native[name][1]);
    assert.equal(stableOoxml(native[name][0], options), stableOoxml(native[name][1], options));
  }
});
test("native-derived exports with a comment and tracked header/format edits preserve review state while normalizing comment-paragraph rsids", () => {
  for (const [name, options] of [
    ["target", target],
    ["body", body],
  ]) {
    assert.notEqual(reviewed[name][0], reviewed[name][1]);
    assert.equal(stableOoxml(reviewed[name][0], options), stableOoxml(reviewed[name][1], options));
  }
  const raw = reviewed.body[0];
  assert.ok(raw.includes('w:author="Synthetic Reviewer"'));
  assert.notEqual(
    stableOoxml(raw, body),
    stableOoxml(raw.replace('w:author="Synthetic Reviewer"', 'w:author="Another Reviewer"'), body),
  );
});
test("native paragraph IDs are required before generated paragraph export IDs can be normalized", () => {
  assert.notEqual(stableOoxml(native.target[0]), stableOoxml(native.target[1]));
});
test("real run formatting, text, stable paragraph IDs and nonobserved rsids remain exact evidence", () => {
  const raw = native.body[0];
  for (const [before, after] of [
    ['w:color w:val="336699"', 'w:color w:val="FF0000"'],
    ["Formatting target.", "Formatting changed."],
    ['w14:paraId="7DBC6D43"', 'w14:paraId="12345678"'],
    ['w:rsidRPr="00903E77"', 'w:rsidRPr="12345678"'],
    ['w:rsidRoot w:val="00FC3148"', 'w:rsidRoot w:val="12345678"'],
  ]) {
    assert.ok(raw.includes(before), before);
    assert.notEqual(stableOoxml(raw, body), stableOoxml(raw.replace(before, after), body), before);
  }
});
test("review metadata and reference IDs are never filtered, including references to a generated paragraph", () => {
  const wrap = (content) =>
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${content}</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
  const content =
    '<w:p w14:paraId="12345678"><w:pPr><w:pStyle w:val="Style1"/></w:pPr><w:bookmarkStart w:id="1" w:name="Anchor"/><w:commentRangeStart w:id="2"/><w:ins w:id="3" w:author="Alice" w:date="2026-09-12T00:00:00Z"><w:r><w:t>&lt;w:p w14:paraId="ABCDEF12"/&gt;</w:t></w:r></w:ins><w:commentRangeEnd w:id="2"/><w:hyperlink r:id="rId7"/><w15:commentEx w15:paraId="12345678" w15:paraIdParent="99999999"/></w:p>';
  for (const [before, after] of [
    ['w:id="1"', 'w:id="9"'],
    ['w:name="Anchor"', 'w:name="Other"'],
    ['w:id="2"', 'w:id="9"'],
    ['w:id="3"', 'w:id="9"'],
    ['w:author="Alice"', 'w:author="Bob"'],
    ["2026-09-12T00:00:00Z", "2026-09-13T00:00:00Z"],
    ["Style1", "Style2"],
    ["rId7", "rId8"],
    ["ABCDEF12", "ABCDEF34"],
    ['w14:paraId="12345678"', 'w14:paraId="12345679"'],
    ['w15:paraIdParent="99999999"', 'w15:paraIdParent="88888888"'],
  ])
    assert.notEqual(
      stableOoxml(wrap(content), target),
      stableOoxml(wrap(content.replace(before, after)), target),
      before,
    );
});
test("plain text, malformed XML, unknown namespaces, comments and CDATA fail closed or remain verbatim", () => {
  for (const raw of [
    'literal w:rsidR="12345678"',
    '{"ooxml":"w14:paraId=12345678"}',
    '<pkg:package><w:p w14:paraId="12345678"/></pkg:package>',
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><broken></pkg:package>',
    '<pkg:package xmlns:pkg="urn:other"><w:p w14:paraId="12345678"/></pkg:package>',
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><!-- <w:p w14:paraId="12345678"/> --><![CDATA[<w:p w14:paraId="12345678"/>]]></pkg:package>',
  ])
    assert.equal(stableOoxml(raw, target), raw);
});

test("native collapsed replacement proof accepts only selected insertion unwrap and exact later exported revision ID decrements", () => {
  assert.equal(verifyAcceptedInsertion(accepted.before, accepted.after, [accepted.selected]), true);
  assert.equal(
    verifyAcceptedInsertion(accepted.before, accepted.before, [accepted.selected]),
    false,
  );
  assert.equal(verifyAcceptedInsertion(accepted.before, accepted.after, []), false);
  assert.equal(
    verifyAcceptedInsertion(accepted.before, accepted.after, [
      accepted.selected,
      accepted.selected,
    ]),
    false,
  );
  assert.equal(
    verifyAcceptedInsertion(accepted.before, accepted.after, [
      { ...accepted.selected, type: "Deleted" },
    ]),
    false,
  );
  // The counterpart was already present before accept and must remain deleted,
  // not be accepted with the selected insertion or substituted with new text.
  for (const [old, changed] of [
    [
      "<w:delText>DRAFTSPECT REVISION DIAGNOSTIC 100</w:delText>",
      "<w:delText>DRAFTSPECT REVISION DIAGNOSTIC 999</w:delText>",
    ],
    ['<w:del w:id="4"', '<w:del w:id="99"'],
    ['w:author="Synthetic Reviewer"', 'w:author="Other Reviewer"'],
    ['w:date="2026-09-12T14:50:00Z"', 'w:date="2026-09-12T14:51:00Z"'],
    ['w:rsidDel="00F5329F"', 'w:rsidDel="00F5329E"'],
    ["<w:rPrChange ", '<w:rPrChange bogus="1" '],
  ]) {
    assert.ok(accepted.after.includes(old), old);
    assert.equal(
      verifyAcceptedInsertion(accepted.before, accepted.after.replace(old, changed), [
        accepted.selected,
      ]),
      false,
      old,
    );
  }
  const removedCounterpart = accepted.after.replace(/<w:del w:id="4"[^>]*>[\s\S]*?<\/w:del>/, "");
  assert.notEqual(removedCounterpart, accepted.after);
  assert.equal(
    verifyAcceptedInsertion(accepted.before, removedCounterpart, [accepted.selected]),
    false,
  );
});

test("insertion proof rejects ambiguous duplicate subtrees and malformed packages", () => {
  const selectedNode = /<w:ins w:id="4"[^>]*>[\s\S]*?<\/w:ins>/.exec(accepted.before)?.[0];
  assert.ok(selectedNode);
  const duplicate = accepted.before.replace(
    selectedNode,
    selectedNode + selectedNode.replace('w:id="4"', 'w:id="44"'),
  );
  assert.equal(verifyAcceptedInsertion(duplicate, accepted.after, [accepted.selected]), false);
  assert.equal(verifyAcceptedInsertion("not XML", accepted.after, [accepted.selected]), false);
  assert.equal(verifyAcceptedInsertion(accepted.before, "not XML", [accepted.selected]), false);
});

test("body insertion proof cannot select an identical insertion found only in a header story", () => {
  const packageXml = (documentContent, headerContent = null) =>
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document><w:body><w:p>${documentContent}</w:p></w:body></w:document></pkg:xmlData></pkg:part>` +
    (headerContent === null
      ? ""
      : `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData><w:hdr><w:p>${headerContent}</w:p></w:hdr></pkg:xmlData></pkg:part>`) +
    "</pkg:package>";
  const insertion = (id) =>
    `<w:ins w:id="${id}" w:author="Synthetic Reviewer"><w:r><w:t>New</w:t></w:r></w:ins>`;
  const row = {
    type: "Added",
    current: "New",
    original: "",
    text: "New",
    ooxml: packageXml(insertion(0)),
  };
  assert.equal(
    verifyAcceptedInsertion(
      packageXml("", insertion(1)),
      packageXml("", "<w:r><w:t>New</w:t></w:r>"),
      [row],
    ),
    false,
  );
});
