import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toolReadParagraphs,
  toolGetDocumentText,
  toolGetSelection,
  toolApplyTextReplacement,
} from "../taskpane/shared/tools-word.js";
import { planReplacement } from "../taskpane/shared/word-edit-plan.js";
import { createOperationQueue } from "../taskpane/shared/operation-queue.js";

// Synthetic Office.js proxy: proves call-path semantics, NOT native Word rendering.
function fixture({
  reviewed = true,
  wrongReadback = false,
  duplicateSearch = false,
  serializedTracking = false,
} = {}) {
  let current = "The price is USD 100.";
  let writes = 0,
    restores = 0;
  const reviewedRange = () => ({
    getOoxml: () => ({
      value: serializedTracking
        ? `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>${current}</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/settings.xml"><pkg:xmlData><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${doc.changeTrackingMode === "TrackAll" ? "<w:trackRevisions/>" : ""}</w:settings></pkg:xmlData></pkg:part></pkg:package>`
        : `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>${current}</w:t></w:r></w:p>`,
    }),
    getReviewedText(version) {
      assert.equal(version, "Current");
      if (!reviewed) throw new Error("Unsupported");
      return { value: current };
    },
  });
  const paragraph = {
    uniqueLocalId: "A1",
    text: "The price is USD 90. The price is USD 100.",
    style: "Normal",
    listItemOrNullObject: { isNullObject: true },
    getRange: reviewedRange,
    search(find) {
      const item = {
        text: find,
        insertText(replacement) {
          assert.equal(doc.changeTrackingMode, "TrackAll");
          writes++;
          current = wrongReadback ? "unexpected" : current.replace(find, replacement);
        },
      };
      return { load() {}, items: duplicateSearch ? [item, item] : [item] };
    },
  };
  const paragraphs = { items: [paragraph], load() {} };
  const selection = {
    ...reviewedRange(),
    paragraphs,
    text: paragraph.text,
    load() {},
    select() {
      restores++;
    },
  };
  const body = { ...reviewedRange(), paragraphs, text: paragraph.text, load() {} };
  const doc = { body, changeTrackingMode: "Off", load() {}, getSelection: () => selection };
  const context = { document: doc, sync: async () => {} };
  globalThis.Word = {
    run: async (fn) => fn(context),
    ChangeTrackingVersion: { current: "Current" },
    ChangeTrackingMode: { trackAll: "TrackAll" },
    InsertLocation: { replace: "Replace" },
  };
  return {
    doc,
    setCurrent: (text) => (current = text),
    get current() {
      return current;
    },
    get writes() {
      return writes;
    },
    get restores() {
      return restores;
    },
  };
}
const args = { paragraph_ids: ["A1"], find: "100", replace: "120", match_case: true };

test("recording Off does not bypass Current reads, including preview:false and selection", async () => {
  const f = fixture();
  const body = await toolGetDocumentText();
  const full = await toolReadParagraphs({ preview: false });
  const preview = await toolReadParagraphs();
  const selection = await toolGetSelection();
  assert.equal(body.text, f.current);
  assert.equal(full.paragraphs[0].text, f.current);
  assert.equal(preview.paragraphs[0].text, f.current);
  assert.equal(selection.text, f.current);
  assert.equal(body.coverage.text_version, "current");
  assert.ok(body.coverage.excluded.includes("footnotes"));
  assert.equal(body.coverage.scope, "body");
  assert.ok(f.restores > 0);
  assert.equal(f.writes, 0);
});

test("unsupported Current API fails explicitly without substituting raw revision text", async () => {
  fixture({ reviewed: false });
  await assert.rejects(toolGetDocumentText(), /REVIEWED_TEXT_UNAVAILABLE/);
  await assert.rejects(toolReadParagraphs({ preview: false }), /REVIEWED_TEXT_UNAVAILABLE/);
  await assert.rejects(toolGetSelection(), /REVIEWED_TEXT_UNAVAILABLE/);
});

test("controlled replacement preserves surrounding text, tracks, reads back and restores mode", async () => {
  const f = fixture();
  let approvals = 0;
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  assert.equal(plan[0].after, "The price is USD 120.");
  const result = await toolApplyTextReplacement(args, plan, {
    guard() {},
    beforeWrite() {
      approvals++;
    },
  });
  assert.equal(f.current, "The price is USD 120.");
  assert.equal(f.writes, 1);
  assert.equal(approvals, 1);
  assert.equal(f.doc.changeTrackingMode, "Off");
  assert.equal(result.verification, "current_text_readback_matched");
});

test("manual edits after preview and cancellation stop before mutation", async () => {
  const f = fixture();
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  f.setCurrent("Manually updated USD 100.");
  await assert.rejects(
    toolApplyTextReplacement(args, plan, {
      guard() {},
      beforeWrite() {
        assert.fail("must not write");
      },
    }),
    /STALE_PROPOSAL/,
  );
  assert.equal(f.writes, 0);
  f.setCurrent(plan[0].before);
  await assert.rejects(
    toolApplyTextReplacement(args, plan, {
      guard() {
        throw new Error("cancelled");
      },
      beforeWrite() {},
    }),
    /cancelled/,
  );
  assert.equal(f.writes, 0);
});

test("tracking mode serialized into native OOXML is changed only after the final precondition", async () => {
  const f = fixture({ serializedTracking: true });
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  const result = await toolApplyTextReplacement(args, plan, { guard() {}, beforeWrite() {} });
  assert.equal(f.current, "The price is USD 120.");
  assert.equal(f.writes, 1);
  assert.equal(f.doc.changeTrackingMode, "Off");
  assert.equal(result.verification, "current_text_readback_matched");
});

test("pre-write cancellation or receipt failure never restores a tracking mode this call did not change", async () => {
  for (const failureStage of ["guard", "receipt"]) {
    const f = fixture();
    const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
    let mode = "Off",
      guards = 0,
      receipts = 0;
    const modeWrites = [];
    Object.defineProperty(f.doc, "changeTrackingMode", {
      get: () => mode,
      set(value) {
        modeWrites.push(value);
        mode = value;
      },
    });
    await assert.rejects(
      toolApplyTextReplacement(args, plan, {
        guard() {
          if (++guards === 2) {
            mode = "TrackAll"; // Native user change, outside this operation.
            if (failureStage === "guard") throw new Error("cancelled before write");
          }
        },
        beforeWrite() {
          if (failureStage === "receipt") throw new Error("receipt unavailable");
          receipts++;
        },
      }),
      /cancelled before write|receipt unavailable/,
    );
    assert.equal(receipts, 0);
    assert.equal(f.writes, 0);
    assert.deepEqual(modeWrites, [], failureStage);
    assert.equal(mode, "TrackAll", "User tracking choice must remain untouched");
  }
});

test("tracking setter failure still attempts restoration and remains RESULT_UNKNOWN", async () => {
  const f = fixture();
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  let mode = "Off",
    receipts = 0;
  const modeWrites = [];
  Object.defineProperty(f.doc, "changeTrackingMode", {
    get: () => mode,
    set(value) {
      modeWrites.push(value);
      mode = value;
      if (value === "TrackAll")
        throw new Error("tracking setter failed after attempted assignment");
    },
  });
  await assert.rejects(
    toolApplyTextReplacement(args, plan, {
      guard() {},
      beforeWrite() {
        receipts++;
      },
    }),
    /RESULT_UNKNOWN.*tracking setter failed/,
  );
  assert.equal(receipts, 1);
  assert.equal(f.writes, 0);
  assert.deepEqual(modeWrites, ["TrackAll", "Off"]);
  assert.equal(mode, "Off");
});

test("ambiguous text, positional IDs, and inconsistent Word ranges are rejected", async () => {
  assert.throws(() => planReplacement({ addressing: "index", paragraphs: [] }, args), /Stable/);
  assert.throws(
    () =>
      planReplacement(
        { addressing: "uniqueLocalId", paragraphs: [{ id: "A1", text: "100 and 100" }] },
        args,
      ),
    /unambiguous/,
  );
  const f = fixture({ duplicateSearch: true });
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  await assert.rejects(
    toolApplyTextReplacement(args, plan, { guard() {}, beforeWrite() {} }),
    /AMBIGUOUS_RANGE/,
  );
  assert.equal(f.writes, 0);
});

test("write/readback mismatch is RESULT_UNKNOWN, not success or automatic retry", async () => {
  const f = fixture({ wrongReadback: true });
  const plan = planReplacement(await toolReadParagraphs({ preview: false }), args);
  await assert.rejects(
    toolApplyTextReplacement(args, plan, { guard() {}, beforeWrite() {} }),
    /RESULT_UNKNOWN/,
  );
  assert.equal(f.writes, 1);
  assert.equal(f.doc.changeTrackingMode, "Off");
});

test("duplicate write ID executes once; reload refuses a persisted attempted ID", async () => {
  const data = new Map();
  const storage = { getItem: (k) => data.get(k), setItem: (k, v) => data.set(k, v) };
  const q = createOperationQueue({ storage });
  let calls = 0;
  const options = { mutates: true, deadline: Date.now() + 1000 };
  const action = async ({ beforeWrite }) => {
    beforeWrite();
    calls++;
    return "verified";
  };
  assert.deepEqual(
    await Promise.all([q.run("id", options, action), q.run("id", options, action)]),
    ["verified", "verified"],
  );
  assert.equal(calls, 1);
  await assert.rejects(q.run("id", options, action), /RESULT_UNKNOWN/);
  await assert.rejects(
    createOperationQueue({ storage }).run("id", options, action),
    /RESULT_UNKNOWN/,
  );
  assert.equal(calls, 1);
});

test("settled queue records are evicted while pending duplicate reads still share a promise", async () => {
  const q = createOperationQueue();
  const options = { mutates: false, deadline: Date.now() + 1000 };
  let release,
    calls = 0;
  const wait = new Promise((resolve) => {
    release = resolve;
  });
  const first = q.run("read", options, async () => {
    calls++;
    await wait;
    return "first";
  });
  assert.equal(
    q.run("read", options, () => assert.fail()),
    first,
  );
  release();
  assert.equal(await first, "first");
  assert.equal(
    await q.run("read", options, () => {
      calls++;
      return "second";
    }),
    "second",
  );
  assert.equal(calls, 2);
  await assert.rejects(
    q.run("failed", options, () => {
      throw new Error("fixture failure");
    }),
    /fixture failure/,
  );
  assert.equal(await q.run("failed", options, () => "retry"), "retry");
});

test("cancelled confirmation and expired queued writes cannot later apply", async () => {
  const q = createOperationQueue();
  let release,
    calls = 0;
  const waiting = new Promise((r) => (release = r));
  const first = q.run(
    "a",
    { mutates: true, deadline: Date.now() + 1000 },
    async ({ beforeWrite }) => {
      await waiting;
      beforeWrite();
      calls++;
    },
  );
  await Promise.resolve();
  const second = q.run("b", { mutates: true, deadline: Date.now() - 1 }, async () => {
    calls++;
  });
  const assertions = [
    assert.rejects(first, /CANCELLED_OR_EXPIRED/),
    assert.rejects(second, /CANCELLED_OR_EXPIRED/),
  ];
  q.cancel();
  release();
  await Promise.all(assertions);
  assert.equal(calls, 0);
});

test("a cancel that arrives while a write function is in flight is not success", async () => {
  const q = createOperationQueue();
  let release;
  const gate = new Promise((resolve) => {
    release = resolve;
  });
  const running = q.run(
    "late-cancel",
    { mutates: true, deadline: Date.now() + 1000 },
    async ({ beforeWrite }) => {
      beforeWrite();
      await gate;
      return "wrote";
    },
  );
  await Promise.resolve();
  q.cancel();
  release();
  await assert.rejects(running, /CANCELLED_OR_EXPIRED/);
});
