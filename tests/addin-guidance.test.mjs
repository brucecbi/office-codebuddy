import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { runInNewContext } from "node:vm";
import {
  addinActivationInstructions,
  addinInstallRecoveryInstructions,
} from "../app/addin-guidance.mjs";

test("Mac activation copy starts with the ordinary gallery and treats initialization as fallback", () => {
  const text = addinActivationInstructions("darwin");
  assert.match(text, /save your work.*fully quit and reopen/);
  assert.match(text, /Home → Add-ins → Developer Add-ins and single-click/);
  assert.match(text, /only if manual menu activation still fails/);
  assert.match(text, /registration, pane connection and model availability are checked separately/);
});

test("manual recovery respects platform and rendered HTTPS, without assuming an upload dialog", () => {
  const mac = addinInstallRecoveryInstructions("darwin");
  assert.match(mac, /Data\/Documents\/wef\//);
  assert.match(mac, /HTTP templates unchanged into an HTTPS installation/);
  assert.match(mac, /Upload My Add-in dialog is not required/);
  assert.match(mac, /Do not clear Office caches or change certificate trust/);
  for (const text of [
    addinActivationInstructions("win32"),
    addinInstallRecoveryInstructions("win32"),
  ])
    assert.doesNotMatch(text, /on macOS|Developer Add-ins|Initialize Word/);
  assert.match(addinInstallRecoveryInstructions("win32"), /WEF\\Developer/);
  for (const text of [
    addinActivationInstructions("linux"),
    addinInstallRecoveryInstructions("linux"),
  ])
    assert.match(text, /unsupported|not been verified/);
});

test("actual Electron install success and failure dialogs use platform-specific guidance", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const body = source.slice(
    source.indexOf("async function runInstall("),
    source.indexOf("async function runUninstall("),
  );
  const dialogs = [];
  let fail = false,
    refreshes = 0;
  const runInstall = runInNewContext(body + "\nrunInstall", {
    installAddin: async () => {
      if (fail) throw new Error("Synthetic install failure");
      return { catalog: "synthetic" };
    },
    refreshAddinInstalled: async () => {
      refreshes++;
    },
    dialog: {
      showMessageBox: async (options) => {
        dialogs.push(options);
      },
    },
    process: { platform: "darwin" },
    logStream: { write: () => {} },
    addinActivationInstructions: () => addinActivationInstructions("darwin"),
    addinInstallRecoveryInstructions: () => addinInstallRecoveryInstructions("darwin"),
  });
  await runInstall({ interactive: true });
  assert.equal(refreshes, 1);
  assert.match(dialogs[0].detail, /Developer Add-ins and single-click/);
  assert.match(dialogs[0].message, /activation must be checked separately/);
  fail = true;
  await runInstall({ interactive: true });
  assert.equal(refreshes, 1);
  assert.equal(dialogs[1].type, "error");
  assert.match(dialogs[1].detail, /Upload My Add-in dialog is not required/);
  assert.doesNotMatch(dialogs[1].detail, /Insert → My Add-ins → Upload/);
  assert.ok(source.includes('from "./addin-guidance.mjs"'));
  assert.match(
    source.slice(source.indexOf("async function offerFirstRunInstall(")),
    /addinActivationInstructions\(\)/,
  );
});

test("actual Word initialization dialog makes the per-process fallback conditional", async () => {
  const source = await readFile(new URL("../app/main.mjs", import.meta.url), "utf8");
  const start = source.indexOf("async function runInitializeWordAddin(");
  const end = source.indexOf("async function refreshAddinInstalled(", start);
  assert.ok(start >= 0 && end > start, "Word initialization test anchors must remain valid");
  const dialogs = [];
  let fail = false,
    opens = 0;
  const runInitializeWordAddin = runInNewContext(
    source.slice(start, end) + "\nrunInitializeWordAddin",
    {
      openWordBootstrap: async ({ openPath }) => {
        if (fail) throw new Error("Synthetic bootstrap failure");
        await openPath("synthetic.docx");
      },
      shell: {
        openPath: async (path) => {
          assert.equal(path, "synthetic.docx");
          opens++;
          return "";
        },
      },
      dialog: {
        showMessageBox: async (options) => {
          dialogs.push(options);
        },
      },
      logStream: { write: () => {} },
    },
  );
  assert.equal(await runInitializeWordAddin(), true);
  assert.equal(opens, 1);
  assert.equal(dialogs[0].type, "info");
  assert.match(dialogs[0].message, /Wait for the Draftspect pane in the synthetic document/);
  assert.match(dialogs[0].detail, /Close the synthetic document without saving/);
  assert.match(
    dialogs[0].detail,
    /Repeat Initialize Word Add-in after fully quitting and restarting Word only if manual menu activation still fails/,
  );
  assert.match(dialogs[0].detail, /not permanent registration/);
  assert.match(dialogs[0].detail, /does not start the daemon/);
  fail = true;
  assert.equal(await runInitializeWordAddin(), false);
  assert.equal(opens, 1);
  assert.equal(dialogs[1].type, "error");
  assert.equal(dialogs[1].message, "Synthetic bootstrap failure");
});
