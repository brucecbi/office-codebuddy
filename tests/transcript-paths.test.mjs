import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { locateSessionFile, projectsDirectory, readTranscript } from "../daemon/transcript.mjs";
import { resumeCommands } from "../daemon/session-archive.mjs";
import { codebuddyConfigEnvironment } from "../daemon/runtime/runtime-codebuddy.mjs";

test("CodeBuddy history locator and exact-ID command honor custom config; Claude default unchanged", async () => {
  const root = await mkdtemp(join(tmpdir(), "draftspect-transcript-paths-"));
  const env = { CODEBUDDY_CONFIG_DIR: join(root, "custom config") };
  const projectsDir = projectsDirectory("codebuddy", { env, home: "/synthetic-home" });
  assert.deepEqual(codebuddyConfigEnvironment(env), {
    env: { CODEBUDDY_CONFIG_DIR: env.CODEBUDDY_CONFIG_DIR },
  });
  assert.deepEqual(codebuddyConfigEnvironment({}), {});
  assert.equal(projectsDir, join(env.CODEBUDDY_CONFIG_DIR, "projects"));
  assert.equal(
    projectsDirectory("claude", { env, home: "/synthetic-home" }),
    "/synthetic-home/.claude/projects",
  );
  await mkdir(join(projectsDir, "synthetic-project"), { recursive: true });
  const source = join(projectsDir, "synthetic-project", "fixture-id.jsonl");
  await writeFile(
    source,
    JSON.stringify({
      type: "message",
      role: "assistant",
      content: [{ type: "output_text", text: "Synthetic custom config history" }],
    }) + "\n",
  );
  assert.equal(
    await locateSessionFile("fixture-id", { provider: "codebuddy", projectsDir }),
    source,
  );
  assert.equal(
    (await readTranscript("fixture-id", { provider: "codebuddy", projectsDir })).events[0].text,
    "Synthetic custom config history",
  );
  assert.equal(await locateSessionFile("missing", { provider: "codebuddy", projectsDir }), null);
  const commands = resumeCommands(
    { provider: "codebuddy", session_id: "fixture-id", cwd: root },
    { env },
  );
  assert.ok(
    commands.resume.includes(
      `CODEBUDDY_CONFIG_DIR='${env.CODEBUDDY_CONFIG_DIR}' codebuddy --resume 'fixture-id'`,
    ),
  );
  assert.ok(commands.reuse_as_new_session.includes("--fork-session"));
});
