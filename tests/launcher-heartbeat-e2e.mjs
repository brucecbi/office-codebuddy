// Compiles the actual Swift liveness policy; synthetic time, NOT native sleep evidence.
import assert from "node:assert/strict";
import { mkdir, writeFile, readFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { resolve, join } from "node:path";
import { createHash } from "node:crypto";
const out = resolve(process.argv[2]);
await mkdir(out, { recursive: true, mode: 0o700 });
const source = resolve("app/macos-launch-monitor.swift");
const probe = join(out, "HeartbeatProbe.swift");
await writeFile(
  probe,
  `
import Foundation
@main struct Probe {
 static func main() {
  var w = LaunchHeartbeatWatch()
  precondition(w.sample(beat: 1, now: 0) == nil)
  for n in 1...14 { precondition(w.sample(beat: 1, now: Double(n * 2)) == nil) }
  precondition(w.sample(beat: 1, now: 30) == "main-unresponsive-suspected")
  precondition(w.sample(beat: 1, now: 32) == nil) // one visible warning
  precondition(w.sample(beat: 2, now: 34) == "main-responsive")
  precondition(w.sample(beat: nil, now: 36) == nil) // partial file read tolerated
  precondition(w.sample(beat: 3, now: 38) == nil)
  precondition(w.sample(beat: 3, now: 100) == nil) // delayed monitor gets grace
  precondition(w.sample(beat: 3, now: 102, sleepGeneration: 1, suspended: true) == nil)
  precondition(w.sample(beat: 3, now: 200, sleepGeneration: 2) == nil) // wake gets grace
  for n in 1...14 { precondition(w.sample(beat: 3, now: 200 + Double(n * 2), sleepGeneration: 2) == nil) }
  precondition(w.sample(beat: 3, now: 230, sleepGeneration: 2) == "main-unresponsive-suspected")
  precondition(w.sample(beat: 4, now: 232, sleepGeneration: 2) == "main-responsive")
  let activity = LaunchHostActivity(); activity.update(suspended: true)
  precondition(activity.snapshot().suspended && activity.snapshot().generation == 1)
  activity.update(suspended: false)
  precondition(!activity.snapshot().suspended && activity.snapshot().generation == 2)
  print("PASS: threshold, one-warning, progress-recovery, partial-read, monitor-delay, sleep-wake-grace, host-activity")
 }
}
`,
  { flag: "wx" },
);
const binary = join(out, "HeartbeatProbe");
const compile = spawnSync(
  "/usr/bin/xcrun",
  [
    "swiftc",
    "-parse-as-library",
    "-module-cache-path",
    join(out, "module-cache"),
    source,
    probe,
    "-o",
    binary,
  ],
  { encoding: "utf8", timeout: 120000 },
);
await writeFile(join(out, "compile.log"), compile.stdout + compile.stderr, { flag: "wx" });
assert.equal(compile.status, 0, compile.stderr);
const run = spawnSync(binary, [], { encoding: "utf8", timeout: 10000 });
await writeFile(join(out, "run.log"), run.stdout + run.stderr, { flag: "wx" });
assert.equal(run.status, 0, run.stderr);
await writeFile(
  join(out, "result.json"),
  JSON.stringify(
    {
      status: "PASS",
      scope: "COMPILED_SWIFT_SYNTHETIC_CLOCK",
      checks: 7,
      modelRequests: 0,
      sourceSha256: createHash("sha256")
        .update(await readFile(source))
        .digest("hex"),
    },
    null,
    2,
  ),
  { flag: "wx" },
);
console.log(run.stdout.trim());
