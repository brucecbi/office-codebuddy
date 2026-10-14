// Opt-in: compiled production Swift monitor + controlled Node children; not Electron/Office.
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdir, writeFile, readFile, mkdtemp } from "node:fs/promises";
import { createHash } from "node:crypto";
import { resolve, join } from "node:path";
import { tmpdir } from "node:os";

const output = resolve(
  process.argv[2] || (await mkdtemp(join(tmpdir(), "draftspect-monitor-e2e-"))),
);
await mkdir(output, { recursive: true, mode: 0o700 });
const root = resolve(new URL("..", import.meta.url).pathname);
const monitorSource = join(root, "app/macos-launch-monitor.swift");
const probe = join(output, "MonitorProbe.swift"),
  binary = join(output, "MonitorProbe");
await writeFile(
  probe,
  `
import Foundation
import Darwin
@main struct Probe {
  static func main() throws {
    let args = CommandLine.arguments
    if args.count == 2 && args[1] == "diagnostic-write-failure" {
      let pipe = Pipe(); try pipe.fileHandleForWriting.close()
      let saved = writeLaunchDiagnostic(pipe.fileHandleForWriting, event: "incident", pid: 0)
      guard !saved && diagnosticSummary(saved: saved).contains("未完整保存") && safeDiagnosticCode("DESKTOP_SECRET_SENTINEL") == nil else { exit(1) }
      print(#"{"diagnosticFailureVisible":true}"#); return
    }
    let state = URL(fileURLWithPath: args[1])
    try FileManager.default.createDirectory(at: state, withIntermediateDirectories: true, attributes: [.posixPermissions: 0o700])
    let attempt = try createLaunchAttempt(stateDirectory: state)
    let child = Process()
    child.executableURL = URL(fileURLWithPath: args[2]); child.arguments = [args[3], args[4]]
    child.environment = ["DRAFTSPECT_LAUNCH_STATUS_PATH": attempt.statusURL.path, "DRAFTSPECT_LAUNCH_TOKEN": attempt.token]
    child.standardOutput = FileHandle.nullDevice; child.standardError = FileHandle.nullDevice
    try child.run()
    try JSONSerialization.data(withJSONObject: ["pid": child.processIdentifier]).write(to: state.appendingPathComponent("owned-child.json"))
    let outcome = monitorLaunch(child: child, attempt: attempt, timeout: 0.7, pollInterval: 0.02, terminateGrace: 0.2)
    let bytes = try JSONEncoder().encode(outcome)
    print(String(data: bytes, encoding: .utf8)!)
  }
}
`,
  { flag: "wx", mode: 0o600 },
);
const compile = spawnSync(
  "/usr/bin/xcrun",
  [
    "swiftc",
    "-parse-as-library",
    "-O",
    "-module-cache-path",
    join(output, "module-cache"),
    monitorSource,
    probe,
    "-o",
    binary,
  ],
  { encoding: "utf8", timeout: 120000 },
);
await writeFile(join(output, "compile.txt"), compile.stdout + compile.stderr, {
  flag: "wx",
  mode: 0o600,
});
assert.equal(compile.status, 0, compile.stderr);
const childSource = join(output, "synthetic-child.cjs");
await writeFile(
  childSource,
  `
const fs = require('node:fs');
const mode = process.argv[2], file = process.env.DRAFTSPECT_LAUNCH_STATUS_PATH;
const state = {schemaVersion:1,token:process.env.DRAFTSPECT_LAUNCH_TOKEN,pid:process.pid,phase:'js-started',desktopVisible:false,serviceReady:false};
const publish = () => fs.writeFileSync(file,JSON.stringify(state));
console.log('SECRET_SENTINEL_NOT_DIAGNOSTIC');
if(mode==='early-exit'){process.exit(7);}
if(mode==='service-only'){state.phase='service-ready';state.serviceReady=true;publish();setTimeout(()=>process.exit(0),100);}
if(mode==='duplicate'){state.phase='duplicate';publish();setTimeout(()=>process.exit(0),100);}
if(['normal','post-ready-signal','unmarked-exit'].includes(mode)){state.phase='desktop-visible';state.desktopVisible=true;publish();setTimeout(()=>{if(mode==='post-ready-signal')process.kill(process.pid,'SIGTERM');else{if(mode==='normal'){state.phase='quitting';publish();}process.exit(0);}},180);}
if(['timeout','wrong-token','wrong-pid','public-file','linked-file'].includes(mode)){if(mode!=='timeout'){state.phase='desktop-visible';state.desktopVisible=true;}if(mode==='wrong-token')state.token='BAD';if(mode==='wrong-pid')state.pid=1;publish();if(mode==='public-file')fs.chmodSync(file,0o644);if(mode==='linked-file')fs.linkSync(file,file+'.link');setInterval(()=>{},1000);}
`,
  { flag: "wx", mode: 0o600 },
);
const cases = [
  ["normal", { startupConfirmed: true, manualQuit: true, exitCode: 0 }],
  [
    "early-exit",
    {
      startupConfirmed: false,
      failureCode: "LAUNCH_EXIT_BEFORE_READY",
      exitCode: 7,
      lastPhase: "before-js-marker",
    },
  ],
  ["service-only", { startupConfirmed: false, failureCode: "LAUNCH_EXIT_BEFORE_READY" }],
  ["duplicate", { startupConfirmed: false, duplicate: true, exitCode: 0 }],
  [
    "post-ready-signal",
    {
      startupConfirmed: true,
      manualQuit: false,
      signal: 15,
      failureCode: "LAUNCH_MAIN_EXIT_UNEXPECTED",
    },
  ],
  [
    "unmarked-exit",
    { startupConfirmed: true, manualQuit: false, failureCode: "LAUNCH_MAIN_EXIT_UNEXPECTED" },
  ],
  ...["timeout", "wrong-token", "wrong-pid", "public-file", "linked-file"].map((name) => [
    name,
    { startupConfirmed: false, manualQuit: false, failureCode: "LAUNCH_READY_TIMEOUT" },
  ]),
];
const results = [];
const badLog = spawnSync(binary, ["diagnostic-write-failure"], { encoding: "utf8", timeout: 3000 });
assert.equal(badLog.status, 0, badLog.stderr);
assert.equal(JSON.parse(badLog.stdout).diagnosticFailureVisible, true);
results.push({
  name: "diagnostic-write-failure",
  status: "PASS",
  scope: "closed descriptor simulation, not a real full disk",
});
for (const [name, expected] of cases) {
  const began = Date.now();
  const result = spawnSync(
    binary,
    [join(output, "state-" + name), process.execPath, childSource, name],
    { encoding: "utf8", timeout: 7000 },
  );
  assert.equal(result.status, 0, `${name}: ${result.stderr}`);
  const outcome = JSON.parse(result.stdout);
  for (const [key, value] of Object.entries(expected))
    assert.equal(outcome[key], value, `${name}.${key}`);
  assert.doesNotMatch(result.stdout + result.stderr, /SECRET_SENTINEL|DRAFTSPECT_LAUNCH_TOKEN/);
  if (name === "normal") assert.ok(Date.now() - began >= 160);
  results.push({ name, status: "PASS", elapsedMs: Date.now() - began, outcome });
}
const sourceSha256 = createHash("sha256")
  .update(await readFile(monitorSource))
  .digest("hex");
await writeFile(
  join(output, "result.json"),
  JSON.stringify(
    {
      status: "PASS",
      scope: "COMPILED_PRODUCTION_SWIFT_MONITOR_WITH_SYNTHETIC_CHILDREN",
      modelRequests: 0,
      sourceSha256,
      results,
    },
    null,
    2,
  ) + "\n",
  { flag: "wx", mode: 0o600 },
);
console.log(
  JSON.stringify({
    status: "PASS",
    cases: results.length,
    output,
    scope: "compiled monitor; Electron and Office not exercised",
  }),
);
