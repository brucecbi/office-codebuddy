import Darwin
import Foundation

// One monitor per explicit launcher invocation. No login item, upload or auto-restart.
struct LaunchAttempt {
  let directory: URL
  let statusURL: URL
  let token: String
}

struct LaunchStatus: Decodable {
  let schemaVersion: Int
  let token: String
  let pid: Int32
  let phase: String
  let desktopVisible: Bool
  let serviceReady: Bool
  let failureCode: String?
  let heartbeat: UInt64?
}

// Monotonic liveness, not a crash diagnosis. Sleep/wake and a delayed monitor
// receive a fresh grace period. An unresponsive main is never killed/restarted.
struct LaunchHeartbeatWatch {
  var lastTick: TimeInterval?
  var lastProgress: TimeInterval?
  var lastBeat: UInt64?
  var generation = 0
  var warned = false
  mutating func sample(
    beat: UInt64?, now: TimeInterval, timeout: TimeInterval = 30,
    sleepGeneration: Int = 0, suspended: Bool = false
  ) -> String? {
    let delayed = lastTick.map { now - $0 > 6 } ?? false
    lastTick = now
    if suspended || generation != sleepGeneration || delayed {
      generation = sleepGeneration
      lastProgress = now
      lastBeat = beat
      // Keep an existing warning until fresh JS progress proves recovery.
      return nil
    }
    if lastProgress == nil { lastProgress = now }
    if let beat = beat, lastBeat != beat {
      lastBeat = beat
      lastProgress = now
      if warned {
        warned = false
        return "main-responsive"
      }
    } else if now - (lastProgress ?? now) >= timeout && !warned {
      warned = true
      return "main-unresponsive-suspected"
    }
    return nil
  }
}

final class LaunchHostActivity: @unchecked Sendable {
  private let lock = NSLock()
  private var generation = 0
  private var suspended = false
  func update(suspended value: Bool) {
    lock.lock()
    defer { lock.unlock() }
    generation += 1
    suspended = value
  }
  func snapshot() -> (generation: Int, suspended: Bool) {
    lock.lock()
    defer { lock.unlock() }
    return (generation, suspended)
  }
}

struct LaunchOutcome: Codable {
  let startupConfirmed: Bool
  let duplicate: Bool
  let manualQuit: Bool
  let exitCode: Int32
  let signal: Int32?
  let failureCode: String?
  let lastPhase: String
}

func privateDirectory(_ url: URL) throws {
  let descriptor = open(url.path, O_RDONLY | O_DIRECTORY | O_NOFOLLOW)
  guard descriptor >= 0 else { throw CocoaError(.fileReadNoPermission) }
  defer { close(descriptor) }
  var info = stat()
  guard fstat(descriptor, &info) == 0, info.st_uid == getuid(),
    info.st_mode & 0o077 == 0
  else { throw CocoaError(.fileReadNoPermission) }
}

func createLaunchAttempt(stateDirectory: URL) throws -> LaunchAttempt {
  let root = stateDirectory.resolvingSymlinksInPath()
  try privateDirectory(root)
  let attempts = root.appendingPathComponent("launch-attempts", isDirectory: true)
  try FileManager.default.createDirectory(
    at: attempts, withIntermediateDirectories: true, attributes: [.posixPermissions: 0o700])
  try privateDirectory(attempts)
  let token = UUID().uuidString
  let directory = attempts.appendingPathComponent(token, isDirectory: true)
  try FileManager.default.createDirectory(
    at: directory, withIntermediateDirectories: false, attributes: [.posixPermissions: 0o700])
  try privateDirectory(directory)
  let statusURL = directory.appendingPathComponent("status.json")
  let descriptor = open(statusURL.path, O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW, 0o600)
  guard descriptor >= 0 else { throw CocoaError(.fileWriteNoPermission) }
  let handle = FileHandle(fileDescriptor: descriptor, closeOnDealloc: true)
  defer { try? handle.close() }
  let initial: [String: Any] = [
    "schemaVersion": 1, "token": token, "pid": 0, "phase": "launcher-started",
    "desktopVisible": false, "serviceReady": false,
  ]
  try handle.write(
    contentsOf: JSONSerialization.data(withJSONObject: initial, options: [.sortedKeys]))
  try handle.synchronize()
  return LaunchAttempt(directory: directory, statusURL: statusURL, token: token)
}

func readLaunchStatus(_ attempt: LaunchAttempt, pid: Int32) -> LaunchStatus? {
  let descriptor = open(attempt.statusURL.path, O_RDONLY | O_NOFOLLOW | O_NONBLOCK)
  guard descriptor >= 0 else { return nil }
  let handle = FileHandle(fileDescriptor: descriptor, closeOnDealloc: true)
  defer { try? handle.close() }
  var info = stat()
  guard fstat(descriptor, &info) == 0, info.st_uid == getuid(),
    info.st_mode & mode_t(S_IFMT) == mode_t(S_IFREG), info.st_mode & 0o077 == 0,
    info.st_nlink == 1, info.st_size <= 8192,
    let data = try? handle.readToEnd(),
    let status = try? JSONDecoder().decode(LaunchStatus.self, from: data),
    status.schemaVersion == 1, status.token == attempt.token, status.pid == pid,
    ["js-started", "desktop-visible", "service-ready", "startup-failed", "duplicate", "quitting"]
      .contains(status.phase)
  else { return nil }
  return status
}

func safeDiagnosticCode(_ code: String?) -> String? {
  let allowed: Set<String> = [
    "LAUNCH_READY_TIMEOUT", "LAUNCH_SPAWN_FAILED", "LAUNCH_STATE_UNAVAILABLE", "LAUNCH_FAILED",
    "LAUNCH_MAIN_EXIT_UNEXPECTED", "LAUNCH_EXIT_BEFORE_READY",
    "DESKTOP_MAIN_UNRESPONSIVE_SUSPECTED",
    "DESKTOP_START_FAILED", "DESKTOP_LAUNCH_CHANNEL_INVALID", "DESKTOP_STATUS_LOAD_FAILED",
    "DESKTOP_STATUS_TIMEOUT", "DESKTOP_MAIN_EXCEPTION", "DESKTOP_MAIN_REJECTION",
    "DESKTOP_DAEMON_START_TIMEOUT", "DESKTOP_DIAGNOSTIC_UNAVAILABLE", "DESKTOP_LOG_UNAVAILABLE",
    "DESKTOP_SINGLE_INSTANCE_FAILED", "DESKTOP_TLS_REQUIRED", "DESKTOP_STATE_UNAVAILABLE",
    "DESKTOP_TRANSPORT_INVALID", "DESKTOP_PORT_INVALID", "DESKTOP_PORT_CHECK_TIMEOUT",
    "DESKTOP_PORT_IN_USE", "DESKTOP_PORT_UNAVAILABLE", "DESKTOP_ENV_INVALID",
    "DESKTOP_PROJECT_INVALID", "DESKTOP_ASSET_MISSING", "DESKTOP_DEPENDENCY_MISSING",
    "DESKTOP_TLS_INVALID", "DESKTOP_HTTPS_REQUIRED", "DESKTOP_CLI_INVALID",
    "DESKTOP_CLI_NOT_FOUND", "DESKTOP_NODE_INVALID", "DESKTOP_NODE_NOT_FOUND",
  ]
  guard let code = code, allowed.contains(code) else { return nil }
  return code
}

func launchVersions(projectRoot: String) -> [String: String] {
  var values: [String: String] = [:]
  for (label, relative) in [
    ("projectVersion", "package.json"), ("electronVersion", "node_modules/electron/package.json"),
  ] {
    let path = URL(fileURLWithPath: projectRoot).appendingPathComponent(relative)
    if let data = try? Data(contentsOf: path), data.count <= 65536,
      let object = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
      let version = object["version"] as? String,
      version.utf8.count <= 64,
      version.range(of: "^[0-9]+(?:\\.[0-9]+){1,3}(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?$", options: .regularExpression) != nil
    {
      values[label] = version
    }
  }
  let os = ProcessInfo.processInfo.operatingSystemVersion
  values["macOSVersion"] = "\(os.majorVersion).\(os.minorVersion).\(os.patchVersion)"
  #if arch(arm64)
    values["architecture"] = "arm64"
  #else
    values["architecture"] = "x86_64"
  #endif
  return values
}

@discardableResult
func writeLaunchDiagnostic(
  _ handle: FileHandle, event: String, pid: Int32, code: String? = nil,
  exitCode: Int32? = nil, signal: Int32? = nil, versions: [String: String] = [:]
) -> Bool {
  var entry: [String: Any] = [
    "schemaVersion": 1, "time": ISO8601DateFormatter().string(from: Date()),
    "event": event, "pid": pid,
  ]
  if let code = safeDiagnosticCode(code) { entry["code"] = code }
  if let exitCode = exitCode { entry["exitCode"] = exitCode }
  if let signal = signal { entry["signal"] = signal }
  for key in ["projectVersion", "electronVersion", "macOSVersion", "architecture"] {
    if let value = versions[key] { entry[key] = value }
  }
  do {
    let bytes = try JSONSerialization.data(withJSONObject: entry, options: [.sortedKeys])
    try handle.write(contentsOf: bytes + Data("\n".utf8))
    try handle.synchronize()
    return true
  } catch { return false }
}

func diagnosticSummary(saved: Bool) -> String {
  saved ? "本机启动阶段与退出记录已保存。" : "本机诊断未完整保存，请检查目录权限与磁盘空间。"
}

// Process is the exact child handle created by this invocation. Never kill by name/port.
func stopOwnedLaunchChild(_ child: Process, grace: TimeInterval = 3) {
  if !child.isRunning { return }
  child.terminate()
  let deadline = Date().addingTimeInterval(grace)
  while child.isRunning && Date() < deadline { Thread.sleep(forTimeInterval: 0.05) }
  if child.isRunning { _ = kill(child.processIdentifier, SIGKILL) }
  child.waitUntilExit()
}

func monitorLaunch(
  child: Process, attempt: LaunchAttempt, timeout: TimeInterval = 30,
  pollInterval: TimeInterval = 0.25, terminateGrace: TimeInterval = 3,
  heartbeatTimeout: TimeInterval = 30,
  hostActivity: LaunchHostActivity = LaunchHostActivity(),
  onHealth: (String) -> Void = { _ in },
  onEvent: (String, String?) -> Void = { _, _ in }
) -> LaunchOutcome {
  let deadline = Date().addingTimeInterval(timeout)
  var confirmed = false
  var status: LaunchStatus?
  var timedOut = false
  while child.isRunning {
    if let current = readLaunchStatus(attempt, pid: child.processIdentifier) { status = current }
    if status?.desktopVisible == true {
      confirmed = true
      onEvent("desktop-visible", nil)
      break
    }
    if Date() >= deadline {
      timedOut = true
      onEvent("startup-timeout", "LAUNCH_READY_TIMEOUT")
      stopOwnedLaunchChild(child, grace: terminateGrace)
      break
    }
    Thread.sleep(forTimeInterval: pollInterval)
  }
  var watch = LaunchHeartbeatWatch()
  while child.isRunning && confirmed {
    let current = readLaunchStatus(attempt, pid: child.processIdentifier)
    if let current = current { status = current }
    let activity = hostActivity.snapshot()
    if status?.phase != "quitting",
      let event = watch.sample(
        beat: current?.heartbeat, now: ProcessInfo.processInfo.systemUptime,
        timeout: heartbeatTimeout, sleepGeneration: activity.generation,
        suspended: activity.suspended)
    {
      onEvent(
        event, event == "main-unresponsive-suspected" ? "DESKTOP_MAIN_UNRESPONSIVE_SUSPECTED" : nil)
      onHealth(event)
    }
    Thread.sleep(forTimeInterval: max(0.02, min(2, pollInterval * 8)))
  }
  child.waitUntilExit()
  if let current = readLaunchStatus(attempt, pid: child.processIdentifier) { status = current }
  if !confirmed && status?.desktopVisible == true {
    confirmed = true
    onEvent("desktop-visible", nil)
  }
  let duplicate = status?.phase == "duplicate" && child.terminationStatus == 0
  let signal = child.terminationReason == .uncaughtSignal ? child.terminationStatus : nil
  let manualQuit =
    confirmed && status?.phase == "quitting" && signal == nil
    && child.terminationStatus == 0
  let failureCode: String? =
    timedOut
    ? "LAUNCH_READY_TIMEOUT"
    : (!confirmed && !duplicate
      ? (safeDiagnosticCode(status?.failureCode) ?? "LAUNCH_EXIT_BEFORE_READY")
      : (!manualQuit && !duplicate ? "LAUNCH_MAIN_EXIT_UNEXPECTED" : nil))
  return LaunchOutcome(
    startupConfirmed: confirmed, duplicate: duplicate, manualQuit: manualQuit,
    exitCode: child.terminationStatus, signal: signal, failureCode: failureCode,
    lastPhase: status?.phase ?? "before-js-marker")
}
