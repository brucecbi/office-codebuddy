import AppKit
import Darwin
import Foundation

// Local-only launch supervisor. It never reads credentials or starts a model request.
struct LauncherConfiguration: Decodable {
  let schemaVersion: Int
  let projectRoot: String
  let cliPath: String
  let nodePath: String
  let environment: [String: String]
}

struct LauncherFailure: Error, CustomStringConvertible {
  let description: String
}

let allowedEnvironment: Set<String> = [
  "DRAFTSPECT_STATE_DIR", "CODEBUDDY_CONFIG_DIR", "CODEBUDDY_INTERNET_ENVIRONMENT",
  "DRAFTSPECT_TLS_KEY", "DRAFTSPECT_TLS_CERT",
]

func hasControlCharacters(_ value: String) -> Bool {
  value.unicodeScalars.contains { $0.value < 0x20 || $0.value == 0x7f }
}

func requireAbsolutePath(_ path: String, label: String) throws {
  guard path.hasPrefix("/"), !hasControlCharacters(path) else {
    throw LauncherFailure(description: "\(label) 必须为不含控制字符的绝对路径。")
  }
}

func requireFile(_ path: String, label: String, executable: Bool = false) throws {
  try requireAbsolutePath(path, label: label)
  var isDirectory: ObjCBool = false
  guard FileManager.default.fileExists(atPath: path, isDirectory: &isDirectory),
    !isDirectory.boolValue,
    !executable || FileManager.default.isExecutableFile(atPath: path)
  else {
    throw LauncherFailure(description: "\(label) 不存在或不可用：\n\(path)\n请恢复原路径或重新构建启动器。")
  }
}

func readConfiguration() throws -> LauncherConfiguration {
  guard let configURL = Bundle.main.url(forResource: "launcher-config", withExtension: "json")
  else {
    throw LauncherFailure(description: "启动器配置丢失，请重新构建 Draftspect.app。")
  }
  let data = try Data(contentsOf: configURL)
  guard let raw = try JSONSerialization.jsonObject(with: data) as? [String: Any],
    Set(raw.keys) == Set(["schemaVersion", "projectRoot", "cliPath", "nodePath", "environment"])
  else {
    throw LauncherFailure(description: "启动器配置字段无效，请重新构建 Draftspect.app。")
  }
  let config = try JSONDecoder().decode(LauncherConfiguration.self, from: data)
  guard config.schemaVersion == 1,
    Set(config.environment.keys).isSubset(of: allowedEnvironment)
  else {
    throw LauncherFailure(description: "启动器配置版本或环境变量不受支持。")
  }
  guard let accountEnvironment = config.environment["CODEBUDDY_INTERNET_ENVIRONMENT"],
    ["external", "internal", "ioa", "cloudhosted"].contains(accountEnvironment) else {
    throw LauncherFailure(description: "请明确配置与你账户匹配的 CodeBuddy 网络环境；启动器不采用默认账户环境。")
  }
  for (key, value) in config.environment {
    guard !value.isEmpty, !hasControlCharacters(value) else {
      throw LauncherFailure(description: "启动器环境变量格式无效：\(key)。")
    }
    if key == "CODEBUDDY_INTERNET_ENVIRONMENT" {
      guard value.range(of: "^[A-Za-z0-9_-]{1,32}$", options: .regularExpression) != nil else {
        throw LauncherFailure(description: "CodeBuddy 网络环境名称无效。")
      }
    } else {
      try requireAbsolutePath(value, label: key)
    }
  }
  guard
    (config.environment["DRAFTSPECT_TLS_KEY"] == nil)
      == (config.environment["DRAFTSPECT_TLS_CERT"] == nil)
  else {
    throw LauncherFailure(description: "TLS 证书与私钥路径必须同时配置。")
  }
  try requireAbsolutePath(config.projectRoot, label: "项目路径")
  try requireAbsolutePath(config.cliPath, label: "CodeBuddy CLI 路径")
  try requireAbsolutePath(config.nodePath, label: "Node.js 路径")
  let root = URL(fileURLWithPath: config.projectRoot)
  try requireFile(root.appendingPathComponent("package.json").path, label: "项目 package.json")
  try requireFile(root.appendingPathComponent("app/main.mjs").path, label: "Draftspect 入口")
  try requireFile(
    root.appendingPathComponent("node_modules/electron/dist/Electron.app/Contents/MacOS/Electron")
      .path,
    label: "项目 Electron", executable: true)
  try requireFile(config.cliPath, label: "CodeBuddy CLI", executable: true)
  try requireFile(config.nodePath, label: "Node.js", executable: true)
  let package =
    try JSONSerialization.jsonObject(
      with: Data(contentsOf: root.appendingPathComponent("package.json"))) as? [String: Any]
  guard package?["name"] as? String == "draftspect-office-addins",
    package?["main"] as? String == "app/main.mjs"
  else {
    throw LauncherFailure(description: "项目路径未指向预期的 Draftspect 项目。")
  }
  for key in ["DRAFTSPECT_TLS_KEY", "DRAFTSPECT_TLS_CERT"] {
    if let path = config.environment[key] { try requireFile(path, label: key) }
  }
  return config
}

func userHomeDirectory() throws -> String {
  guard let record = getpwuid(getuid()), let home = record.pointee.pw_dir else {
    throw LauncherFailure(description: "无法确定当前 macOS 用户的主目录。")
  }
  return String(cString: home)
}

func launchEnvironment(_ config: LauncherConfiguration) throws -> [String: String] {
  // Foundation Process receives this exact dictionary, without inherited injection variables.
  var environment = config.environment
  environment["HOME"] = try userHomeDirectory()
  environment["USER"] = NSUserName()
  environment["TMPDIR"] = NSTemporaryDirectory()
  environment["LANG"] = "en_US.UTF-8"
  environment["DRAFTSPECT_RUNTIME"] = "codebuddy"
  environment["DRAFTSPECT_DESKTOP_LAUNCH"] = "1"
  environment["CODEBUDDY_CODE_PATH"] = config.cliPath
  environment["DRAFTSPECT_NODE_PATH"] = config.nodePath
  guard let accountEnvironment = config.environment["CODEBUDDY_INTERNET_ENVIRONMENT"],
    ["external", "internal", "ioa", "cloudhosted"].contains(accountEnvironment) else {
    throw LauncherFailure(description: "请在构建前明确选择 CodeBuddy 账号环境；启动器不会假定公司环境。")
  }
  environment["CODEBUDDY_INTERNET_ENVIRONMENT"] = accountEnvironment
  environment["PATH"] = [
    URL(fileURLWithPath: config.nodePath).deletingLastPathComponent().path,
    "/opt/homebrew/bin", "/usr/bin", "/bin", "/usr/sbin", "/sbin",
  ].joined(separator: ":")
  return environment
}

func openLaunchLog(_ config: LauncherConfiguration) throws -> FileHandle {
  let defaultDirectory = URL(fileURLWithPath: try userHomeDirectory())
    .appendingPathComponent(".draftspect-preview").path
  let directory =
    config.environment["DRAFTSPECT_STATE_DIR"]
    ?? defaultDirectory
  let path = URL(fileURLWithPath: directory).standardizedFileURL.path
  try FileManager.default.createDirectory(
    atPath: path, withIntermediateDirectories: true,
    attributes: [.posixPermissions: 0o700])
  let directoryFD = open(path, O_RDONLY | O_DIRECTORY | O_NOFOLLOW)
  guard directoryFD >= 0 else {
    throw LauncherFailure(description: "无法安全打开启动日志目录；请检查目录权限及符号链接。")
  }
  defer { close(directoryFD) }
  var directoryInfo = stat()
  guard fstat(directoryFD, &directoryInfo) == 0,
    directoryInfo.st_uid == getuid(), directoryInfo.st_mode & 0o022 == 0
  else {
    throw LauncherFailure(description: "启动日志目录必须由当前用户拥有，且不能允许其他用户写入。")
  }
  let descriptor = openat(
    directoryFD, "launcher.log", O_WRONLY | O_APPEND | O_CREAT | O_NOFOLLOW | O_NONBLOCK, 0o600)
  guard descriptor >= 0 else {
    throw LauncherFailure(description: "无法安全打开 launcher.log；请检查权限及符号链接。")
  }
  var info = stat()
  guard fstat(descriptor, &info) == 0, info.st_uid == getuid(),
    info.st_mode & mode_t(S_IFMT) == mode_t(S_IFREG), info.st_mode & 0o077 == 0,
    info.st_nlink == 1
  else {
    close(descriptor)
    throw LauncherFailure(description: "launcher.log 必须为当前用户拥有的私有普通文件，不能使用链接。")
  }
  return FileHandle(fileDescriptor: descriptor, closeOnDealloc: true)
}

@MainActor
func showFailure(_ message: String, projectRoot: String?, diagnosticDirectory: URL? = nil) {
  NSApplication.shared.setActivationPolicy(.accessory)
  NSApplication.shared.activate(ignoringOtherApps: true)
  let alert = NSAlert()
  alert.alertStyle = .critical
  alert.messageText = "Draftspect 需要处理"
  alert.informativeText = message
  alert.addButton(withTitle: "关闭")
  alert.addButton(
    withTitle: diagnosticDirectory != nil
      ? "查看诊断记录"
      : (projectRoot != nil ? "查看项目位置" : "查看启动器位置"))
  if alert.runModal() == .alertSecondButtonReturn {
    let target =
      diagnosticDirectory
      ?? projectRoot.map { URL(fileURLWithPath: $0) } ?? Bundle.main.bundleURL
    NSWorkspace.shared.activateFileViewerSelecting([target])
  }
}

// Keep the main AppKit loop responsive to Finder/NSWorkspace reopen events while
// one background monitor waits for its exact child. Reopen spawns only a normal
// single-instance contender; Electron owns the lock and forwards window focus.
@MainActor
final class LauncherController: NSObject, NSApplicationDelegate {
  let configuration: LauncherConfiguration
  var reopenInFlight = false
  let hostActivity = LaunchHostActivity()
  var healthWindow: NSWindow?
  var healthDirectory: URL?
  var sleepObservers: [NSObjectProtocol] = []

  init(configuration: LauncherConfiguration) {
    self.configuration = configuration
    super.init()
    let center = NSWorkspace.shared.notificationCenter
    let activity = hostActivity
    sleepObservers = [
      center.addObserver(forName: NSWorkspace.willSleepNotification, object: nil, queue: nil) { _ in
        activity.update(suspended: true)
      },
      center.addObserver(forName: NSWorkspace.didWakeNotification, object: nil, queue: nil) { _ in
        activity.update(suspended: false)
      },
    ]
  }

  func displayHealth(_ event: String, directory: URL) {
    guard event == "main-unresponsive-suspected" else {
      healthWindow?.close()
      healthWindow = nil
      return
    }
    guard healthWindow == nil else { return }
    healthDirectory = directory
    let window = NSWindow(
      contentRect: NSRect(x: 0, y: 0, width: 470, height: 200),
      styleMask: [.titled, .closable], backing: .buffered, defer: false)
    window.title = "Draftspect 桌面响应异常"
    window.isReleasedWhenClosed = false
    let text = NSTextField(
      wrappingLabelWithString:
        "桌面状态超过 30 秒未更新。主进程可能无响应，或本机状态记录不可用；尚未确认崩溃。\n\n可稍候观察或查看诊断记录。未自动终止、重启或重发任务；若正在编辑，请先核对文档实际结果。恢复响应后此提示会自动关闭。"
    )
    text.frame = NSRect(x: 20, y: 62, width: 430, height: 115)
    window.contentView?.addSubview(text)
    let button = NSButton(title: "查看诊断记录", target: self, action: #selector(openHealthRecords))
    button.frame = NSRect(x: 295, y: 16, width: 155, height: 30)
    window.contentView?.addSubview(button)
    window.center()
    window.makeKeyAndOrderFront(nil)
    NSApplication.shared.activate(ignoringOtherApps: true)
    healthWindow = window
  }

  @objc func openHealthRecords() {
    if let directory = healthDirectory {
      NSWorkspace.shared.activateFileViewerSelecting([directory])
    }
  }

  func applicationShouldHandleReopen(_ sender: NSApplication, hasVisibleWindows: Bool) -> Bool {
    if !reopenInFlight {
      reopenInFlight = true
      start(primary: false)
    }
    return false
  }

  func start(primary: Bool) {
    let config = configuration
    let activity = hostActivity
    DispatchQueue.global(qos: .utility).async {
      var diagnosticDirectory: URL?
      var failure: String?
      do {
        try superviseLaunch(
          config, diagnosticDirectory: &diagnosticDirectory,
          hostActivity: activity,
          onHealth: { event, directory in
            DispatchQueue.main.async {
              if primary { self.displayHealth(event, directory: directory) }
            }
          })
      } catch {
        failure = (error as? LauncherFailure)?.description ?? "无法准备本机启动状态，请检查私有目录权限。"
        writeLaunchDiagnostic(
          FileHandle.standardError, event: "launcher-failed", pid: 0,
          code: "LAUNCH_FAILED")
      }
      let resultMessage = failure
      let resultDirectory = diagnosticDirectory
      DispatchQueue.main.async {
        if primary {
          self.healthWindow?.close()
          self.healthWindow = nil
        }
        self.reopenInFlight = false
        if let message = resultMessage {
          showFailure(
            message, projectRoot: config.projectRoot, diagnosticDirectory: resultDirectory)
        }
        if primary { exit(resultMessage == nil ? 0 : 1) }
      }
    }
  }
}

func superviseLaunch(
  _ config: LauncherConfiguration, diagnosticDirectory: inout URL?,
  hostActivity: LaunchHostActivity, onHealth: @escaping (String, URL) -> Void
) throws {
  let log = try openLaunchLog(config)
  defer { try? log.close() }
  var diagnosticsSaved = true
  func record(
    _ event: String, pid: Int32, code: String? = nil, exitCode: Int32? = nil, signal: Int32? = nil,
    versions: [String: String] = [:]
  ) {
    let saved = writeLaunchDiagnostic(
      log, event: event, pid: pid, code: code, exitCode: exitCode, signal: signal,
      versions: versions)
    diagnosticsSaved = saved && diagnosticsSaved
  }
  let defaultStateDirectory = URL(fileURLWithPath: try userHomeDirectory()).appendingPathComponent(
    ".draftspect-preview"
  ).path
  let stateDirectory = URL(
    fileURLWithPath: config.environment["DRAFTSPECT_STATE_DIR"] ?? defaultStateDirectory)
  let attempt: LaunchAttempt
  do {
    attempt = try createLaunchAttempt(stateDirectory: stateDirectory)
  } catch {
    record("launch-preparation-failed", pid: 0, code: "LAUNCH_STATE_UNAVAILABLE")
    throw LauncherFailure(description: "无法建立私有启动握手目录。请检查本机状态目录权限；不会启动或重试 Electron。")
  }
  diagnosticDirectory = attempt.directory
  let child = Process()
  let electron = URL(fileURLWithPath: config.projectRoot).appendingPathComponent(
    "node_modules/electron/dist/Electron.app/Contents/MacOS/Electron")
  child.executableURL = electron
  child.arguments = [config.projectRoot]
  child.currentDirectoryURL = URL(fileURLWithPath: config.projectRoot)
  child.environment = try launchEnvironment(config)
  child.environment?["DRAFTSPECT_LAUNCH_STATUS_PATH"] = attempt.statusURL.path
  child.environment?["DRAFTSPECT_LAUNCH_TOKEN"] = attempt.token
  child.environment?["DRAFTSPECT_LAUNCH_PARENT_PID"] = String(getpid())
  child.standardInput = FileHandle.nullDevice
  // Metadata-only launcher records; do not copy arbitrary child output into diagnostics.
  child.standardOutput = FileHandle.nullDevice
  child.standardError = FileHandle.nullDevice
  record(
    "launch-requested", pid: 0,
    versions: launchVersions(projectRoot: config.projectRoot))
  do {
    try child.run()
  } catch {
    record("spawn-failed", pid: 0, code: "LAUNCH_SPAWN_FAILED")
    throw LauncherFailure(
      description: "macOS 未能启动项目内的 Electron。请检查安装完整性及系统安全提示；不要关闭系统安全保护。")
  }
  record("spawned", pid: child.processIdentifier)
  let outcome = monitorLaunch(
    child: child, attempt: attempt, hostActivity: hostActivity,
    onHealth: { event in onHealth(event, attempt.directory) },
    onEvent: { event, code in
      record(event, pid: child.processIdentifier, code: code)
    })
  record(
    outcome.manualQuit
      ? "manual-quit" : (outcome.duplicate ? "duplicate-handoff" : "main-exit"),
    pid: child.processIdentifier, code: outcome.failureCode,
    exitCode: outcome.exitCode, signal: outcome.signal)
  let outcomeData = try JSONEncoder().encode(outcome)
  let outcomePath = attempt.directory.appendingPathComponent("outcome.json").path
  let descriptor = open(outcomePath, O_WRONLY | O_CREAT | O_EXCL | O_NOFOLLOW, 0o600)
  if descriptor >= 0 {
    let outcomeHandle = FileHandle(fileDescriptor: descriptor, closeOnDealloc: true)
    do {
      try outcomeHandle.write(contentsOf: outcomeData)
      try outcomeHandle.synchronize()
    } catch { diagnosticsSaved = false }
    try? outcomeHandle.close()
  } else {
    diagnosticsSaved = false
  }
  if !outcome.manualQuit && !outcome.duplicate {
    let stage = outcome.startupConfirmed ? "运行中" : "界面就绪前"
    let reason = outcome.signal.map { "信号 \($0)" } ?? "退出码 \(outcome.exitCode)"
    throw LauncherFailure(
      description:
        "Electron 在\(stage)退出（\(reason)）。\n" + diagnosticSummary(saved: diagnosticsSaved)
        + "未自动重启或重发任务。"
        + "\n请查看诊断记录并重新打开 Draftspect。如故障发生在编辑期间，先重新读取并核对文档中的实际结果。"
        + "\n系统保护和 renderer 沙箱保持开启。")
  }
  if !diagnosticsSaved {
    throw LauncherFailure(description: "Electron 已正常退出或完成单实例交接，但本机诊断未完整保存。请检查私有目录权限与磁盘空间；不会重发任务。")
  }
}

@main
struct DraftspectLauncher {
  @MainActor
  static func main() {
    let checkOnly = CommandLine.arguments.dropFirst().elementsEqual(["--check"])
    var configuration: LauncherConfiguration?
    do {
      let config = try readConfiguration()
      configuration = config
      if checkOnly {
        print("{\"ok\":true,\"schemaVersion\":1,\"mode\":\"check-only\"}")
        exit(0)
      }
      guard CommandLine.arguments.count == 1 else {
        throw LauncherFailure(description: "启动器不接受文档或自定义命令参数。")
      }
      let application = NSApplication.shared
      application.setActivationPolicy(.accessory)
      let controller = LauncherController(configuration: config)
      application.delegate = controller
      controller.start(primary: true)
      withExtendedLifetime(controller) { application.run() }
    } catch {
      let message = (error as? LauncherFailure)?.description ?? "无法读取启动器配置，请重新构建 Draftspect.app。"
      if checkOnly {
        let output = try! JSONSerialization.data(
          withJSONObject: ["ok": false, "error": message], options: [.sortedKeys])
        FileHandle.standardError.write(output)
        FileHandle.standardError.write(Data("\n".utf8))
      } else {
        // A failed preflight may have no safe writable state directory. Still emit
        // bounded metadata to the invoking console, never the invalid config value.
        writeLaunchDiagnostic(
          FileHandle.standardError, event: "launcher-failed", pid: 0,
          code: "LAUNCH_FAILED")
        showFailure(
          message, projectRoot: configuration?.projectRoot)
      }
      exit(1)
    }
  }
}
