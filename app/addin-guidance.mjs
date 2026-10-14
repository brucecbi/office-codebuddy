// Installation copy must describe the host's entry point, not imply that a
// copied manifest, an OS-open result or a local service proves activation.
export function addinActivationInstructions(platform = process.platform) {
  const preparation =
    "Start the local service. After installing or updating a manifest, save your work, " +
    "fully quit and reopen Word / Excel. ";
  const boundary =
    "Manifest registration, pane connection and model availability are checked separately.";
  if (platform === "darwin")
    return (
      preparation +
      "In an editable file, choose Home → Add-ins → Developer Add-ins and single-click " +
      "Draftspect for Word / Excel. Once activated, use Home → 显示 Draftspect 窗格 in the target file. " +
      "Use Initialize Word / Excel Add-in only if manual menu activation still fails. " +
      boundary
    );
  if (platform === "win32")
    return (
      preparation +
      "Find the developer-registered add-in under Home → Add-ins (Insert → Add-ins on older builds). " +
      boundary
    );
  return "Automatic Office add-in installation is unsupported on this platform. See the README.";
}

export function addinInstallRecoveryInstructions(platform = process.platform) {
  const transport =
    "Use manifests rendered for the same local service origin; do not copy the HTTP templates " +
    "unchanged into an HTTPS installation. See the README's Sideload section. ";
  if (platform === "darwin")
    return (
      transport +
      "On macOS, register each manifest in that Office app's Data/Documents/wef/ folder, " +
      "then fully quit and reopen the host. Choose Home → Add-ins → Developer Add-ins " +
      "and single-click the Draftspect tile. An Upload My Add-in dialog is not required. " +
      "Do not clear Office caches or change certificate trust as an automatic recovery step."
    );
  if (platform === "win32")
    return (
      transport +
      "On Windows, register the manifest under HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer " +
      "with its full path as both the value name and data, then restart Office."
    );
  return "Manual registration on this platform has not been verified. See the README for supported hosts.";
}
