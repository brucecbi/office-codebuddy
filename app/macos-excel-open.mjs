// Open the private initialization workbook in the intended Office host, not
// the user's default XLSX application. Do not change LaunchServices settings.
import { spawn } from "node:child_process";
import { isAbsolute } from "node:path";

export async function openInMicrosoftExcel(
  path,
  { platform = process.platform, spawnImpl = spawn } = {},
) {
  if (platform !== "darwin") return "Excel initialization is only verified on macOS.";
  if (typeof path !== "string" || !isAbsolute(path) || /[\x00-\x1f\x7f]/.test(path))
    return "The initialization workbook path is invalid.";
  return new Promise((resolve) => {
    try {
      const child = spawnImpl("/usr/bin/open", ["-b", "com.microsoft.Excel", path], {
        stdio: "ignore",
        timeout: 10000,
        shell: false,
      });
      child.once("error", () =>
        resolve("Microsoft Excel could not open the initialization workbook."),
      );
      child.once("close", (code, signal) =>
        resolve(
          code === 0 && !signal
            ? ""
            : "Microsoft Excel could not open the initialization workbook.",
        ),
      );
    } catch {
      resolve("Microsoft Excel could not open the initialization workbook.");
    }
  });
}
