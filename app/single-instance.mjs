import { chmodSync, lstatSync, mkdirSync, realpathSync } from "node:fs";
import { join, resolve } from "node:path";
import { stateDirectory } from "../daemon/state-paths.mjs";

// Call once, before whenReady() or any application lifecycle listeners. Using
// the state directory (not the checkout or generic Electron profile) gives
// every launcher for this installation the same process-level lock identity.
export function acquireDesktopInstance({ app, stateDir, onSecondInstance = () => {} }) {
  let ownsLock = false;
  try {
    if (typeof onSecondInstance !== "function") {
      throw new TypeError("onSecondInstance must be a function");
    }
    if (stateDir !== undefined && (typeof stateDir !== "string" || !stateDir.trim())) {
      throw new TypeError("stateDir must be a non-empty path");
    }
    const root = stateDir === undefined ? stateDirectory() : resolve(stateDir);
    mkdirSync(root, { recursive: true, mode: 0o700 });
    const userData = join(realpathSync(root), "desktop");
    mkdirSync(userData, { recursive: true, mode: 0o700 });
    // Do not redirect Electron state or chmod an unrelated directory through
    // a pre-existing symlink. Tighten an existing real profile as well.
    const profile = lstatSync(userData);
    if (!profile.isDirectory() || profile.isSymbolicLink()) {
      throw new Error("Draftspect desktop profile must be a real directory");
    }
    chmodSync(userData, 0o700);
    app.setName("Draftspect");
    app.setPath("userData", userData);
    ownsLock = app.requestSingleInstanceLock() === true;
    if (ownsLock) {
      app.on("second-instance", onSecondInstance);
      return true;
    }
  } catch (error) {
    if (ownsLock) {
      // Registration failure must not leave an apparently usable owner behind.
      try {
        app.removeListener("second-instance", onSecondInstance);
      } catch {
        // Still attempt to release ownership if listener cleanup fails.
      }
      try {
        app.releaseSingleInstanceLock();
      } catch {
        // The caller must quit after reporting initialization failure.
      }
    }
    const failure = new Error(
      `Draftspect single-instance initialization failed: ${error?.message || "unknown error"}`,
      { cause: error },
    );
    failure.code = "DESKTOP_SINGLE_INSTANCE_FAILED";
    throw failure;
  }
  try {
    app.quit();
  } catch {
    try {
      app.exit(1);
    } catch {
      // The caller still receives false and must not initialize the app.
    }
  }
  return false;
}
