import { homedir } from "node:os";
import { join, resolve } from "node:path";

// This fork has independent state. Ports are shared: run one installation at a time.
export function stateDirectory(env = process.env) {
  const dir = resolve(env.DRAFTSPECT_STATE_DIR || join(homedir(), ".draftspect-preview"));
  if (
    [resolve(homedir(), ".claude", "office-addins"), resolve(homedir(), ".draftspect-cb")].some(
      (protectedRoot) => dir === protectedRoot || dir.startsWith(protectedRoot + "/"),
    )
  ) {
    throw new Error(
      "DRAFTSPECT_STATE_DIR must not point at an existing Draftspect installation's state",
    );
  }
  return dir;
}
export const stateFile = (name, env = process.env) => join(stateDirectory(env), name);
