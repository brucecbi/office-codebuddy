import { execFile } from "node:child_process";
import { promisify } from "node:util";

const run = promisify(execFile);

// execFile has no `input`/`stdio` option. Close its pipe explicitly and bound
// fixture generation so an interactive OpenSSL/LibreSSL command cannot hang.
export function runOpenSsl(args) {
  const pending = run("openssl", args, { timeout: 20000 });
  pending.child.stdin.end();
  return pending;
}
