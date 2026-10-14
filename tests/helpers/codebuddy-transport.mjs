import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { pathToFileURL } from "node:url";

// Resolve the pinned package entry, including pnpm/hoisted layouts. Transport
// internals are intentionally tested: they are not part of the SDK's exports.
// Missing package is a declared skip; installed package/API drift must fail.
export async function loadProcessTransport() {
  const require = createRequire(import.meta.url);
  let entry;
  try {
    entry = require.resolve("@tencent-ai/agent-sdk");
  } catch (error) {
    if (
      error.code === "MODULE_NOT_FOUND" &&
      error.message.split("\n")[0] === "Cannot find module '@tencent-ai/agent-sdk'"
    )
      return null;
    throw error;
  }
  const file = require.resolve(join(dirname(entry), "transport", "process-transport.js"));
  const transport = (await import(pathToFileURL(file).href)).ProcessTransport;
  if (typeof transport !== "function")
    throw new Error("Pinned SDK ProcessTransport API unavailable");
  return transport;
}
