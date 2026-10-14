// Read-only local transport preflight. No listeners, provider/runtime imports,
// native trust changes, config writes or certificate/private-key output.
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { loadLocalTransport } from "../daemon/local-transport.mjs";

export async function localTransportStatus(env = process.env, options = {}) {
  const transport = await loadLocalTransport(env, { ...options, warn: null });
  return {
    ok: true,
    source: transport.source,
    config_path: transport.configPath ?? null,
    protocol: transport.secure ? "https/wss" : "http/ws",
    origin: transport.origin,
    certificate: transport.certificate,
    warnings: transport.warnings,
    native_trust: "NOT_CHECKED",
    listeners_started: false,
    model_requests: 0,
  };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    let configPath;
    let requireHttps = false;
    const args = process.argv.slice(2);
    for (let i = 0; i < args.length; i++) {
      if (args[i] === "--config" && args[i + 1]) configPath = resolve(args[++i]);
      else if (args[i] === "--require-https") requireHttps = true;
      else
        throw new Error(
          "Usage: node scripts/local-transport-status.mjs [--config PATH] [--require-https]",
        );
    }
    const result = await localTransportStatus(process.env, { configPath });
    if (requireHttps && result.protocol !== "https/wss")
      throw new Error("HTTPS is required, but no TLS configuration was found");
    process.stdout.write(JSON.stringify(result, null, 2) + "\n");
  } catch (error) {
    process.stderr.write(JSON.stringify({ ok: false, error: error.message }) + "\n");
    process.exitCode = 1;
  }
}
