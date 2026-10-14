// Explicit file-only setup. Never installs certificate trust or starts services.
import { mkdir, mkdtemp, lstat, chmod, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { stateDirectory } from "../daemon/state-paths.mjs";
import { loadLocalTransport } from "../daemon/local-transport.mjs";

export async function createPreviewTls({ env = process.env } = {}) {
  const root = stateDirectory(env);
  const configPath = join(root, "transport.json");
  try {
    await lstat(configPath);
    throw new Error(
      "PREVIEW_TLS_EXISTS: existing configuration is preserved; renewal needs a separate reviewed switch.",
    );
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
  }
  await mkdir(root, { recursive: true, mode: 0o700 });
  const info = await lstat(root);
  if (!info.isDirectory() || info.isSymbolicLink() || info.mode & 0o077)
    throw new Error("PREVIEW_TLS_DIRECTORY: use a private real directory with mode 0700.");
  const tlsDir = await mkdtemp(join(root, "tls-"));
  const keyPath = join(tlsDir, "loopback-key.pem"),
    certPath = join(tlsDir, "loopback-cert.pem");
  const result = spawnSync(
    "openssl",
    [
      "req",
      "-x509",
      "-newkey",
      "rsa:2048",
      "-nodes",
      "-days",
      "90",
      "-subj",
      "/CN=Draftspect Preview loopback",
      "-addext",
      "subjectAltName=IP:127.0.0.1",
      "-addext",
      "basicConstraints=critical,CA:FALSE",
      "-addext",
      "extendedKeyUsage=serverAuth",
      "-keyout",
      keyPath,
      "-out",
      certPath,
    ],
    { stdio: "ignore", timeout: 20000, env },
  );
  if (result.error || result.status !== 0)
    throw new Error(
      "PREVIEW_TLS_GENERATION: OpenSSL with req -addext is required; no trust/configuration changed. A private partial TLS directory may remain.",
    );
  await chmod(keyPath, 0o600);
  const settings = await loadLocalTransport(
    { ...env, DRAFTSPECT_TLS_KEY: keyPath, DRAFTSPECT_TLS_CERT: certPath },
    { warn: null },
  );
  await writeFile(
    configPath,
    JSON.stringify({ version: 1, protocol: "https", keyPath, certPath }, null, 2) + "\n",
    { flag: "wx", mode: 0o600 },
  );
  return {
    configPath,
    certificatePath: certPath,
    certificate: settings.certificate,
    trustInstalled: false,
    listenersStarted: false,
  };
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  if (process.argv.slice(2).join(" ") !== "--create") {
    console.error(
      "Usage: node scripts/setup-preview-tls.mjs --create (creates private files only; never installs trust)",
    );
    process.exitCode = 2;
  } else {
    try {
      console.log(JSON.stringify(await createPreviewTls(), null, 2));
    } catch (error) {
      console.error(error.message);
      process.exitCode = 1;
    }
  }
}
