// Read-only per-host manifest inspection. Matching files do not prove that an
// Office process loaded the add-in or that its task pane connected.
import { readFile } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { loadLocalTransport, renderManifestTransport } from "../daemon/local-transport.mjs";
import { previewManifestName } from "./preview-registration.mjs";

const ROOT = join(dirname(fileURLToPath(import.meta.url)), "..");
const HOSTS = ["word", "excel"];
const cleanXml = (xml) =>
  xml
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/>\s+</g, "><")
    .trim();
function identity(xml) {
  if (typeof xml !== "string" || /<!DOCTYPE|<!ENTITY/i.test(xml)) return null;
  const clean = cleanXml(xml);
  const one = (tag) => {
    const matches = [...clean.matchAll(new RegExp(`<${tag}>\\s*([^<]+?)\\s*</${tag}>`, "g"))];
    return matches.length === 1 ? matches[0][1] : null;
  };
  const id = one("Id"),
    version = one("Version"),
    permissions = one("Permissions");
  return id && version && permissions ? { id: id.toLowerCase(), version, permissions } : null;
}

export function inspectManifestRegistration({ projectXml, installedXml, transport }) {
  const expected = identity(projectXml),
    actual = identity(installedXml);
  if (!expected || !actual) return { state: "invalid", label: "manifest 无效或无法识别" };
  if (JSON.stringify(expected) !== JSON.stringify(actual))
    return { state: "mismatch", label: "GUID、版本或权限不匹配" };
  if (cleanXml(renderManifestTransport(projectXml, transport)) !== cleanXml(installedXml))
    return { state: "mismatch", label: "manifest 内容或服务地址需更新" };
  return { state: "matched", label: "manifest 匹配（尚不代表已激活）" };
}

export async function readAddinRegistration({
  projectRoot = ROOT,
  home = homedir(),
  platform = process.platform,
  transport,
} = {}) {
  if (platform !== "darwin")
    return HOSTS.map((host) => ({
      host,
      state: "unverified",
      label: "此平台未验证 manifest 匹配",
    }));
  const settings = transport || (await loadLocalTransport());
  return Promise.all(
    HOSTS.map(async (host) => {
      const path = join(
        home,
        "Library",
        "Containers",
        `com.microsoft.${host === "word" ? "Word" : "Excel"}`,
        "Data",
        "Documents",
        "wef",
        previewManifestName(host),
      );
      try {
        const [projectXml, installedXml] = await Promise.all([
          readFile(join(projectRoot, "manifests", `${host}.xml`), "utf8"),
          readFile(path, "utf8"),
        ]);
        return {
          host,
          ...inspectManifestRegistration({ projectXml, installedXml, transport: settings }),
        };
      } catch (error) {
        return {
          host,
          state: error.code === "ENOENT" ? "missing" : "unreadable",
          label: error.code === "ENOENT" ? "manifest 缺失" : "manifest 无法读取",
        };
      }
    }),
  );
}

export function officeConnectionDetail(registration, health, now = Date.now()) {
  // A metadata heartbeat covers only currently connected panes. It never proves
  // model availability, a specific user's document, or successful document IO.
  const fresh =
    health?.phase === "ready" &&
    Number.isFinite(health.receivedAt) &&
    now >= health.receivedAt &&
    now - health.receivedAt <= 35000;
  return HOSTS.map((host) => {
    const manifest = registration?.find((row) => row.host === host)?.label || "manifest 尚未检查";
    const count = health?.connected_panes_by_host?.[host];
    const connection =
      fresh && Number.isSafeInteger(count) && count >= 0
        ? count > 0
          ? `${count} 个侧边栏已连接（不代表模型可用）`
          : "尚无侧边栏连接"
        : "侧边栏连接尚未确认";
    return `${host === "word" ? "Word" : "Excel"}：${manifest}；${connection}`;
  }).join("\n");
}
