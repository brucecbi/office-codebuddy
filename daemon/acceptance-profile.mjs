// Explicit, process-local synthetic acceptance scope. No config writes, trust
// changes, provider imports or network calls. Normal production remains opt-out.
import { realpathSync, statSync } from "node:fs";
import { dirname, extname, isAbsolute, join, sep } from "node:path";
import { homedir } from "node:os";
import { fileURLToPath } from "node:url";

export function loadAcceptanceProfile(env = process.env) {
  const document = env.DRAFTSPECT_ACCEPTANCE_DOCUMENT;
  const model = env.DRAFTSPECT_ACCEPTANCE_MODEL;
  if (document === undefined && model === undefined) return null;
  if (!document || !model || !isAbsolute(document) || !/^[a-zA-Z0-9_.\[\]-]+$/.test(model))
    throw new Error("Acceptance requires an absolute document path and an explicit model ID");
  if (env.DRAFTSPECT_RUNTIME !== "codebuddy")
    throw new Error("Acceptance permits only the CodeBuddy runtime");
  const declaredState = env.DRAFTSPECT_STATE_DIR;
  if (!declaredState || !isAbsolute(declaredState))
    throw new Error("Acceptance requires an explicit, independent absolute DRAFTSPECT_STATE_DIR");
  const stateDir = realpathSync(declaredState);
  const activeDocument = realpathSync(document);
  if (
    !statSync(stateDir).isDirectory() ||
    !statSync(activeDocument).isFile() ||
    extname(activeDocument).toLowerCase() !== ".docx"
  )
    throw new Error("Acceptance requires an existing state directory and synthetic .docx file");
  const home = realpathSync(homedir());
  if (
    stateDir === home ||
    stateDir === dirname(activeDocument) ||
    [".draftspect-cb", ".codebuddy", ".claude"].some((name) => {
      const protectedRoot = join(home, name);
      return stateDir === protectedRoot || stateDir.startsWith(protectedRoot + sep);
    })
  )
    throw new Error("Acceptance state must be independent of production and provider state");
  return Object.freeze({
    document: activeDocument,
    workspace: dirname(activeDocument),
    stateDir,
    model,
  });
}

export const acceptanceProfile = loadAcceptanceProfile();

export function acceptanceAllowsDocument(host, value, profile = acceptanceProfile) {
  if (!profile) return true;
  try {
    const path =
      typeof value === "string" && value.startsWith("file:") ? fileURLToPath(value) : value;
    return (
      host === "word" &&
      typeof path === "string" &&
      isAbsolute(path) &&
      realpathSync(path) === profile.document
    );
  } catch {
    return false;
  }
}

export function assertAcceptanceWorkspace(cwd, profile = acceptanceProfile) {
  if (profile && (!isAbsolute(cwd || "") || realpathSync(cwd) !== profile.workspace))
    throw new Error("Acceptance workspace is fixed to the synthetic document's directory");
}

export function acceptanceSdkOptions(options, profile = acceptanceProfile) {
  if (!profile) return options;
  assertAcceptanceWorkspace(options.cwd, profile);
  if (options.model !== profile.model)
    throw new Error("Acceptance model must match the explicitly pinned model");
  const office = options.mcpServers?.office;
  if (
    !office ||
    office.type !== "sdk" ||
    Object.keys(options.mcpServers).some((name) => name !== "office")
  )
    throw new Error("Acceptance sessions require only the in-process Office MCP server");
  return {
    ...options,
    cwd: profile.workspace,
    model: profile.model,
    settingSources: [],
    strictMcpConfig: true,
    tools: [],
    allowedTools: [],
    extraArgs: {},
    maxTurns: 12,
    hooks: {},
    agents: {},
    mcpServers: { office },
    canUseTool: (name, input, context) =>
      /^(?:mcp__office__)?office_[a-z0-9_]+$/.test(name)
        ? (options.canUseTool?.(name, input, context) ??
          Promise.resolve({
            behavior: "deny",
            message: "Acceptance Office approval handler is missing",
          }))
        : Promise.resolve({
            behavior: "deny",
            message: "Acceptance exposes Office MCP tools only",
          }),
  };
}

export function acceptanceCliArguments(profile = acceptanceProfile) {
  return profile
    ? [
        "--model",
        profile.model,
        "--strict-mcp-config",
        "--mcp-config",
        JSON.stringify({ mcpServers: {} }),
        "--tools",
        "",
        "--setting-sources",
        "none",
        "--max-turns",
        "12",
      ]
    : [];
}
