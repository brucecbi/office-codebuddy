// Process-local policy. This module reads no configuration files and starts nothing.
export function agentAccess(env = process.env) {
  const value = env.DRAFTSPECT_AGENT_ACCESS ?? "office-only";
  if (!["office-only", "extended"].includes(value)) {
    throw new Error("PREVIEW_ACCESS_INVALID: choose office-only or extended.");
  }
  return value;
}

export function codebuddyEnvironment(env = process.env) {
  const value = env.CODEBUDDY_INTERNET_ENVIRONMENT;
  // Values declared by the pinned SDK. Availability depends on the user's account.
  if (!["external", "internal", "ioa", "cloudhosted"].includes(value)) {
    throw new Error(
      "PREVIEW_ENV_REQUIRED: explicitly select CODEBUDDY_INTERNET_ENVIRONMENT for your account; no account environment was assumed.",
    );
  }
  return value;
}

export function userSettingsSources(env = process.env) {
  if (agentAccess(env) === "office-only") return [];
  const value = env.DRAFTSPECT_CODEBUDDY_SETTING_SOURCES ?? "none";
  if (!["none", "user"].includes(value))
    throw new Error("PREVIEW_SETTINGS_INVALID: choose none or user.");
  return value === "user" ? ["user"] : [];
}

export function forwardUserMcp(env = process.env) {
  if (agentAccess(env) === "office-only") return false;
  const value = env.DRAFTSPECT_FORWARD_USER_MCP ?? "0";
  if (!["0", "1"].includes(value)) throw new Error("PREVIEW_MCP_INVALID: choose 0 or 1.");
  return value === "1";
}

export function previewSdkOptions(options, env = process.env) {
  const settings = userSettingsSources(env);
  if (agentAccess(env) === "extended")
    return {
      ...options,
      settingSources: settings,
      strictMcpConfig: true,
      ...(settings.length ? {} : { hooks: {} }),
    };
  const office = options.mcpServers?.office;
  return {
    ...options,
    systemPrompt:
      typeof options.systemPrompt === "string"
        ? options.systemPrompt +
          "\nDeveloper Preview office-only mode: use only the supplied Office MCP tools; filesystem, shell, skills and external MCP tools are unavailable."
        : {
            ...options.systemPrompt,
            append:
              (options.systemPrompt?.append || "") +
              "\nDeveloper Preview office-only mode: use only the supplied Office MCP tools; filesystem, shell, skills and external MCP tools are unavailable.",
          },
    settingSources: [],
    strictMcpConfig: true,
    tools: [],
    allowedTools: [],
    hooks: {},
    agents: {},
    extraArgs: {},
    mcpServers: office ? { office } : {},
    canUseTool: (name, input, context) =>
      /^mcp__office__(?:office|excel)_[a-z0-9_]+$/.test(name)
        ? (options.canUseTool?.(name, input, context) ??
          Promise.resolve({
            behavior: "deny",
            message: "Office permission handler is missing.",
          }))
        : Promise.resolve({
            behavior: "deny",
            message: "Developer Preview exposes Office MCP tools only by default.",
          }),
  };
}
