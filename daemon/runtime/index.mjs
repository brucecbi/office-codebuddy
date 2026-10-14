// Runtime selector. DRAFTSPECT_RUNTIME=codebuddy (default) | claude.
//
// Contract every runtime module implements (see runtime-claude.mjs for
// the reference implementation):
//
//   id: string                    stable provider key — namespaces the
//                                 session store (sessions.json) so Claude
//                                 and CodeBuddy resume ids never collide
//   supportsStderr: boolean       whether onStderr is ever called
//   mcp: {createSdkMcpServer, tool}  MCP registration primitives FROM THE
//                                 SAME SDK as the transport — office tools
//                                 must be built with them or the provider
//                                 rejects requests (400)
//   knownModels: string[]|null    synchronous model catalog for validation
//                                 (claude: the fixed legacy provider list; other
//                                 runtimes: null — discovery-based, the
//                                 daemon caches listModels() results)
//   probe()                       → Promise<{ok, detail}> — availability
//                                 check for diagnostics (never throws)
//   listModels()                  → Promise<Array|null> — dynamic model
//                                 catalog, or null when the transport has
//                                 none (Phase 4 wires this into the pane)
//   runSession(params)            → { cancel(), close() } — starts one
//                                 pane's agent loop (detached; reports
//                                 through callbacks, never throws)
//
//   runSession params:
//     cwd, resumeId, model, systemPromptAppend, mcpServers,
//     disallowedTools, canUseTool, includePartialMessages,
//     turnSource            pull-based user-message source (turns.mjs)
//     isStale()             true once this pane's loop was superseded
//     onMessage(msg)        one SDK message (shared handleAgentMessage)
//     onStderr(data)        Claude-only stderr stream
//     onTurnEnd({sawResult}) per-turn terminal note (CodeBuddy)
//     onEnd({reason, error}) reason: "ended" | "aborted" | "error"

import { acceptanceProfile } from "../acceptance-profile.mjs";

// Import only the selected provider; selection alone must not initialize both SDKs.
const RUNTIMES = { claude: "./runtime-claude.mjs", codebuddy: "./runtime-codebuddy.mjs" };
const requested = String(
  acceptanceProfile ? "codebuddy" : process.env.DRAFTSPECT_RUNTIME || "codebuddy",
)
  .trim()
  .toLowerCase();
if (!Object.hasOwn(RUNTIMES, requested)) {
  throw new Error(`Unknown DRAFTSPECT_RUNTIME="${requested}". Expected "codebuddy" or "claude".`);
}
const runtime = await import(RUNTIMES[requested]);

export default runtime;
export { RUNTIMES };
