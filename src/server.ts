import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { GoogleDocsTools } from "./tools.js";
import { grantedScopes, isToolGranted } from "./scopes.js";

const SERVER_NAME = "Google Docs";
const SERVER_VERSION = "0.1.0";

// Built once: createServer runs per request, and rebuilding every Zod shape
// each time would put that cost on every call
const TOOLS = Object.entries(GoogleDocsTools.getTools()) as Array<[string, any]>;

// Resolved at import so a bad PROFILE fails the boot instead of every request
const GRANTED = grantedScopes();

const registered = TOOLS.filter(([, t]) => isToolGranted(t.handler?.scope, GRANTED)).map(([name]) => name);
const withheld = TOOLS.map(([name]) => name).filter((name) => !registered.includes(name));
console.log(
  `[gdocs-hosted] scopes=${GRANTED === null ? "unrestricted" : [...GRANTED].join(",")}`,
);
console.log(`[gdocs-hosted] tools=${registered.join(",") || "(none)"}`);
if (withheld.length > 0) {
  console.log(`[gdocs-hosted] withheld (scope not granted)=${withheld.join(",")}`);
}

export function createServer(granted: Set<string> | null = GRANTED): McpServer {
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });

  for (const [toolName, t] of TOOLS) {
    if (!isToolGranted(t.handler?.scope, granted)) continue;

    server.registerTool(
      toolName,
      {
        description: t.description,
        inputSchema: t.schema,
        outputSchema: t.outputSchema,
        annotations: {
          readOnlyHint: t.readOnlyHint ?? false,
          destructiveHint: t.destructiveHint ?? false,
        },
      },
      async (args: Record<string, unknown>) => t.handler(args),
    );
  }

  return server;
}
