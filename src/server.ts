import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { GoogleDocsTools } from "./tools.js";
import { grantedScopes, isToolGranted } from "./scopes.js";

const SERVER_NAME = "Google Docs";
const SERVER_VERSION = "0.1.0";

export function createServer(granted = grantedScopes()): McpServer {
  const server = new McpServer({ name: SERVER_NAME, version: SERVER_VERSION });

  const tools = GoogleDocsTools.getTools();
  const registered: string[] = [];
  const skipped: string[] = [];

  for (const [toolName, toolConfig] of Object.entries(tools)) {
    const t = toolConfig as any;

    if (!isToolGranted(t.handler?.scope, granted)) {
      skipped.push(toolName);
      continue;
    }

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
    registered.push(toolName);
  }

  console.log(
    `[gdocs-hosted] scopes=${granted === null ? "unrestricted" : [...granted].join(",")}`,
  );
  console.log(`[gdocs-hosted] tools=${registered.join(",") || "(none)"}`);
  if (skipped.length > 0) {
    console.log(`[gdocs-hosted] withheld (scope not granted)=${skipped.join(",")}`);
  }

  return server;
}
