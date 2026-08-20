import express, { type Request, type Response } from "express";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { createServer } from "./server.js";
import { requestContext } from "./auth.js";

const PORT = Number(process.env.PORT) || 8000;
const MCP_PATH = "/mcp";

const app = express();
app.use(express.json({ limit: "10mb" }));

app.get("/healthz", (_req, res) => {
  res.json({ status: "ok" });
});
app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

app.post(MCP_PATH, async (req: Request, res: Response) => {
  const authHeader = req.header("authorization") ?? req.header("Authorization") ?? "";
  const accessToken = authHeader.startsWith("Bearer ")
    ? authHeader.slice("Bearer ".length).trim()
    : "";

  // A fresh McpServer AND transport per request. The MCP Protocol owns one
  // transport at a time, so a shared module-level server throws
  // "Already connected to a transport" as soon as two tool calls overlap.
  const server = createServer();
  const transport = new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined,
  });

  requestContext.run({ accessToken }, async () => {
    try {
      res.on("close", () => {
        transport.close().catch(() => {});
        server.close().catch(() => {});
      });
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (err) {
      console.error("[gdocs-hosted] MCP request error:", err);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: {
            code: -32603,
            message: err instanceof Error ? err.message : "Internal error",
          },
          id: null,
        });
      }
    }
  });
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`[gdocs-hosted] listening on 0.0.0.0:${PORT}${MCP_PATH}`);
});
