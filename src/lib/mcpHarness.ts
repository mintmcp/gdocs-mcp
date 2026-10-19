import { expect, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "../server.js";
import { requestContext } from "../auth.js";
import { jsonResponse, stubFetch } from "./labelStubs.js";

export const DOC_METADATA = { name: "Doc", mimeType: "application/vnd.google-apps.document", webViewLink: "https://docs.google.com/document/d/doc1" };

export const paragraph = (text: string, startIndex: number, extra: Record<string, unknown> = {}) => ({
  startIndex,
  endIndex: startIndex + text.length,
  paragraph: { elements: [{ startIndex, endIndex: startIndex + text.length, textRun: { content: text } }], ...extra },
});

/** Calls one tool over an in-memory MCP transport with a fake access token. */
export async function callTool(name: string, args: Record<string, unknown>) {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0.0.0" });
  await createServer(null).connect(serverTransport);
  await client.connect(clientTransport);
  vi.spyOn(process.stdout, "write").mockImplementation(() => true);
  return requestContext.run({ accessToken: "tok" } as any, () =>
    client.callTool({ name, arguments: { document_id: "doc1", ...args } }),
  );
}

/** Stubs the Drive metadata read and text export, the Docs document read and every batchUpdate for doc1. */
export function stubDoc(doc: unknown, batchReply: unknown = {}) {
  return stubFetch([
    ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
    ["/export?mimeType=", () => new Response("exported text", { status: 200 })],
    [":batchUpdate", () => jsonResponse(batchReply)],
    ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(doc)],
  ]);
}

export function batchBodies(calls: Array<{ url: string; init?: RequestInit }>) {
  return calls.filter(({ url }) => url.endsWith(":batchUpdate")).map(({ init }) => JSON.parse(String(init?.body)));
}

/** The single batchUpdate body a call sent. */
export function batchBody(calls: Array<{ url: string; init?: RequestInit }>) {
  const bodies = batchBodies(calls);
  expect(bodies).toHaveLength(1);
  return bodies[0];
}

export function errorText(result: any): string {
  expect(result.isError).toBe(true);
  return JSON.parse(result.content[0].text).error;
}
