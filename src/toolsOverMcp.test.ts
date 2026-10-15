import { afterEach, describe, expect, it, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "./server.js";
import { requestContext } from "./auth.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, { startIndex: 1, endIndex: 7, paragraph: { elements: [{ startIndex: 1, endIndex: 7, textRun: { content: "Title\n" } }], paragraphStyle: { namedStyleType: "HEADING_1" } } }] },
};

function json(data: unknown) {
  return new Response(JSON.stringify(data), { status: 200, headers: { "content-type": "application/json" } });
}

async function getDocument(args: Record<string, unknown>) {
  const fetchMock = vi.fn(async (input: string | URL | Request) => {
    const url = String(input);
    if (url.includes("/drive/v3/files/doc1?")) return json({ name: "Doc", mimeType: "application/vnd.google-apps.document", webViewLink: "https://docs.google.com/document/d/doc1" });
    if (url.includes("/export?mimeType=text%2Fmarkdown")) return new Response("# Title\n\nSome **bold** text 🚀", { status: 200 });
    if (url.includes("/export?mimeType=text%2Fplain")) return new Response("Title\r\nSome bold text 🚀", { status: 200 });
    if (url.includes("docs.googleapis.com/v1/documents/doc1")) return json(DOC);
    throw new Error(`unexpected fetch ${url}`);
  });
  vi.stubGlobal("fetch", fetchMock);
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0.0.0" });
  await createServer(null).connect(serverTransport);
  await client.connect(clientTransport);
  vi.spyOn(process.stdout, "write").mockImplementation(() => true);
  const result = await requestContext.run({ accessToken: "tok" } as any, () =>
    client.callTool({ name: "get_document", arguments: { document_id: "doc1", ...args } }),
  );
  return { result, urls: fetchMock.mock.calls.map(([u]) => String(u)) };
}

describe("get_document format", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("exports markdown when asked and says so", async () => {
    const { result, urls } = await getDocument({ format: "markdown" });
    expect(result.isError).toBeFalsy();
    const output = result.structuredContent as any;
    expect(output.content).toBe("# Title\n\nSome **bold** text 🚀");
    expect(output.format).toBe("markdown");
    expect(urls.some((u) => u.includes("mimeType=text%2Fmarkdown"))).toBe(true);
  });

  it("defaults to plain text", async () => {
    const { result } = await getDocument({});
    const output = result.structuredContent as any;
    expect(output.content).toBe("Title\r\nSome bold text 🚀");
    expect(output.format).toBeUndefined();
  });

  it("still returns structure indices alongside markdown content", async () => {
    const { result } = await getDocument({ format: "markdown", include_structure: true });
    const output = result.structuredContent as any;
    expect(output.content).toBe("# Title\n\nSome **bold** text 🚀");
    expect(output.headings).toEqual([{ level: 1, text: "Title", startIndex: 1, endIndex: 7 }]);
  });

  it("refuses markdown together with a tab or comments, naming the conflict", async () => {
    const tab = await getDocument({ format: "markdown", tab_id: "t.0" });
    expect(tab.result.isError).toBe(true);
    expect((tab.result.content as any)[0].text).toMatch(/tab_id/);
    const comments = await getDocument({ format: "markdown", include_comments: true });
    expect(comments.result.isError).toBe(true);
    expect((comments.result.content as any)[0].text).toMatch(/include_comments/);
  });
});

describe("append_text", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
  });

  it("starts a new paragraph and resets the style it inherited from a heading", async () => {
    const calls: any[] = [];
    vi.stubGlobal("fetch", vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = String(input);
      if (url.includes("/drive/v3/files/doc1?")) return json({ name: "Doc", mimeType: "application/vnd.google-apps.document" });
      if (url.endsWith(":batchUpdate")) { calls.push(JSON.parse(String(init?.body))); return json({}); }
      if (url.includes("docs.googleapis.com/v1/documents/doc1")) return json(DOC);
      throw new Error(`unexpected fetch ${url}`);
    }));
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    const client = new Client({ name: "test", version: "0.0.0" });
    await createServer(null).connect(serverTransport);
    await client.connect(clientTransport);
    vi.spyOn(process.stdout, "write").mockImplementation(() => true);
    const result = await requestContext.run({ accessToken: "tok" } as any, () =>
      client.callTool({ name: "append_text", arguments: { document_id: "doc1", text: "Body" } }),
    );
    expect(result.isError).toBeFalsy();
    expect(calls).toHaveLength(1);
    const [insert, chars, style, bullets] = calls[0].requests;
    expect(insert).toEqual({ insertText: { location: { index: 6 }, text: "\nBody" } });
    expect(chars.updateTextStyle.range).toEqual({ startIndex: 6, endIndex: 11 });
    expect(style.updateParagraphStyle.range).toEqual({ startIndex: 7, endIndex: 11 });
    expect(bullets.deleteParagraphBullets.range).toEqual({ startIndex: 7, endIndex: 11 });
  });
});
