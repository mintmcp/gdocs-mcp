import { afterEach, describe, expect, it, vi } from "vitest";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { createServer } from "./server.js";
import { requestContext } from "./auth.js";
import { jsonResponse, stubFetch } from "./lib/labelStubs.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, { startIndex: 1, endIndex: 7, paragraph: { elements: [{ startIndex: 1, endIndex: 7, textRun: { content: "Title\n" } }], paragraphStyle: { namedStyleType: "HEADING_1" } } }] },
};

const DOC_METADATA = { name: "Doc", mimeType: "application/vnd.google-apps.document", webViewLink: "https://docs.google.com/document/d/doc1" };

async function callTool(name: string, args: Record<string, unknown>) {
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "test", version: "0.0.0" });
  await createServer(null).connect(serverTransport);
  await client.connect(clientTransport);
  vi.spyOn(process.stdout, "write").mockImplementation(() => true);
  return requestContext.run({ accessToken: "tok" } as any, () =>
    client.callTool({ name, arguments: { document_id: "doc1", ...args } }),
  );
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("get_document format", () => {
  function stubDocument(markdown = "# Title\n\nSome **bold** text 🚀") {
    return stubFetch([
      ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
      ["/export?mimeType=text%2Fmarkdown", () => new Response(markdown, { status: 200 })],
      ["/export?mimeType=text%2Fplain", () => new Response("Title\r\nSome bold text 🚀", { status: 200 })],
      ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(DOC)],
    ]);
  }

  it("exports markdown when asked and says so", async () => {
    const calls = stubDocument();
    const result = await callTool("get_document", { format: "markdown" });
    expect(result.isError).toBeFalsy();
    const output = result.structuredContent as any;
    expect(output.content).toBe("# Title\n\nSome **bold** text 🚀");
    expect(output.format).toBe("markdown");
    expect(calls.some(({ url }) => url.includes("mimeType=text%2Fmarkdown"))).toBe(true);
  });

  it("defaults to plain text", async () => {
    stubDocument();
    const output = (await callTool("get_document", {})).structuredContent as any;
    expect(output.content).toBe("Title\r\nSome bold text 🚀");
    expect(output.format).toBeUndefined();
  });

  it("still returns structure indices alongside markdown content", async () => {
    stubDocument();
    const output = (await callTool("get_document", { format: "markdown", include_structure: true })).structuredContent as any;
    expect(output.content).toBe("# Title\n\nSome **bold** text 🚀");
    expect(output.headings).toEqual([{ level: 1, text: "Title", startIndex: 1, endIndex: 7 }]);
  });

  it("drops embedded images from markdown and points to get_document_images", async () => {
    stubDocument("![][image1]\n\n[image1]: <data:image/png;base64,iVBORw0KGgo=>");
    const output = (await callTool("get_document", { format: "markdown" })).structuredContent as any;
    expect(output.content).toBe("![][image1]\n\n[image1]: <embedded-image-omitted>");
    expect(output.message).toMatch(/1 embedded image.*get_document_images/);
    expect(output.truncated).toBeUndefined();
  });

  it("refuses markdown together with a tab or comments, naming the conflict", async () => {
    stubDocument();
    const tab = await callTool("get_document", { format: "markdown", tab_id: "t.0" });
    expect(tab.isError).toBe(true);
    expect((tab.content as any)[0].text).toMatch(/tab_id/);
    const comments = await callTool("get_document", { format: "markdown", include_comments: true });
    expect(comments.isError).toBe(true);
    expect((comments.content as any)[0].text).toMatch(/include_comments/);
  });
});

describe("append_text", () => {
  function stubAppend() {
    return stubFetch([
      ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
      [":batchUpdate", () => jsonResponse({})],
      ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(DOC)],
    ]);
  }

  function batchRequests(calls: Array<{ url: string; init?: RequestInit }>) {
    const updates = calls.filter(({ url }) => url.endsWith(":batchUpdate"));
    expect(updates).toHaveLength(1);
    return JSON.parse(String(updates[0].init?.body)).requests;
  }

  it("starts a new paragraph and resets the style it inherited from a heading", async () => {
    const calls = stubAppend();
    const result = await callTool("append_text", { text: "Body" });
    expect(result.isError).toBeFalsy();
    const [insert, chars, style, bullets] = batchRequests(calls);
    expect(insert).toEqual({ insertText: { location: { index: 6 }, text: "\nBody" } });
    expect(chars.updateTextStyle.range).toEqual({ startIndex: 6, endIndex: 11 });
    expect(style.updateParagraphStyle.range).toEqual({ startIndex: 7, endIndex: 11 });
    expect(bullets.deleteParagraphBullets.range).toEqual({ startIndex: 7, endIndex: 11 });
  });

  it("keeps the heading when the text brings its own leading newline", async () => {
    const calls = stubAppend();
    await callTool("append_text", { text: "\nBody" });
    const [insert, , style, bullets] = batchRequests(calls);
    expect(insert).toEqual({ insertText: { location: { index: 6 }, text: "\nBody" } });
    expect(style.updateParagraphStyle.range).toEqual({ startIndex: 7, endIndex: 11 });
    expect(bullets.deleteParagraphBullets.range).toEqual({ startIndex: 7, endIndex: 11 });
  });
});
