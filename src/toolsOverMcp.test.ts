import { afterEach, describe, expect, it, vi } from "vitest";
import { buildDocx } from "./lib/__fixtures__/index.js";
import { jsonResponse, stubFetch } from "./lib/labelStubs.js";
import { batchBody, callTool, DOC_METADATA, stubDoc } from "./lib/mcpHarness.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, { startIndex: 1, endIndex: 7, paragraph: { elements: [{ startIndex: 1, endIndex: 7, textRun: { content: "Title\n" } }], paragraphStyle: { namedStyleType: "HEADING_1" } } }] },
};

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
  it("starts a new paragraph and resets the style it inherited from a heading", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("append_text", { text: "Body" });
    expect(result.isError).toBeFalsy();
    const [insert, chars, bullets, style] = batchBody(calls).requests;
    expect(insert).toEqual({ insertText: { location: { index: 6 }, text: "\nBody" } });
    expect(chars.updateTextStyle.range).toEqual({ startIndex: 7, endIndex: 11 });
    expect(bullets.deleteParagraphBullets.range).toEqual({ startIndex: 7, endIndex: 11 });
    expect(style.updateParagraphStyle).toEqual({
      range: { startIndex: 7, endIndex: 11 },
      paragraphStyle: { namedStyleType: "NORMAL_TEXT" },
      fields: "namedStyleType,alignment,lineSpacing,spaceAbove,spaceBelow,indentStart,indentEnd,indentFirstLine,keepWithNext,pageBreakBefore",
    });
  });

  it("keeps the heading when the text brings its own leading newline", async () => {
    const calls = stubDoc(DOC);
    await callTool("append_text", { text: "\nBody" });
    const [insert, chars, bullets, style] = batchBody(calls).requests;
    expect(insert).toEqual({ insertText: { location: { index: 6 }, text: "\nBody" } });
    for (const request of [chars.updateTextStyle, bullets.deleteParagraphBullets, style.updateParagraphStyle]) {
      expect(request.range).toEqual({ startIndex: 7, endIndex: 11 });
    }
  });
});

describe("get_document on a Word upload", () => {
  it("reads the text and notes what does not apply, instead of refusing options meant for Google Docs", async () => {
    const docx = buildDocx(["Hello Word"]);
    stubFetch([
      ["?alt=media", () => new Response(docx, { status: 200 })],
      ["/drive/v3/files/doc1?", () => jsonResponse({ name: "Report.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", size: String(docx.byteLength) })],
    ]);
    const result = await callTool("get_document", { format: "markdown", tab_id: "t.0", suggestions_view_mode: "SUGGESTIONS_INLINE" });
    expect(result.isError).toBeFalsy();
    const output = result.structuredContent as any;
    expect(output.content).toContain("Hello Word");
    expect(output.message).toMatch(/Word upload.*suggestions and markdown are unavailable/);
  });
});

