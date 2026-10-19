import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBodies, batchBody, callTool, errorText, paragraph, stubDoc } from "../lib/mcpHarness.js";
import { jsonResponse, stubFetch } from "../lib/labelStubs.js";
import { DOC_METADATA } from "../lib/mcpHarness.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1)] },
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("insert_break", () => {
  it("inserts a page break at an index", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("insert_break", { index: 3, type: "page" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ insertPageBreak: { location: { index: 3 } } }]);
  });

  it("appends a section break at the end of the tab when no index is given", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("insert_break", { type: "section_next_page", tab_id: undefined });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([
      { insertSectionBreak: { sectionType: "NEXT_PAGE", endOfSegmentLocation: {} } },
    ]);
  });

  it("continuous sections and suggest mode", async () => {
    const calls = stubDoc(DOC);
    await callTool("insert_break", { index: 1, type: "section_continuous", suggest: true });
    const body = batchBody(calls);
    expect(body.requests[0].insertSectionBreak.sectionType).toBe("CONTINUOUS");
    expect(body.writeControl).toEqual({ writeMode: "SUGGEST" });
  });

  it("refuses index 0", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("insert_break", { index: 0, type: "page" }))).toMatch(/index must be >= 1/);
  });
});

describe("set_header_footer", () => {
  const WITH_HEADER = {
    ...DOC,
    documentStyle: { defaultHeaderId: "kix.h1" },
    headers: { "kix.h1": { headerId: "kix.h1", content: [paragraph("Old header\n", 0)] } },
  };

  it("creates the header when the document has none, then writes the text into it", async () => {
    const calls = stubFetch([
      ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
      [":batchUpdate", (url) => jsonResponse(url ? { replies: [{ createHeader: { headerId: "kix.new" } }] } : {})],
      ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(DOC)],
    ]);
    const result = await callTool("set_header_footer", { kind: "header", text: "Confidential" });
    expect(result.isError).toBeFalsy();
    const [create, write] = batchBodies(calls);
    expect(create.requests).toEqual([{ createHeader: { type: "DEFAULT" } }]);
    expect(write.requests).toEqual([{ insertText: { location: { index: 0, segmentId: "kix.new" }, text: "Confidential" } }]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", kind: "header", segment_id: "kix.new", created: true });
  });

  it("replaces the existing text when the header already exists", async () => {
    const calls = stubDoc(WITH_HEADER);
    const result = await callTool("set_header_footer", { kind: "header", text: "New" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([
      { deleteContentRange: { range: { startIndex: 0, endIndex: 10, segmentId: "kix.h1" } } },
      { insertText: { location: { index: 0, segmentId: "kix.h1" }, text: "New" } },
    ]);
    expect(result.structuredContent).toMatchObject({ segment_id: "kix.h1", created: false });
  });

  it("clears the header with empty text and removes it with remove=true", async () => {
    const cleared = stubDoc(WITH_HEADER);
    await callTool("set_header_footer", { kind: "header", text: "" });
    expect(batchBody(cleared).requests).toEqual([
      { deleteContentRange: { range: { startIndex: 0, endIndex: 10, segmentId: "kix.h1" } } },
    ]);
    const removed = stubDoc(WITH_HEADER);
    const result = await callTool("set_header_footer", { kind: "header", remove: true });
    expect(batchBody(removed).requests).toEqual([{ deleteHeader: { headerId: "kix.h1" } }]);
    expect(result.structuredContent).toMatchObject({ removed: true });
  });

  it("says when there is nothing to remove, and needs text or remove", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("set_header_footer", { kind: "footer", remove: true }))).toMatch(/no footer/);
    stubDoc(DOC);
    expect(errorText(await callTool("set_header_footer", { kind: "footer" }))).toMatch(/text or remove/);
  });

  it("targets the tab's footer on a tabbed document", async () => {
    const tabbed = {
      revisionId: "rev-1",
      tabs: [{
        tabProperties: { tabId: "t.0", title: "First", index: 0 },
        documentTab: {
          body: { content: [paragraph("x\n", 1)] },
          documentStyle: { defaultFooterId: "kix.f1" },
          footers: { "kix.f1": { footerId: "kix.f1", content: [paragraph("\n", 0)] } },
        },
      }],
    };
    const calls = stubDoc(tabbed);
    const result = await callTool("set_header_footer", { kind: "footer", text: "Page", tab_id: "t.0" });
    expect(result.isError).toBeFalsy();
    // An empty footer holds only its newline, so nothing is deleted first
    expect(batchBody(calls).requests).toEqual([
      { insertText: { location: { index: 0, tabId: "t.0", segmentId: "kix.f1" }, text: "Page" } },
    ]);
  });
});

describe("insert_footnote", () => {
  it("creates the footnote at the index, then writes its text into the new segment", async () => {
    const calls = stubFetch([
      ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
      [":batchUpdate", () => jsonResponse({ replies: [{ createFootnote: { footnoteId: "kix.fn" } }] })],
      ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(DOC)],
    ]);
    const result = await callTool("insert_footnote", { index: 6, text: "See appendix." });
    expect(result.isError).toBeFalsy();
    const [create, write] = batchBodies(calls);
    expect(create.requests).toEqual([{ createFootnote: { location: { index: 6 } } }]);
    expect(write.requests).toEqual([{ insertText: { location: { index: 0, segmentId: "kix.fn" }, text: "See appendix." } }]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", footnote_id: "kix.fn" });
  });

  it("reports a created footnote whose text could not be written, so it is not created twice", async () => {
    let batches = 0;
    stubFetch([
      ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
      [":batchUpdate", () => (++batches === 1
        ? jsonResponse({ replies: [{ createFootnote: { footnoteId: "kix.fn" } }] })
        : jsonResponse({ error: { code: 400, status: "INVALID_ARGUMENT", message: "nope" } }, 400))],
      ["docs.googleapis.com/v1/documents/doc1", () => jsonResponse(DOC)],
    ]);
    const result = await callTool("insert_footnote", { index: 6, text: "x" });
    expect(result.isError).toBe(true);
    const payload = JSON.parse((result.content as any)[0].text);
    expect(payload).toMatchObject({ partialSuccess: true, footnote_id: "kix.fn", status: 400 });
    expect(payload.error).toMatch(/footnote was created/);
  });
});

describe("get_document headers and footers", () => {
  it("reports header and footer text and ids with the structure", async () => {
    stubDoc({
      ...DOC,
      documentStyle: { defaultHeaderId: "kix.h1", firstPageHeaderId: "kix.h2", defaultFooterId: "kix.f1" },
      headers: {
        "kix.h1": { headerId: "kix.h1", content: [paragraph("Running head\n", 0)] },
        "kix.h2": { headerId: "kix.h2", content: [paragraph("Cover\n", 0)] },
      },
      footers: { "kix.f1": { footerId: "kix.f1", content: [paragraph("p. 1\n", 0)] } },
    });
    const result = await callTool("get_document", { include_structure: true });
    expect(result.isError).toBeFalsy();
    const output = result.structuredContent as any;
    expect(output.headers).toEqual([
      { id: "kix.h1", type: "DEFAULT", text: "Running head\n" },
      { id: "kix.h2", type: "FIRST_PAGE", text: "Cover\n" },
    ]);
    expect(output.footers).toEqual([{ id: "kix.f1", type: "DEFAULT", text: "p. 1\n" }]);
  });

  it("omits them when the document has none", async () => {
    stubDoc(DOC);
    const output = (await callTool("get_document", { include_structure: true })).structuredContent as any;
    expect(output).not.toHaveProperty("headers");
    expect(output).not.toHaveProperty("footers");
  });
});
