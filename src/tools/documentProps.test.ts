import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, paragraph, stubDoc } from "../lib/mcpHarness.js";
import { jsonResponse, stubFetch } from "../lib/labelStubs.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1)] },
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("rename_document", () => {
  it("patches the Drive name and returns the new one", async () => {
    const calls = stubFetch([
      ["/drive/v3/files/doc1?", (url) => jsonResponse(url.includes("fields=name,mimeType")
        ? { name: "Old", mimeType: "application/vnd.google-apps.document", webViewLink: "https://docs.google.com/document/d/doc1" }
        : { id: "doc1", name: "New name", webViewLink: "https://docs.google.com/document/d/doc1" })],
    ]);
    const result = await callTool("rename_document", { name: "New name" });
    expect(result.isError).toBeFalsy();
    const patch = calls.find(({ init }) => init?.method === "PATCH");
    expect(patch?.url).toContain("/drive/v3/files/doc1?");
    expect(patch?.url).toContain("supportsAllDrives=true");
    expect(JSON.parse(String(patch?.init?.body))).toEqual({ name: "New name" });
    expect(result.structuredContent).toMatchObject({ id: "doc1", name: "New name", previous_name: "Old" });
  });

  it("refuses an empty or whitespace name and a Word upload", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("rename_document", { name: "   " }))).toMatch(/name/);
    stubFetch([["/drive/v3/files/doc1?", () => jsonResponse({ name: "x.docx", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", webViewLink: "https://drive.google.com/file/d/doc1/view" })]]);
    expect(errorText(await callTool("rename_document", { name: "y" }))).toMatch(/Word/);
  });
});

describe("update_document_style", () => {
  it("sets margins, page size, orientation, first-page header and page number start", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("update_document_style", {
      margin_top_pt: 36, margin_bottom_pt: 36, margin_left_pt: 54, margin_right_pt: 54,
      page_size: "A4", orientation: "landscape", first_page_header_footer: true, page_number_start: 3,
    });
    expect(result.isError).toBeFalsy();
    const { documentStyle, fields } = batchBody(calls).requests[0].updateDocumentStyle;
    expect(documentStyle).toEqual({
      marginTop: { magnitude: 36, unit: "PT" },
      marginBottom: { magnitude: 36, unit: "PT" },
      marginLeft: { magnitude: 54, unit: "PT" },
      marginRight: { magnitude: 54, unit: "PT" },
      pageSize: { width: { magnitude: 841.9, unit: "PT" }, height: { magnitude: 595.3, unit: "PT" } },
      useFirstPageHeaderFooter: true,
      pageNumberStart: 3,
    });
    expect(fields).toBe("marginTop,marginBottom,marginLeft,marginRight,pageSize,useFirstPageHeaderFooter,pageNumberStart");
  });

  it("flips the current page size when only the orientation changes", async () => {
    const calls = stubDoc({ ...DOC, documentStyle: { pageSize: { width: { magnitude: 612, unit: "PT" }, height: { magnitude: 792, unit: "PT" } } } });
    await callTool("update_document_style", { orientation: "landscape" });
    expect(batchBody(calls).requests[0].updateDocumentStyle.documentStyle).toEqual({
      pageSize: { width: { magnitude: 792, unit: "PT" }, height: { magnitude: 612, unit: "PT" } },
    });
  });

  it("targets a tab and needs at least one property", async () => {
    const tabbed = { revisionId: "rev-1", tabs: [{ tabProperties: { tabId: "t.0", index: 0 }, documentTab: DOC }] };
    const calls = stubDoc(tabbed);
    await callTool("update_document_style", { margin_top_pt: 10, tab_id: "t.0" });
    expect(batchBody(calls).requests[0].updateDocumentStyle.tabId).toBe("t.0");
    stubDoc(DOC);
    expect(errorText(await callTool("update_document_style", {}))).toMatch(/At least one/);
  });
});
