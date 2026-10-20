import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, stubDoc, TITLE_DOC } from "../lib/mcpHarness.js";

const DOC = TITLE_DOC;

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
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
    expect(fields).toBe("marginTop,marginBottom,marginLeft,marginRight,useFirstPageHeaderFooter,pageNumberStart,pageSize");    expect(batchBody(calls)).not.toHaveProperty("writeControl");
  });

  it("flips the current page size when only the orientation changes", async () => {
    const calls = stubDoc({ ...DOC, documentStyle: { pageSize: { width: { magnitude: 612, unit: "PT" }, height: { magnitude: 792, unit: "PT" } } } });
    await callTool("update_document_style", { orientation: "landscape" });
    expect(batchBody(calls).requests[0].updateDocumentStyle.documentStyle).toEqual({
      pageSize: { width: { magnitude: 792, unit: "PT" }, height: { magnitude: 612, unit: "PT" } },
    });
    expect(batchBody(calls).writeControl).toEqual({ requiredRevisionId: DOC.revisionId });
  });

  it("targets a tab and needs at least one property", async () => {
    const tabbed = { revisionId: "rev-1", tabs: [{ tabProperties: { tabId: "t.0", index: 0 }, documentTab: DOC }] };
    const calls = stubDoc(tabbed);
    await callTool("update_document_style", { margin_top_pt: 10, tab_id: "t.0" });
    expect(batchBody(calls).requests[0].updateDocumentStyle.tabId).toBe("t.0");
    expect(calls.some(({ url }) => url.includes("docs.googleapis.com/v1/documents/doc1?"))).toBe(false);
    stubDoc(DOC);
    expect(errorText(await callTool("update_document_style", {}))).toMatch(/At least one/);
  });
});
