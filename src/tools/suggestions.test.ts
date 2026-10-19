import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, paragraph, stubDoc } from "../lib/mcpHarness.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1)] },
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("suggest mode on write tools", () => {
  it("sends writeMode SUGGEST with the revision pin and says so", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("insert_text", { text: "x", index: 1, suggest: true, required_revision_id: "rev-1" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).writeControl).toEqual({ requiredRevisionId: "rev-1", writeMode: "SUGGEST" });
    expect((result.structuredContent as any).message).toMatch(/as a suggestion/);
  });

  it("sends no writeControl at all when neither is asked for", async () => {
    const calls = stubDoc(DOC);
    await callTool("insert_text", { text: "x", index: 1 });
    expect(batchBody(calls)).not.toHaveProperty("writeControl");
  });

  it.each([
    ["append_text", { text: "x" }],
    ["replace_text", { old_text: "a", new_text: "b" }],
    ["delete_content", { startIndex: 1, endIndex: 3 }],
    ["update_text_style", { startIndex: 1, endIndex: 3, bold: true }],
    ["update_paragraph_style", { startIndex: 1, endIndex: 3, alignment: "CENTER" }],
    ["insert_table", { rows: [["a"]], index: 1, clear_inherited_formatting: false }],
  ])("%s accepts suggest", async (tool, args) => {
    const calls = stubDoc({ ...DOC, replies: [] });
    const result = await callTool(tool, { ...args, suggest: true });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).writeControl).toEqual({ writeMode: "SUGGEST" });
  });
});

describe("segment_id plumbing", () => {
  it("inserts into a header at index 0 with the segment on the location", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("insert_text", { text: "Hdr", index: 0, segment_id: "kix.h1" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests[0].insertText.location).toEqual({ index: 0, segmentId: "kix.h1" });
  });

  it("still refuses index 0 in the body", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("insert_text", { text: "x", index: 0 }))).toMatch(/index must be >= 1/);
  });

  it("deletes and styles ranges inside a segment", async () => {
    const calls = stubDoc(DOC);
    await callTool("delete_content", { startIndex: 0, endIndex: 3, segment_id: "kix.f1", tab_id: "t.0" });
    expect(batchBody(calls).requests[0].deleteContentRange.range).toEqual({ startIndex: 0, endIndex: 3, tabId: "t.0", segmentId: "kix.f1" });

    const styled = stubDoc(DOC);
    await callTool("update_text_style", { startIndex: 0, endIndex: 3, segment_id: "kix.f1", bold: true });
    expect(batchBody(styled).requests[0].updateTextStyle.range).toEqual({ startIndex: 0, endIndex: 3, segmentId: "kix.f1" });
  });
});

describe("review_suggestion", () => {
  it("accepts several suggestions in one batch", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("review_suggestion", { suggestion_ids: ["suggest.a", "suggest.b"], action: "accept" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([
      { acceptSuggestion: { suggestionId: "suggest.a" } },
      { acceptSuggestion: { suggestionId: "suggest.b" } },
    ]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", action: "accept", count: 2 });
  });

  it("rejects and deletes with their own request kinds", async () => {
    const rejected = stubDoc(DOC);
    await callTool("review_suggestion", { suggestion_ids: ["suggest.a"], action: "reject" });
    expect(batchBody(rejected).requests).toEqual([{ rejectSuggestion: { suggestionId: "suggest.a" } }]);
    const deleted = stubDoc(DOC);
    await callTool("review_suggestion", { suggestion_ids: ["suggest.a"], action: "delete" });
    expect(batchBody(deleted).requests).toEqual([{ deleteSuggestion: { suggestionId: "suggest.a" } }]);
  });
});

describe("get_document suggestions", () => {
  const SUGGESTED_DOC = {
    revisionId: "rev-1",
    body: {
      content: [
        { startIndex: 0, endIndex: 1, sectionBreak: {} },
        {
          startIndex: 1, endIndex: 12,
          paragraph: { elements: [
            { startIndex: 1, endIndex: 5, textRun: { content: "New ", suggestedInsertionIds: ["suggest.ins"] } },
            { startIndex: 5, endIndex: 9, textRun: { content: "old ", suggestedDeletionIds: ["suggest.del"] } },
            { startIndex: 9, endIndex: 12, textRun: { content: "txt\n" } },
          ] },
        },
        { startIndex: 12, endIndex: 20, table: { rows: 1, columns: 1, suggestedInsertionIds: ["suggest.tbl"], tableRows: [] } },
      ],
    },
  };

  it("passes suggestions_view_mode to the Docs read and lists suggestion ids on elements", async () => {
    const calls = stubDoc(SUGGESTED_DOC);
    const result = await callTool("get_document", { tab_id: undefined, include_structure: true, suggestions_view_mode: "PREVIEW_WITHOUT_SUGGESTIONS" });
    expect(result.isError).toBeFalsy();
    const docsRead = calls.find(({ url }) => url.includes("docs.googleapis.com/v1/documents/doc1?"));
    expect(docsRead?.url).toContain("suggestionsViewMode=PREVIEW_WITHOUT_SUGGESTIONS");
    const structure = (result.structuredContent as any).structure;
    expect(structure[1].suggestions).toEqual([
      { id: "suggest.ins", kind: "insertion" },
      { id: "suggest.del", kind: "deletion" },
    ]);
    expect(structure[2].suggestions).toEqual([{ id: "suggest.tbl", kind: "insertion" }]);
  });

  it("omits suggestions on elements that have none", async () => {
    stubDoc(SUGGESTED_DOC);
    const result = await callTool("get_document", { include_structure: true });
    expect((result.structuredContent as any).structure[0]).not.toHaveProperty("suggestions");
  });
});

describe("get_document suggestion and footnote details", () => {
  it("lists formatting suggestions and footnote references", async () => {
    stubDoc({
      revisionId: "rev-1",
      body: { content: [
        { startIndex: 1, endIndex: 9, paragraph: {
          suggestedParagraphStyleChanges: { "suggest.para": {} },
          elements: [
            { startIndex: 1, endIndex: 5, textRun: { content: "Bold", suggestedTextStyleChanges: { "suggest.fmt": {} } } },
            { startIndex: 5, endIndex: 6, footnoteReference: { footnoteId: "kix.fn1" } },
            { startIndex: 6, endIndex: 9, textRun: { content: "..\n" } },
          ],
        } },
      ] },
      footnotes: { "kix.fn1": { content: [paragraph("Note text\n", 0)] } },
    });
    const output = (await callTool("get_document", { include_structure: true })).structuredContent as any;
    expect(output.structure[0].suggestions).toEqual([{ id: "suggest.para", kind: "formatting" }, { id: "suggest.fmt", kind: "formatting" }]);
    expect(output.structure[1]).toEqual({ type: "footnoteReference", startIndex: 5, endIndex: 6, footnoteId: "kix.fn1" });
    expect(output.footnotes).toEqual([{ id: "kix.fn1", text: "Note text\n" }]);
  });

  it("refuses suggestions_view_mode without a Docs read to apply it to", async () => {
    stubDoc({ revisionId: "rev-1", body: { content: [] } });
    expect(errorText(await callTool("get_document", { suggestions_view_mode: "PREVIEW_SUGGESTIONS_ACCEPTED" }))).toMatch(/include_structure=true or tab_id/);
  });
});

describe("table suggestions", () => {
  it("lists suggestions made on rows and cells, such as a suggested cell style", async () => {
    stubDoc({
      revisionId: "rev-1",
      body: { content: [{ startIndex: 1, endIndex: 9, table: { rows: 1, columns: 1, tableRows: [
        { suggestedInsertionIds: ["suggest.row"], tableCells: [{ suggestedTableCellStyleChanges: { "suggest.cell": {} }, content: [] }] },
      ] } }] },
    });
    const output = (await callTool("get_document", { include_structure: true })).structuredContent as any;
    expect(output.structure[0].suggestions).toEqual([{ id: "suggest.row", kind: "insertion" }, { id: "suggest.cell", kind: "formatting" }]);
  });
});
