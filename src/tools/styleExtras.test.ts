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

describe("update_text_style extras", () => {
  it("sets font family, background colour, baseline offset and small caps", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("update_text_style", {
      startIndex: 1, endIndex: 6, font_family: "Roboto", background_color: "#FFFF00", baseline_offset: "SUPERSCRIPT", small_caps: true,
    });
    expect(result.isError).toBeFalsy();
    const { textStyle, fields } = batchBody(calls).requests[0].updateTextStyle;
    expect(textStyle).toEqual({
      weightedFontFamily: { fontFamily: "Roboto" },
      backgroundColor: { color: { rgbColor: { red: 1, green: 1, blue: 0 } } },
      baselineOffset: "SUPERSCRIPT",
      smallCaps: true,
    });
    expect(fields).toBe("weightedFontFamily,backgroundColor,baselineOffset,smallCaps");
  });

  it("clears a background colour with an empty string", async () => {
    const calls = stubDoc(DOC);
    await callTool("update_text_style", { startIndex: 1, endIndex: 6, background_color: "" });
    const { textStyle, fields } = batchBody(calls).requests[0].updateTextStyle;
    expect(textStyle).toEqual({ backgroundColor: {} });
    expect(fields).toBe("backgroundColor");
  });
});

describe("update_paragraph_style extras", () => {
  it("sets spacing, indents, line spacing and keep-with-next", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("update_paragraph_style", {
      startIndex: 1, endIndex: 6, line_spacing: 150, space_above_pt: 12, space_below_pt: 6, indent_start_pt: 36, indent_end_pt: 0, indent_first_line_pt: 18, keep_with_next: true, page_break_before: true,
    });
    expect(result.isError).toBeFalsy();
    const { paragraphStyle, fields } = batchBody(calls).requests[0].updateParagraphStyle;
    expect(paragraphStyle).toEqual({
      lineSpacing: 150,
      spaceAbove: { magnitude: 12, unit: "PT" },
      spaceBelow: { magnitude: 6, unit: "PT" },
      indentStart: { magnitude: 36, unit: "PT" },
      indentEnd: { magnitude: 0, unit: "PT" },
      indentFirstLine: { magnitude: 18, unit: "PT" },
      keepWithNext: true,
      pageBreakBefore: true,
    });
    expect(fields).toBe("lineSpacing,spaceAbove,spaceBelow,indentStart,indentEnd,indentFirstLine,keepWithNext,pageBreakBefore");
  });

  it("applies TITLE and SUBTITLE through named_style and refuses it together with heading_level", async () => {
    const calls = stubDoc(DOC);
    await callTool("update_paragraph_style", { startIndex: 1, endIndex: 6, named_style: "TITLE" });
    expect(batchBody(calls).requests[0].updateParagraphStyle.paragraphStyle).toEqual({ namedStyleType: "TITLE" });
    stubDoc(DOC);
    expect(errorText(await callTool("update_paragraph_style", { startIndex: 1, endIndex: 6, named_style: "TITLE", heading_level: 1 }))).toMatch(/one of heading_level or named_style/);
  });

  it("takes a Google bullet preset for list_style", async () => {
    const calls = stubDoc(DOC);
    await callTool("update_paragraph_style", { startIndex: 1, endIndex: 6, list_style: "bullet", bullet_preset: "BULLET_ARROW_DIAMOND_DISC" });
    expect(batchBody(calls).requests[0].createParagraphBullets.bulletPreset).toBe("BULLET_ARROW_DIAMOND_DISC");
    stubDoc(DOC);
    expect(errorText(await callTool("update_paragraph_style", { startIndex: 1, endIndex: 6, bullet_preset: "BULLET_ARROW_DIAMOND_DISC" }))).toMatch(/bullet_preset needs list_style/);
  });
});

describe("replace_text use_regex", () => {
  it("passes searchByRegex and keeps matchCase", async () => {
    const calls = stubDoc(DOC, { replies: [{ replaceAllText: { occurrencesChanged: 2 } }] });
    const result = await callTool("replace_text", { old_text: "v(\\d+)", new_text: "version", use_regex: true, match_case: false });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests[0].replaceAllText.containsText).toEqual({ text: "v(\\d+)", matchCase: false, searchByRegex: true });
    expect((result.structuredContent as any).occurrencesChanged).toBe(2);
  });

  it("does not send searchByRegex when off", async () => {
    const calls = stubDoc(DOC, { replies: [] });
    await callTool("replace_text", { old_text: "a", new_text: "b" });
    expect(batchBody(calls).requests[0].replaceAllText.containsText).toEqual({ text: "a", matchCase: true });
  });
});
