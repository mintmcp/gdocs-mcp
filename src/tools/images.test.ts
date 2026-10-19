import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, paragraph, stubDoc } from "../lib/mcpHarness.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1)] },
  inlineObjects: { "kix.img1": { inlineObjectProperties: { embeddedObject: { imageProperties: { contentUri: "https://lh3.googleusercontent.com/x" } } } } },
};

const REPLY = { replies: [{ insertInlineImage: { objectId: "kix.new" } }] };

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("insert_image", () => {
  it("inserts a public image at an index with a size and returns its object id", async () => {
    const calls = stubDoc(DOC, REPLY);
    const result = await callTool("insert_image", { image_url: "https://example.com/a.png", index: 3, width_pt: 200 });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{
      insertInlineImage: {
        uri: "https://example.com/a.png",
        location: { index: 3 },
        objectSize: { width: { magnitude: 200, unit: "PT" } },
      },
    }]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", object_id: "kix.new" });
  });

  it("appends at the end of the tab without a size when index is omitted", async () => {
    const calls = stubDoc(DOC, REPLY);
    await callTool("insert_image", { image_url: "https://example.com/a.png", tab_id: undefined });
    expect(batchBody(calls).requests[0].insertInlineImage).toEqual({
      uri: "https://example.com/a.png",
      endOfSegmentLocation: {},
    });
  });

  it("turns a Drive share link into a fetchable download URL", async () => {
    const calls = stubDoc(DOC, REPLY);
    await callTool("insert_image", { image_url: "https://drive.google.com/file/d/1AbC_dEf/view?usp=sharing", index: 1 });
    expect(batchBody(calls).requests[0].insertInlineImage.uri).toBe("https://drive.google.com/uc?export=download&id=1AbC_dEf");
  });

  it("rejects non-https, data: and over-long URLs before calling Google", async () => {
    stubDoc(DOC, REPLY);
    expect(errorText(await callTool("insert_image", { image_url: "http://example.com/a.png", index: 1 }))).toMatch(/https/);
    stubDoc(DOC, REPLY);
    expect(errorText(await callTool("insert_image", { image_url: "data:image/png;base64,AAAA", index: 1 }))).toMatch(/https/);
    stubDoc(DOC, REPLY);
    expect(errorText(await callTool("insert_image", { image_url: `https://example.com/${"a".repeat(2100)}`, index: 1 }))).toMatch(/2 ?kB|2048/);
  });

  it("explains a Google refusal to fetch the image", async () => {
    stubDoc(DOC, { error: { code: 400, status: "INVALID_ARGUMENT", message: "Invalid requests[0].insertInlineImage: The provided image is in an unsupported format." } });
    vi.stubGlobal("fetch", vi.fn(async (input: any) => {
      const url = String(input);
      if (url.endsWith(":batchUpdate")) return new Response(JSON.stringify({ error: { code: 400, status: "INVALID_ARGUMENT", message: "The provided image is in an unsupported format." } }), { status: 400, headers: { "Content-Type": "application/json" } });
      if (url.includes("/drive/v3/files/doc1?")) return new Response(JSON.stringify({ name: "Doc", mimeType: "application/vnd.google-apps.document" }), { status: 200, headers: { "Content-Type": "application/json" } });
      return new Response(JSON.stringify(DOC), { status: 200, headers: { "Content-Type": "application/json" } });
    }));
    const text = errorText(await callTool("insert_image", { image_url: "https://example.com/a.png", index: 1 }));
    expect(text).toMatch(/unsupported format/);
    expect(text).toMatch(/PNG, JPEG or GIF/);
  });
});

describe("replace_image", () => {
  it("replaces an existing inline image by object id, keeping its size", async () => {
    const calls = stubDoc(DOC);
    const result = await callTool("replace_image", { object_id: "kix.img1", image_url: "https://example.com/b.jpg" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{
      replaceImage: { imageObjectId: "kix.img1", uri: "https://example.com/b.jpg", imageReplaceMethod: "CENTER_CROP" },
    }]);
  });

  it("refuses an object id the document does not have", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("replace_image", { object_id: "kix.nope", image_url: "https://example.com/b.jpg" }))).toMatch(/kix.img1/);
  });
});

describe("replace_image tab resolution", () => {
  it("sends the tab that owns the image when no tab_id is given", async () => {
    const tabbed = {
      revisionId: "rev-1",
      tabs: [
        { tabProperties: { tabId: "t.0", index: 0 }, documentTab: { body: { content: [] }, inlineObjects: {} } },
        { tabProperties: { tabId: "t.1", index: 1 }, documentTab: { body: { content: [] }, inlineObjects: { "kix.img2": {} } } },
      ],
    };
    const calls = stubDoc(tabbed);
    const result = await callTool("replace_image", { object_id: "kix.img2", image_url: "https://example.com/b.jpg" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests[0].replaceImage.tabId).toBe("t.1");
  });
});
