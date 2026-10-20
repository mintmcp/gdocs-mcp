import { afterEach, describe, expect, it, vi } from "vitest";
import { jsonResponse, stubFetch } from "../lib/labelStubs.js";
import { callTool } from "../lib/mcpHarness.js";

afterEach(() => vi.unstubAllGlobals());

const CREATED = { id: "new1", name: "Notes", mimeType: "application/vnd.google-apps.document" };

function stubCreate() {
  return stubFetch([
    ["/upload/drive/v3/files", () => jsonResponse(CREATED)],
    ["/drive/v3/files?", () => jsonResponse(CREATED)],
  ]);
}

describe("create_document", () => {
  it("imports a plain body verbatim in the one upload that creates the file", async () => {
    const calls = stubCreate();
    const body = "    indented\n\ttab\n*not bold*";
    const result = await callTool("create_document", { title: "Notes", body, body_format: "plain" });
    expect(result.isError).toBeFalsy();
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toContain("/upload/drive/v3/files");
    const sent = String(calls[0].init?.body);
    expect(sent).toContain("Content-Type: text/plain; charset=UTF-8");
    expect(sent).toContain(body);
    expect(result.structuredContent).toMatchObject({ id: "new1", webViewLink: "https://docs.google.com/document/d/new1/edit", message: "Document created successfully" });
  });

  it("imports markdown as markdown", async () => {
    const calls = stubCreate();
    const result = await callTool("create_document", { title: "Notes", body: "# Hi" });
    expect(String(calls[0].init?.body)).toContain("Content-Type: text/markdown");
    expect((result.structuredContent as any).message).toMatch(/markdown converted/);
  });

  it("creates an empty Doc from metadata alone when there is no body", async () => {
    const calls = stubCreate();
    await callTool("create_document", { title: "Notes" });
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toContain("/drive/v3/files?");
    expect(JSON.parse(String(calls[0].init?.body))).toEqual({ name: "Notes", mimeType: "application/vnd.google-apps.document" });
  });

  it("refuses to report success when Drive stored the bytes without converting them", async () => {
    stubFetch([["/upload/drive/v3/files", () => jsonResponse({ ...CREATED, mimeType: "text/plain" })]]);
    const result = await callTool("create_document", { title: "Notes", body: "x", body_format: "plain" });
    expect(result.isError).toBe(true);
    expect(JSON.parse((result.content as any)[0].text).error).toMatch(/instead of converting it/);
  });
});
