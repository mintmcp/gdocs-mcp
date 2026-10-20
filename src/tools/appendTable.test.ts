import { afterEach, describe, expect, it, vi } from "vitest";
import { jsonResponse, stubFetch } from "../lib/labelStubs.js";
import { DOC_METADATA, batchBodies, callTool, paragraph } from "../lib/mcpHarness.js";

afterEach(() => vi.unstubAllGlobals());

const before = { revisionId: "r1", body: { content: [paragraph("Intro\n", 1)] } };
const after = {
  revisionId: "r2",
  body: { content: [paragraph("Intro\n", 1), { startIndex: 7, endIndex: 20, table: { rows: 1, columns: 2 } }, paragraph("\n", 20)] },
};

function stubAppend(afterRead: () => Response) {
  let reads = 0;
  return stubFetch([
    ["/drive/v3/files/doc1?", () => jsonResponse(DOC_METADATA)],
    [":batchUpdate", () => jsonResponse({})],
    ["docs.googleapis.com/v1/documents/doc1", () => (reads++ === 0 ? jsonResponse(before) : afterRead())],
  ]);
}

describe("append_table", () => {
  it("inserts at the end of the body under the revision gate, then clears the formatting of the table it finds", async () => {
    const calls = stubAppend(() => jsonResponse(after));
    const result = await callTool("append_table", { rows: [["a", "b"]], required_revision_id: "r1" });

    expect(result.isError).toBeFalsy();
    const [insert, clear] = batchBodies(calls);
    expect(insert.requests[0]).toEqual({ insertTable: { location: { index: 6 }, rows: 1, columns: 2 } });
    expect(insert.writeControl).toEqual({ requiredRevisionId: "r1" });
    expect(clear.requests[0].updateTextStyle.range).toEqual({ startIndex: 8, endIndex: 20 });
    expect(clear.writeControl).toEqual({ targetRevisionId: "r2" });
    expect(result.structuredContent).toMatchObject({ table_start_index: 7, table_start_index_verified: true });
  });

  it("leaves a table of another shape at the insert point alone", async () => {
    const theirs = { ...after, body: { content: [after.body.content[0], { ...after.body.content[1], table: { rows: 3, columns: 2 } }, after.body.content[2]] } };
    const calls = stubAppend(() => jsonResponse(theirs));
    const result = await callTool("append_table", { rows: [["a", "b"]] });

    expect(result.isError).toBeFalsy();
    expect(batchBodies(calls)).toHaveLength(1);
    expect(result.structuredContent).toMatchObject({ table_start_index_verified: false });
    expect((result.structuredContent as any).warning).toMatch(/Do not insert it again/);
  });

  it("warns without repeating Google's error text when the table cannot be re-read", async () => {
    stubAppend(() => jsonResponse({ error: { code: 500, message: "secret upstream detail", status: "INTERNAL" } }, 500));
    const result = await callTool("append_table", { rows: [["a", "b"]] });

    expect(result.isError).toBeFalsy();
    const { warning, table_start_index_verified } = result.structuredContent as any;
    expect(warning).toMatch(/Do not insert it again/);
    expect(warning).not.toMatch(/secret upstream detail/);
    expect(table_start_index_verified).toBe(false);
  });
});

describe("insert_table", () => {
  it("inserts against the revision its nesting check read, or the one the caller requires", async () => {
    const targeted = stubAppend(() => jsonResponse(after));
    await callTool("insert_table", { rows: [["a", "b"]], index: 6 });
    expect(batchBodies(targeted)[0].writeControl).toEqual({ targetRevisionId: "r1" });

    const required = stubAppend(() => jsonResponse(after));
    await callTool("insert_table", { rows: [["a", "b"]], index: 6, required_revision_id: "r0" });
    expect(batchBodies(required)[0].writeControl).toEqual({ requiredRevisionId: "r0" });
  });
});
