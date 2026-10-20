import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, paragraph, stubDoc, TITLE_DOC } from "../lib/mcpHarness.js";

const DOC = TITLE_DOC;

const TABBED = {
  revisionId: "rev-1",
  tabs: [
    { tabProperties: { tabId: "t.0", title: "First", index: 0 }, documentTab: DOC },
    { tabProperties: { tabId: "t.1", title: "Second", index: 1 }, documentTab: DOC },
  ],
};

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("add_comment", () => {
  const THREAD_REPLY = { replies: [{ insertComment: { commentThread: { commentId: "AAAC1", plainTextQuote: "Titl", status: "OPEN" } } }] };

  it("anchors a new comment to a range and returns the thread id", async () => {
    const calls = stubDoc(DOC, THREAD_REPLY);
    const result = await callTool("add_comment", { content: "Please check", startIndex: 1, endIndex: 5, tab_id: undefined });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ insertComment: { content: "Please check", range: { startIndex: 1, endIndex: 5 } } }]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", thread_id: "AAAC1", quoted_text: "Titl" });
  });

  it("assigns the comment and carries the tab", async () => {
    const calls = stubDoc(TABBED, THREAD_REPLY);
    await callTool("add_comment", { content: "Do this", startIndex: 1, endIndex: 5, tab_id: "t.1", assignee_email: "someone@example.com" });
    expect(batchBody(calls).requests[0].insertComment).toEqual({
      content: "Do this",
      assigneeEmailAddress: "someone@example.com",
      range: { startIndex: 1, endIndex: 5, tabId: "t.1" },
    });
  });

  it("replies to an existing thread when thread_id is given", async () => {
    const calls = stubDoc(DOC, { replies: [{ addCommentReply: { post: { postId: "AAAC1.R1" } } }] });
    const result = await callTool("add_comment", { content: "Done", thread_id: "AAAC1" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ addCommentReply: { commentId: "AAAC1", post: { content: "Done" } } }]);
    expect(result.structuredContent).toMatchObject({ thread_id: "AAAC1", reply: true });
  });

  it("needs a range for a new comment and refuses a range with a reply", async () => {
    stubDoc(DOC);
    expect(errorText(await callTool("add_comment", { content: "x" }))).toMatch(/startIndex and endIndex/);
    stubDoc(DOC);
    expect(errorText(await callTool("add_comment", { content: "x", thread_id: "AAAC1", startIndex: 1, endIndex: 2 }))).toMatch(/thread_id/);
  });

  it("reports an unsaved comment as a plain failure, since nothing else changed", async () => {
    stubDoc(DOC, { replies: [{}], commentUpdateState: "ALL_FAILED_UNKNOWN_REASON" });
    const result = await callTool("add_comment", { content: "x", startIndex: 1, endIndex: 5 });
    const payload = JSON.parse((result.content as any)[0].text);
    expect(result.isError).toBe(true);
    expect(payload.error).toMatch(/saved none of the comments/);
    expect(payload).not.toHaveProperty("partialSuccess");
  });

  it("measures the 2048 limit in UTF-8 bytes, as Google does", async () => {
    const calls = stubDoc(DOC);
    expect(errorText(await callTool("add_comment", { content: "😀".repeat(600), thread_id: "AAAC1" }))).toMatch(/2400 bytes/);
    expect(calls.filter(({ url }) => url.endsWith(":batchUpdate"))).toHaveLength(0);
    stubDoc(DOC, { replies: [{}] });
    expect((await callTool("add_comment", { content: "é".repeat(1024), thread_id: "AAAC1" })).isError).toBeFalsy();
  });
});

describe("manage_tab", () => {
  it("adds a tab with a title and reports its id", async () => {
    const calls = stubDoc(TABBED, { replies: [{ addDocumentTab: { tabProperties: { tabId: "t.new", title: "Notes", index: 2 } } }] });
    const result = await callTool("manage_tab", { action: "add", title: "Notes" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ addDocumentTab: { tabProperties: { title: "Notes" } } }]);
    expect(result.structuredContent).toMatchObject({ id: "doc1", action: "add", tab: { tabId: "t.new", title: "Notes", index: 2 } });
  });

  it("renames a tab", async () => {
    const calls = stubDoc(TABBED);
    const result = await callTool("manage_tab", { action: "rename", tab_id: "t.1", title: "Renamed" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ updateDocumentTabProperties: { tabProperties: { tabId: "t.1", title: "Renamed" }, fields: "title" } }]);
    expect((result.structuredContent as any).tab).toEqual({ tabId: "t.1", title: "Renamed", index: 1, nestingLevel: 0 });
  });

  it("deletes a tab, but never the last one", async () => {
    const calls = stubDoc(TABBED);
    const result = await callTool("manage_tab", { action: "delete", tab_id: "t.1" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ deleteTab: { tabId: "t.1" } }]);
    stubDoc({ revisionId: "rev-1", tabs: [TABBED.tabs[0]] });
    expect(errorText(await callTool("manage_tab", { action: "delete", tab_id: "t.0" }))).toMatch(/only top-level tab/);
  });

  it("explains an unknown tab and missing arguments", async () => {
    stubDoc(TABBED);
    expect(errorText(await callTool("manage_tab", { action: "rename", tab_id: "t.9", title: "x" }))).toMatch(/No tab with id 't.9'/);
    stubDoc(TABBED);
    expect(errorText(await callTool("manage_tab", { action: "rename", tab_id: "t.1" }))).toMatch(/title/);
    stubDoc(TABBED);
    expect(errorText(await callTool("manage_tab", { action: "delete" }))).toMatch(/tab_id/);
  });
});

describe("manage_tab nested tabs", () => {
  it("refuses deleting the only top-level tab even when it has children, but allows deleting a child", async () => {
    const nested = {
      revisionId: "rev-1",
      tabs: [{
        tabProperties: { tabId: "t.0", title: "Root", index: 0 },
        documentTab: DOC,
        childTabs: [{ tabProperties: { tabId: "t.child", title: "Child", index: 0, nestingLevel: 1 }, documentTab: DOC }],
      }],
    };
    stubDoc(nested);
    expect(errorText(await callTool("manage_tab", { action: "delete", tab_id: "t.0" }))).toMatch(/only top-level tab/);
    const calls = stubDoc(nested);
    const result = await callTool("manage_tab", { action: "delete", tab_id: "t.child" });
    expect(result.isError).toBeFalsy();
    expect(batchBody(calls).requests).toEqual([{ deleteTab: { tabId: "t.child" } }]);
  });
});

describe("argument checks before any network call", () => {
  it("refuses a reply that names a tab, and a rename without a title", async () => {
    const calls = stubDoc(TABBED);
    expect(errorText(await callTool("add_comment", { content: "x", thread_id: "AAAC1", tab_id: "t.1" }))).toMatch(/tab_id/);
    expect(errorText(await callTool("manage_tab", { action: "rename", tab_id: "t.1" }))).toMatch(/title/);
    expect(calls.filter(({ url }) => url.includes("docs.googleapis.com"))).toHaveLength(0);
  });
});

describe("get_document on one tab", () => {
  it("returns a table's full text in content, past the cap structure puts on a cell", async () => {
    const long = "x".repeat(500);
    stubDoc({
      revisionId: "rev-1",
      tabs: [{
        tabProperties: { tabId: "t.0", title: "Only", index: 0 },
        documentTab: { body: { content: [
          { startIndex: 1, endIndex: 600, table: { rows: 1, columns: 1, tableRows: [{ tableCells: [{ content: [paragraph(`${long}\n`, 2)] }] }] } },
        ] } },
      }],
    });
    const output = (await callTool("get_document", { tab_id: "t.0", include_structure: true })).structuredContent as any;
    expect(output.content).toBe(`${long}\n`);
    expect(output.structure[0].cells[0][0]).toMatch(/\.\.\.$/);
  });
});
