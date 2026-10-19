import { afterEach, describe, expect, it, vi } from "vitest";
import { batchBody, callTool, errorText, paragraph, stubDoc } from "../lib/mcpHarness.js";

const DOC = {
  revisionId: "rev-1",
  body: { content: [{ startIndex: 0, endIndex: 1, sectionBreak: {} }, paragraph("Title\n", 1)] },
};

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
