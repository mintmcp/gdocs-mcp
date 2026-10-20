import { z } from 'zod';
import { docRange, validateIndexRange } from "../lib/docRequests.js";
import {
  docsBatchUpdate,
  documentIdField,
  editTool,
  tabIdField,
  toolResult,
} from './shared.js';

// Google counts this in UTF-8 code units, so emoji and accents take several each
const MAX_COMMENT_BYTES = 2048;

export const commentTools = {
  add_comment: editTool({
    description: 'Add a comment to a Google Doc. Without thread_id it starts a new thread anchored to the text between startIndex and endIndex (from get_document with include_structure=true), optionally assigned to someone by email. With thread_id it replies to an existing thread: the `thread_id` get_document reports with include_comments=true. Comments are plain text, at most 2048 bytes as UTF-8 (fewer characters when they include emoji or accents). Native Google Docs only.',
    outputSchema: {
      id: z.string(),
      thread_id: z.string().optional().describe('Id of the thread, which is also the disco= parameter of its deep link'),
      reply: z.boolean().describe('True when a reply was added rather than a new thread'),
      quoted_text: z.string().optional().describe('The text the new thread is anchored to, as Google saw it'),
      message: z.string(),
    },
    schema: {
      document_id: documentIdField,
      content: z.string().min(1).max(MAX_COMMENT_BYTES).describe('The comment text, at most 2048 bytes as UTF-8'),
      startIndex: z.coerce.number().int().optional().describe('New thread: start of the text to comment on'),
      endIndex: z.coerce.number().int().optional().describe('New thread: end of the text to comment on (exclusive)'),
      tab_id: tabIdField,
      assignee_email: z.string().email().optional().describe('New thread: assign the comment to this collaborator'),
      thread_id: z.string().optional().describe('Reply to this thread instead of starting one'),
    },
    run: async ({ document_id, content, startIndex, endIndex, tab_id, assignee_email, thread_id }, { accessToken }) => {
      const bytes = new TextEncoder().encode(content).length;
      if (bytes > MAX_COMMENT_BYTES) {
        throw new Error(`The comment is ${bytes} bytes as UTF-8; Google accepts at most ${MAX_COMMENT_BYTES}. Shorten it.`);
      }
      if (thread_id) {
        if (startIndex !== undefined || endIndex !== undefined || assignee_email || tab_id) {
          throw new Error('thread_id replies to an existing thread, which already has its anchor, tab and assignee: drop startIndex, endIndex, tab_id and assignee_email, or drop thread_id to start a new thread.');
        }
        await docsBatchUpdate(document_id, accessToken, [
          { addCommentReply: { commentId: thread_id, post: { content } } },
        ]);
        return toolResult({ id: document_id, thread_id, reply: true, message: `Replied to comment thread ${thread_id}` });
      }

      if (startIndex === undefined || endIndex === undefined) {
        throw new Error('A new comment needs startIndex and endIndex for the text it is anchored to; pass thread_id instead to reply to an existing thread.');
      }
      validateIndexRange(startIndex, endIndex);

      const reply = await docsBatchUpdate(document_id, accessToken, [{
        insertComment: {
          content,
          ...(assignee_email ? { assigneeEmailAddress: assignee_email } : {}),
          range: docRange(startIndex, endIndex, tab_id),
        },
      }]);
      const thread = reply?.replies?.[0]?.insertComment?.commentThread;

      return toolResult({
        id: document_id,
        ...(thread?.commentId ? { thread_id: thread.commentId } : {}),
        reply: false,
        ...(typeof thread?.plainTextQuote === 'string' ? { quoted_text: thread.plainTextQuote } : {}),
        message: `Comment added on range ${startIndex}-${endIndex}${assignee_email ? `, assigned to ${assignee_email}` : ''}`,
      });
    },
  }),
};
