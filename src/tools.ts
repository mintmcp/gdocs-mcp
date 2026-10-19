/**
 * Google Docs MCP Tools, assembled from one module per domain under src/tools/.
 * The order here is the order tools/list reports them in.
 */

import { readTools } from "./tools/read.js";
import { documentTools } from "./tools/documents.js";
import { textTools } from "./tools/text.js";
import { styleTools } from "./tools/styles.js";
import { tableTools } from "./tools/tables.js";

export class GoogleDocsTools {
  static getTools() {
    return {
      search_documents: readTools.search_documents,
      get_document: readTools.get_document,
      create_document: documentTools.create_document,
      append_text: textTools.append_text,
      replace_text: textTools.replace_text,
      delete_content: textTools.delete_content,
      insert_text: textTools.insert_text,
      update_text_style: styleTools.update_text_style,
      update_paragraph_style: styleTools.update_paragraph_style,
      append_table: tableTools.append_table,
      insert_table: tableTools.insert_table,
      update_table_style: tableTools.update_table_style,
      convert_to_google_doc: documentTools.convert_to_google_doc,
      get_document_images: readTools.get_document_images,
    };
  }
}
