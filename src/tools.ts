import { readTools } from "./tools/read.js";
import { documentTools } from "./tools/documents.js";
import { textTools } from "./tools/text.js";
import { styleTools } from "./tools/styles.js";
import { tableTools } from "./tools/tables.js";
import { suggestionTools } from "./tools/suggestions.js";
import { structureTools } from "./tools/structure.js";
import { imageTools } from "./tools/images.js";
import { documentPropertyTools } from "./tools/documentProps.js";
import { commentTools } from "./tools/comments.js";
import { tabTools } from "./tools/documentTabs.js";

import type { ToolDefinition } from "./tools/shared.js";

export const TOOL_DEFINITIONS: Record<string, ToolDefinition> = {
  ...readTools,
  ...documentTools,
  ...textTools,
  ...styleTools,
  ...tableTools,
  ...suggestionTools,
  ...structureTools,
  ...imageTools,
  ...documentPropertyTools,
  ...commentTools,
  ...tabTools,
};
