/**
 * Drive label resolution. VENDORED byte-identical across the Google connectors
 * (gdrive-mcp, gdocs-mcp, gsheets-mcp); mirror any change into the siblings.
 * Returns a file's applied Drive labels as human-readable strings for policy
 * middleware (keyed on `_meta.labels` / `_meta.labelsError`). Needs
 * `drive.labels.readonly`, else fails soft.
 */

const GOOGLE_DRIVE_API = 'https://www.googleapis.com/drive/v3';
const DRIVE_LABELS_API = 'https://drivelabels.googleapis.com/v2';
const LABEL_SCHEMA_TTL_MS = 60 * 60 * 1000;
const LABEL_SCHEMA_CACHE_MAX = 500;

async function driveGet(endpoint: string, accessToken: string): Promise<any> {
  const url = endpoint.startsWith('http') ? endpoint : `${GOOGLE_DRIVE_API}${endpoint}`;
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' },
  });
  if (!response.ok) {
    const body = await response.text().catch(() => '');
    throw new Error(`Drive labels API error (${response.status})${body ? `: ${body}` : ''}`);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

export type LabelSchema = Record<string, Record<string, string>>;

const labelSchemaCache = new Map<string, { schema: LabelSchema; expiresAt: number }>();

/** Test hook: cache is module state; lets tests isolate runs. */
export function clearLabelSchemaCache(): void {
  labelSchemaCache.clear();
}

/** Cache by the exact revision the file references, so a later choice rename can't
 *  retroactively change how an already-labeled file resolves. */
export async function getLabelSchema(
  labelId: string,
  revisionId: string | undefined,
  accessToken: string
): Promise<LabelSchema> {
  const labelResource = revisionId
    ? `${encodeURIComponent(labelId)}@${encodeURIComponent(revisionId)}`
    : encodeURIComponent(labelId);
  const cached = labelSchemaCache.get(labelResource);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.schema;
  }

  const data = await driveGet(
    `${DRIVE_LABELS_API}/labels/${labelResource}?view=LABEL_VIEW_FULL`,
    accessToken
  ) as {
    fields?: Array<{
      id: string;
      selectionOptions?: { choices?: Array<{ id: string; properties?: { displayName?: string } }> };
    }>;
  };

  const schema: LabelSchema = {};
  for (const field of data?.fields || []) {
    const choices = field.selectionOptions?.choices;
    if (!choices) continue;
    const choiceNames: Record<string, string> = {};
    for (const choice of choices) {
      if (choice.properties?.displayName) {
        choiceNames[choice.id] = choice.properties.displayName;
      }
    }
    schema[field.id] = choiceNames;
  }

  if (labelSchemaCache.size >= LABEL_SCHEMA_CACHE_MAX) {
    const oldest = labelSchemaCache.keys().next().value;
    if (oldest !== undefined) labelSchemaCache.delete(oldest);
  }
  labelSchemaCache.set(labelResource, { schema, expiresAt: Date.now() + LABEL_SCHEMA_TTL_MS });
  return schema;
}

/**
 * The file's Drive labels as human-readable strings (selection-choice names + text values;
 * date/integer/user fields skipped). `error` flags a real read/resolve failure, not a skip.
 * INVARIANT: never rejects; callers may leave the promise unawaited, so keep throws in try/catch.
 */
export async function getFileLabels(
  fileId: string,
  accessToken: string
): Promise<{ labels: string[]; error?: string }> {
  type AppliedLabel = {
    id: string;
    revisionId?: string;
    fields?: Record<string, { valueType?: string; selection?: string[]; text?: string[] }>;
  };

  const applied: AppliedLabel[] = [];
  const labels: string[] = [];
  let error: string | undefined;

  try {
    let pageToken: string | undefined;
    do {
      const params = new URLSearchParams({ maxResults: '100' });
      if (pageToken) params.set('pageToken', pageToken);
      const page = await driveGet(
        `/files/${encodeURIComponent(fileId)}/listLabels?${params}`,
        accessToken
      ) as { labels?: AppliedLabel[]; nextPageToken?: string };
      if (page?.labels) applied.push(...page.labels);
      pageToken = page?.nextPageToken;
    } while (pageToken);
  } catch (err: any) {
    error = 'label read failed';
    console.warn(`getFileLabels: label read failed fileId=${fileId} error=${err?.message}`);
  }

  for (const label of applied) {
    const fields = Object.entries(label.fields || {});

    for (const [, field] of fields) {
      if (field.valueType === 'text' && field.text) {
        labels.push(...field.text.filter(Boolean));
      }
    }

    const selectionFields = fields.filter(([, f]) => f.valueType === 'selection' && f.selection?.length);
    if (selectionFields.length === 0) continue;
    try {
      const schema = await getLabelSchema(label.id, label.revisionId, accessToken);
      for (const [fieldId, field] of selectionFields) {
        for (const choiceId of field.selection!) {
          const name = schema[fieldId]?.[choiceId];
          if (name) {
            labels.push(name);
          } else {
            labels.push(choiceId);
            if (!error) error = 'incomplete label resolution';
            console.warn(`getFileLabels: unresolved choice fileId=${fileId} labelId=${label.id}`);
          }
        }
      }
    } catch (e: any) {
      if (!error) error = 'incomplete label resolution';
      console.warn(`getFileLabels: label resolve failed fileId=${fileId} labelId=${label.id} error=${e?.message}`);
    }
  }

  const deduped = [...new Set(labels)];
  return error ? { labels: deduped, error } : { labels: deduped };
}
