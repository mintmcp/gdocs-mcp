# gdocs-mcp

A hosted MCP server that wraps the Google Docs and Drive APIs, packaged for
the [MintMCP](https://mintmcp.com) hosted runtime. Speaks streamable HTTP on
port 8000, reads the per-request user access token from the
`Authorization: Bearer <token>` header, and exposes 11 tools across document
discovery, reading, editing, commenting, and image extraction.

## Auth contract

MintMCP performs the OAuth flow on the frontend and forwards the user's
access token to this server on every request:

```
Authorization: Bearer <google-oauth-access-token>
```

That is the **only** credential input. No tenant/account IDs, no client
secrets — there is no per-deployment configuration to inject.

### Required Google OAuth scopes

| Scope | Purpose |
| --- | --- |
| `openid` | Identity assertion |
| `https://www.googleapis.com/auth/userinfo.email` | User email |
| `https://www.googleapis.com/auth/userinfo.profile` | User profile |
| `https://www.googleapis.com/auth/drive.readonly` | List/search docs, fetch metadata, read comments |
| `https://www.googleapis.com/auth/drive.file` | Create new docs in a folder |
| `https://www.googleapis.com/auth/documents` | Read/write document content |
| `https://www.googleapis.com/auth/drive.labels.readonly` | Optional, `labels` profile only: label enrichment (see below) |

### Profiles

A **profile** (`PROFILES` in `src/scopes.ts`) is a frozen, named scope set,
a connector's contract with its users:

| Profile    | Scopes                                                                  |
|------------|--------------------------------------------------------------------------|
| `standard` | `drive.readonly` + `drive.file` + `documents`                            |
| `labels`   | `drive.readonly` + `drive.file` + `documents` + `drive.labels.readonly`  |

Each deployment selects a profile via the `PROFILE` env var; unset registers
every tool (self-hosted default) and enriches best effort, an unknown or
empty value fails at boot. Both profiles register every tool: the label
scope gates the enrichment below, not a tool.

### Drive label enrichment

On deployments whose grant includes `drive.labels.readonly`, `get_document`
and `get_document_images` attach the document's applied Drive labels as
`_meta.applied`, one entry per label, `{labelId, revisionId, title,
resolved, values[], skippedValueTypes?}`, the same tree the gdrive-mcp
connector emits (`src/lib/driveLabels.ts` is vendored from it; keep them
identical). Labels stay in `_meta` only: this connector has no metadata
tool, the labels-enabled Drive connector is the fleet's visible label
console. `user` fields are withheld, text values cap at 256 chars, lookups
pin the applied label revision, and failures degrade to a `labelsError`
code instead of failing the read. Without the scope no label call is made
and no `_meta` is returned; absence means "surfacing not enabled", never
"no labels".

## Tool surface (11 tools)

- **Discovery** — `search_documents`
- **Read** — `get_document` (with optional `include_structure`, `include_comments`, multi-tab summary via `includeTabsContent`), `get_document_images`
- **Create** — `create_document` (optional initial body, optional parent folder)
- **Insert / append text** — `insert_text` (index-based), `append_text` (end-of-doc), `append_table` (with reverse-order cell insertion)
- **Update text** — `replace_text`, `delete_content`, `update_text_style` (bold/italic/underline/strikethrough/link), `update_paragraph_style` (heading level, alignment)

Round 2 added the multi-tab summary: `get_document` now returns a `tabs`
array when the doc uses Google Docs Tabs, plus a `headings` list and the
current `revisionId` for optimistic-concurrency-controlled writes via
`required_revision_id`. Round 2 also added structured `error.details`
passthrough in `toolErrorResponse` envelopes.

## Local build & run

```bash
docker build -t gdocs-mcp:dev .
docker run --rm -p 8000:8000 gdocs-mcp:dev
```

The container listens on `0.0.0.0:8000`. `/healthz` returns `{"status":"ok"}`;
MCP requests POST to `/mcp`.

## Verifying with curl

```bash
# List tools (should return 11).
curl -s -X POST http://localhost:8000/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer <google-access-token>" \
  -d '{"jsonrpc":"2.0","method":"tools/list","id":1,"params":{}}'
```

## Deploying to MintMCP

Build for `linux/amd64` and push:

```bash
docker buildx build \
  --platform linux/amd64 \
  -t mintmcp/gdocs-mcp:latest \
  --push .
```

Then deploy via the hosted-cli:

```bash
hosted-cli build-and-deploy mintmcp/gdocs-mcp:latest
```

## Development

```bash
npm install
npm run dev      # tsx watch on src/index.ts
npm run build    # tsc to dist/
npm test         # vitest unit tests (pure helpers)
npm run smoke    # docker build + protocol smoke test (requires Docker)
```

Pure document-structure helpers live in `src/lib/docs-structure.ts` and are
covered by `src/lib/docs-structure.test.ts` (42 tests).
