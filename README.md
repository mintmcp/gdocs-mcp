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

## Tool surface (14 tools)

- **Discovery** — `search_documents` (finds Word uploads too; `mimeType` says which)
- **Read** — `get_document` (with optional `include_structure`, `include_comments`, `include_table_styles`, multi-tab summary via `includeTabsContent`), `get_document_images`
- **Create** — `create_document` (optional initial body, optional parent folder)
- **Insert / append text** — `insert_text` (index-based), `append_text` (end-of-doc)
- **Tables** — `append_table` (end-of-doc), `insert_table` (index-based, refuses an index inside an existing table unless `allow_nested`), `update_table_style` (background, per-side borders, vertical alignment)
- **Update text** — `replace_text`, `delete_content`, `update_text_style` (bold/italic/underline/strikethrough/link), `update_paragraph_style` (heading level, alignment)
- **Word uploads** — `convert_to_google_doc` turns a `.doc`/`.docx` into a new native Doc, leaving the original untouched

`get_document` reads `.doc` and `.docx` uploads by parsing their bytes, since
Drive's export endpoint 403s on them. Those files are read-only here: every
write tool refuses them and points at `convert_to_google_doc`.

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
# List tools (should return 14).
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
