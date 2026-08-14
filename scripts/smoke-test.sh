#!/usr/bin/env bash
# Local container smoke test for the gdocs-mcp hosted MCP server.
#
# Builds the Docker image, runs it on a free local port, and exercises the
# minimum protocol surface required by the MintMCP hosted runtime:
#
#   1. GET  /healthz                     -> 200 with status=ok
#   2. POST /mcp initialize              -> returns protocolVersion
#   3. POST /mcp tools/list              -> returns the expected tool count (14)
#   4. POST /mcp tools/call search_docs  -> returns a structured 401 (fake token,
#                                          no crash) instead of an internal error
#
# Exits 0 on success, non-zero on failure. The container is always cleaned up
# via a trap, even on early failure.

set -euo pipefail

IMAGE_TAG="gdocs-mcp:smoke-$$"
CONTAINER_NAME="gdocs-mcp-smoke-$$"
EXPECTED_TOOLS=14

# Pick a free local port unless one was forced via SMOKE_PORT. We probe with a
# short-lived Python listen on :0 so the port is genuinely available right
# before `docker run -p` claims it. Falls back to 18000 if Python isn't around.
if [ -n "${SMOKE_PORT:-}" ]; then
  HOST_PORT="${SMOKE_PORT}"
elif command -v python3 >/dev/null 2>&1; then
  HOST_PORT="$(python3 -c 'import socket; s=socket.socket(); s.bind(("127.0.0.1",0)); print(s.getsockname()[1]); s.close()')"
else
  HOST_PORT=18000
fi
MCP_URL="http://127.0.0.1:${HOST_PORT}/mcp"
HEALTH_URL="http://127.0.0.1:${HOST_PORT}/healthz"

# Extract the JSON payload from a streamable-HTTP response. The transport
# answers with SSE (`data: <json>` lines); concatenate every data line so we
# don't miss multi-line events. Falls back to the raw body if no SSE framing
# was found (e.g., a JSON-only response).
extract_payload() {
  local resp="$1"
  local payload
  payload="$(echo "${resp}" | awk '/^data: /{ sub(/^data: /,""); print }' | tr -d '\n')"
  if [ -z "${payload}" ]; then
    payload="${resp}"
  fi
  echo "${payload}"
}

# Resolve repo root regardless of where the script is invoked from.
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

cleanup() {
  local exit_code=$?
  echo
  echo "[smoke] cleaning up..."
  docker rm -f "${CONTAINER_NAME}" >/dev/null 2>&1 || true
  docker rmi -f "${IMAGE_TAG}" >/dev/null 2>&1 || true
  exit "${exit_code}"
}
trap cleanup EXIT INT TERM

fail() {
  echo "[smoke] FAIL: $*" >&2
  exit 1
}

log() {
  echo "[smoke] $*"
}

# --- 1. Build the image -----------------------------------------------------
log "building image ${IMAGE_TAG}"
docker build -t "${IMAGE_TAG}" "${REPO_ROOT}" >/dev/null

# --- 2. Run the container ---------------------------------------------------
log "starting container on host port ${HOST_PORT}"
docker run -d --rm \
  --name "${CONTAINER_NAME}" \
  -p "${HOST_PORT}:8000" \
  "${IMAGE_TAG}" >/dev/null

# --- 3. Wait for healthz ---------------------------------------------------
log "waiting for /healthz to respond..."
ready=0
for _ in $(seq 1 30); do
  if curl -sf -o /dev/null "${HEALTH_URL}"; then
    ready=1
    break
  fi
  sleep 1
done
if [ "${ready}" -ne 1 ]; then
  docker logs "${CONTAINER_NAME}" >&2 || true
  fail "container did not become healthy within 30s"
fi
log "healthz OK"

# --- 4. Verify healthz body ------------------------------------------------
health_body="$(curl -sf "${HEALTH_URL}")"
echo "${health_body}" | grep -q '"status":"ok"' \
  || fail "/healthz returned unexpected body: ${health_body}"

# --- 5. POST /mcp initialize -----------------------------------------------
log "POST /mcp initialize"
init_resp="$(curl -sf -X POST "${MCP_URL}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-smoke-token" \
  -d '{"jsonrpc":"2.0","method":"initialize","id":1,"params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"1"}}}')"

init_json="$(extract_payload "${init_resp}")"
echo "${init_json}" | grep -q '"protocolVersion"' \
  || fail "initialize did not return protocolVersion. body=${init_resp}"
log "initialize OK"

# --- 6. POST /mcp tools/list -----------------------------------------------
log "POST /mcp tools/list"
list_resp="$(curl -sf -X POST "${MCP_URL}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-smoke-token" \
  -d '{"jsonrpc":"2.0","method":"tools/list","id":2,"params":{}}')"
list_json="$(extract_payload "${list_resp}")"
# Count tools by matching every `"name":"<x>"` inside `"tools":[...]`.
tool_count="$(echo "${list_json}" | grep -o '"name":"[^"]\+"' | wc -l | tr -d ' ')"
if [ "${tool_count}" != "${EXPECTED_TOOLS}" ]; then
  echo "${list_json}" >&2
  fail "expected ${EXPECTED_TOOLS} tools, got ${tool_count}"
fi
log "tools/list OK (${tool_count} tools)"

# --- 7. POST /mcp tools/call with fake bearer ------------------------------
# A bogus access token should produce a structured tool error (Drive returns
# 401) — NOT a server crash or 5xx. We accept either:
#   a) jsonrpc result with isError:true and an error payload, or
#   b) jsonrpc error envelope referencing auth.
log "POST /mcp tools/call search_documents (with fake bearer)"
call_resp="$(curl -sf -X POST "${MCP_URL}" \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "Authorization: Bearer fake-smoke-token" \
  -d '{"jsonrpc":"2.0","method":"tools/call","id":3,"params":{"name":"search_documents","arguments":{"name":"smoke-test-nonexistent"}}}')"
call_json="$(extract_payload "${call_resp}")"
# Must not be a 5xx-style internal error envelope. Accept any of:
#   - "isError":true (tool-level structured error)
#   - "Authentication failed" / "401" / "Invalid Credentials" (Google's message)
if echo "${call_json}" | grep -qE '"isError":true|Authentication failed|"status":401|Invalid Credentials|Missing Google access token'; then
  log "tools/call returned a structured auth error (good)"
else
  echo "${call_json}" >&2
  fail "tools/call did not return a structured auth error"
fi

log "ALL SMOKE TESTS PASSED"
