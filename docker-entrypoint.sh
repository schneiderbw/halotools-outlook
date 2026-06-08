#!/bin/sh
# Boot both processes in one container. nginx is the foreground process;
# the MCP server runs as a background child. tini (PID 1) handles signal
# forwarding and zombie reaping for both.

set -e

: "${MCP_PORT:=3001}"
export PORT="$MCP_PORT"

echo "halo-tools: starting MCP server on :$PORT" >&2
node /app/apps/mcp/dist/index.js --http &

echo "halo-tools: starting nginx on :80" >&2
exec nginx -g 'daemon off;'
