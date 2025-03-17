#!/bin/bash

# Start the main MCP server
node index.js &
MCP_PID=$!

# Start the authentication server
node outlook-auth-server.js &
AUTH_PID=$!

# Handle shutdown
function cleanup {
  echo "Shutting down..."
  kill $MCP_PID
  kill $AUTH_PID
  exit 0
}

# Trap SIGINT (Ctrl+C)
trap cleanup SIGINT

echo "Both servers started. Press Ctrl+C to stop."
wait