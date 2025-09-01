#!/bin/bash

# This script starts the Enhanced Outlook MCP server.
# The server communicates over StdIO and will be started by the Claude desktop app.

# The authentication process will automatically start a temporary web server
# on port 3000 (by default) to handle the OAuth 2.0 callback.
# There is no need to run a separate authentication server.

# Change to the script's directory to ensure correct file resolution
cd "$(dirname "$0")"

# Start the main MCP server process
# Use exec to replace the shell process with the node process,
# ensuring signals are handled correctly.
exec node index.js