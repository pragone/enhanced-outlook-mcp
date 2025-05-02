# Enhanced Outlook MCP Authentication

This directory contains the authentication module for the Enhanced Outlook MCP server.

## Authentication System

The authentication system (`auth-service.js` and related files) provides a simple, robust approach to Microsoft authentication using MSAL directly, similar to the approach used in the OneNote MCP server.

Key features:
- Self-contained authentication flow that doesn't require an external auth server
- Uses device code flow for interactive authentication
- Implements token caching with file persistence
- Falls back gracefully between silent token acquisition and interactive auth
- Directly integrates with Microsoft Graph API

## Usage

### Basic Usage

```javascript
// Import the auth service
const { getAuthService } = require('./auth/index');

// Get and initialize the authentication service
const authService = getAuthService();
await authService.initialize();

// Check if authenticated
const isAuthenticated = await authService.isAuthenticated();

// Get the Microsoft Graph client for making API calls
const graphClient = authService.getGraphClient();
```

### Tool Handlers

The authentication system provides MCP tool handlers:

- `authenticateHandler`: Initiates authentication flow
- `checkAuthStatusHandler`: Checks the current authentication status
- `revokeAuthenticationHandler`: Signs out the user

### Configuration

The authentication system uses the following configuration from your `.env` file:

```
MS_CLIENT_ID=your-client-id
MS_AUTHORITY=https://login.microsoftonline.com/common
MS_SCOPES=offline_access,User.Read,Mail.Read,Mail.ReadWrite,Mail.Send,Calendars.ReadWrite
```

## Testing

To test the authentication system, run:

```
node auth/test-auth-service.js
```

This will initiate the authentication flow and verify that the token can be used to access the Microsoft Graph API.

## Architecture

The authentication system consists of:
1. `auth-service.js`: Core MSAL authentication functionality
2. `tools-api.js`: Tool handlers for MCP
3. `index.js`: Exports for use in the main application
4. `test-auth-service.js`: Test script for verifying functionality 