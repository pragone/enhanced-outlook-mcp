# Enhanced Outlook MCP Server

This is an enhanced, modular implementation of the Outlook MCP (Model Context Protocol) server that connects Claude with Microsoft Outlook through the Microsoft Graph API. This server provides a robust set of tools for email, calendar, folder management, and rule creation.

## Features

- **Complete Authentication System**: OAuth 2.0 authentication with Microsoft Graph API with token refresh and multiple user support
- **Email Management**: List, search, read, send, and organize emails with attachment support
- **Calendar Integration**: Create, modify, and manage calendar events with attendee tracking
- **Folder Organization**: Create, manage, and navigate email folders
- **Rules Engine**: Create and manage complex mail processing rules
- **Modular Architecture**: Clean separation of concerns for better maintainability and extensibility
- **Enhanced Error Handling**: Detailed error messages and logging
- **Test Mode**: Simulated responses for testing without real API calls
- **Rate Limiting**: Prevent API throttling with built-in rate limiting
- **Multi-environment Configuration**: Support for development, testing, and production environments

## Directory Structure

```
/enhanced-outlook-mcp/
├── index.js                     # Main entry point
├── config.js                    # Configuration settings
├── .env.example                 # Example environment variables
├── auth/                        # Authentication modules
│   ├── index.js                 # Authentication exports
│   ├── token-manager.js         # Token storage and refresh
│   ├── multi-user-support.js    # Multiple user support
│   ├── tools.js                 # Auth-related tools
│   ├── auth-service.js          # Authentication service
│   ├── unified-index.js         # Backward compatibility
│   └── tools-api.js             # Auth-related tools API
├── email/                       # Email functionality
│   ├── index.js                 # Email exports
│   ├── list.js                  # List emails
│   ├── search.js                # Search emails
│   ├── read.js                  # Read email
│   ├── send.js                  # Send email
│   └── attachments.js           # Handle email attachments
├── calendar/                    # Calendar functionality
│   ├── index.js                 # Calendar exports
│   ├── create-event.js          # Create calendar events
│   ├── list-events.js           # List calendar events
│   ├── update-event.js          # Update calendar events
│   └── delete-event.js          # Delete calendar events
├── folder/                      # Folder management
│   ├── index.js                 # Folder exports
│   ├── list.js                  # List folders
│   ├── create.js                # Create folders
│   └── move.js                  # Move items between folders
├── rules/                       # Mail rules functionality
│   ├── index.js                 # Rules exports
│   ├── create.js                # Create mail rules
│   ├── list.js                  # List mail rules
│   └── delete.js                # Delete mail rules
└── utils/                       # Utility functions
    ├── enhanced-graph-api.js    # Enhanced Graph API client
    ├── graph-api-adapter.js     # Adapter for backward compatibility
    ├── odata-helpers.js         # OData query building
    ├── logger.js                # Logging utility
    ├── rate-limiter.js          # API rate limiting
    └── mock-data/               # Test mode mock data
        ├── emails.js            # Mock email data
        ├── folders.js           # Mock folder data
        ├── calendar.js          # Mock calendar data
        └── rules.js             # Mock rules data
```

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/your-username/enhanced-outlook-mcp.git
   cd enhanced-outlook-mcp
   ```

2. Install dependencies:
   ```
   npm install
   ```

3. Create a `.env` file based on `.env.example` with your Microsoft App Registration details:
   ```
   MS_CLIENT_ID=your_client_id
   MS_CLIENT_SECRET=your_client_secret
   # Additional configuration options
   ```

## Usage with Claude

1. Configure Claude to use the MCP server by adding the following to your Claude configuration:
   ```json
   {
     "tools": [
       {
         "name": "enhanced-outlook-mcp",
         "url": "http://localhost:3000",
         "auth": {
           "type": "none"
         }
       }
     ]
   }
   ```

2. Start the MCP server:
   ```
   npm start
   ```

3. Use the authenticate tool in Claude to initiate the authentication flow.

## Authentication Flow

1. Use the `authenticate` tool in Claude to get an authentication URL
2. Complete the authentication in your browser
3. Tokens are securely stored in the configured location

## Development

To run the server in development mode with auto-reload:
```
npm run dev
```

To run tests:
```
npm test
```

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Authentication System Migration Plan

The current authentication system is being updated to adopt a cleaner architecture inspired by the OneNote MCP server while maintaining the current browser-based authentication flow. This will address reliability and complexity issues without sacrificing user experience.

### Goals

- Create a simpler, more robust authentication system
- Eliminate the need for a separate authentication server
- Maintain user-friendly web-based authentication (no device code flows)
- Improve code maintainability and error handling
- Enable smoother token refresh operations

### Implementation Plan

#### Phase 1: Core Authentication Module ✅

1. Create a new `AuthService` class that encapsulates authentication logic:
   - Token caching with file persistence
   - Silent token acquisition
   - Web-based authentication with browser redirect
   - Token refresh handling
   - Microsoft Graph client initialization

2. Develop a lightweight embedded HTTP server for OAuth callbacks:
   - Handle redirect from Microsoft authentication
   - Exchange authorization code for tokens
   - Store tokens securely using the same mechanisms as the current system

3. Create new tool handlers that utilize the improved authentication service:
   - `authenticateHandler`: Initiates authentication flow
   - `checkAuthStatusHandler`: Checks authentication status
   - `revokeAuthenticationHandler`: Handles sign-out

#### Phase 2: Graph Client Integration ✅

1. Integrate the authentication service with Microsoft Graph:
   - Create authenticated client instances
   - Handle token expiration and refresh
   - Provide consistent error handling

2. Create adapters for existing API helpers to use the new authentication system:
   - Ensure backward compatibility
   - Improve error reporting

#### Phase 3: Migration

1. Implement side-by-side operation:
   - Allow both authentication systems to run concurrently
   - Gradually migrate features to the new system

2. Test and validate:
   - Ensure all auth scenarios work correctly
   - Verify token persistence and refresh
   - Test multi-user scenarios

#### Phase 4: Cleanup

1. Remove dependencies on external auth server
2. Update documentation and examples
3. Deprecate old authentication modules

### Benefits

- **Self-contained**: No need for a separate auth server
- **Simplified Code**: Cleaner architecture with better separation of concerns
- **Improved Reliability**: Better error handling and recovery
- **Same User Experience**: Maintains the familiar browser-based authentication flow
- **Better Maintainability**: More modular design for easier future updates

## Authentication System Update

The authentication system has been fully migrated to the new implementation. The old authentication system has been removed entirely as of May 2024. All users should now be using the new authentication flow.

Benefits of the new authentication system:
- More reliable token management
- Better error handling
- Cleaner API interactions
- Improved security
- Self-contained authentication (no separate auth server needed)

If you experience any authentication issues, please run the `authenticate` tool to create a new authentication session.

## Authentication Migration Completed

As of June 2024, the authentication system migration has been completed with all phases executed:

- Phase 1: Core Authentication Module ✅
- Phase 2: Graph Client Integration ✅ 
- Phase 3: Migration ✅
- Phase 4: Cleanup ✅

The migration has successfully:
- Eliminated dependencies on the external auth server
- Updated documentation and examples
- Removed old authentication modules
- Simplified the authentication flow for users

All authentication now happens through the integrated authentication service.