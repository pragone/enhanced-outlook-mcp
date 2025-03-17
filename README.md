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
│   └── tools.js                 # Auth-related tools
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
    ├── graph-api.js             # Microsoft Graph API helper
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

3. In a separate terminal, start the authentication server:
   ```
   npm run auth-server
   ```

4. Use the authenticate tool in Claude to initiate the authentication flow.

## Authentication Flow

1. Start the authentication server on the configured port (default: 3333)
2. Use the `authenticate` tool to get an authentication URL
3. Complete the authentication in your browser
4. Tokens are securely stored in the configured location

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