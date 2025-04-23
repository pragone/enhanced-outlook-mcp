const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const config = require('./config');
const logger = require('./utils/logger');
const fs = require('fs');
const { z } = require('zod');

// Import handlers from email module
const { 
  listEmailsHandler, 
  searchEmailsHandler,
  readEmailHandler, 
  markEmailHandler,
  sendEmailHandler, 
  createDraftHandler, 
  replyEmailHandler, 
  forwardEmailHandler,
  getAttachmentHandler, 
  listAttachmentsHandler, 
  addAttachmentHandler, 
  deleteAttachmentHandler
} = require('./email');

// Import handlers from auth module
const { 
  authenticateHandler, 
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler
} = require('./auth');

// Import handlers from calendar module
const {
  listEventsHandler,
  getEventHandler,
  listCalendarsHandler,
  createEventHandler,
  updateEventHandler,
  respondToEventHandler,
  deleteEventHandler,
  cancelEventHandler,
  findMeetingTimesHandler
} = require('./calendar');

// Import handlers from folder module
const {
  listFoldersHandler,
  getFolderHandler,
  createFolderHandler,
  updateFolderHandler,
  deleteFolderHandler,
  moveEmailsHandler,
  moveFolderHandler,
  copyEmailsHandler
} = require('./folder');

// Import handlers from rules module
const {
  listRulesHandler,
  getRuleHandler,
  createRuleHandler,
  updateRuleHandler,
  deleteRuleHandler
} = require('./rules');

// Create MCP server instance
const server = new McpServer({
  name: config.server.name,
  version: config.server.version,
  port: config.server.port,
  
  onRequest: async (req) => {
    logger.info(`Received request for tool: ${req.tool}`);
    
    // Rate limiting check
    try {
      const { rateLimiter } = require('./utils/rate-limiter');
      await rateLimiter.check();
    } catch (error) {
      logger.warn(`Rate limit exceeded: ${error.message}`);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'rate_limit_exceeded',
              message: 'Too many requests. Please try again later.',
              retry_after: error.retryAfter
            }
          })
        }]
      };
    }
  },
  
  onError: (error, req) => {
    logger.error(`Error in MCP request for tool ${req?.tool}:`, error);
    
    // Provide more informative error response based on error type
    if (error.name === 'AuthenticationError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'authentication_error',
              message: 'Authentication failed. Please re-authenticate using the authenticate tool.',
            }
          })
        }]
      };
    }
    
    if (error.name === 'GraphAPIError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'api_error',
              message: `Microsoft Graph API error: ${error.message}`,
              code: error.code
            }
          })
        }]
      };
    }
    
    // Generic error fallback
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          error: {
            type: 'internal_error',
            message: 'An unexpected error occurred. Please try again or check the server logs.',
          }
        })
      }]
    };
  }
});

// Helper function for error handling in tool handlers
const withErrorHandling = (handler) => async (params) => {
  try {
    // Rate limiting check
    const { rateLimiter } = require('./utils/rate-limiter');
    await rateLimiter.check();
    
    // For Claude Desktop compatibility
    if (global.__last_message?.params?.arguments) {
      params = { ...params, ...global.__last_message.params.arguments };
    }
    
    return await handler(params);
  } catch (error) {
    logger.error(`Error executing tool:`, error);
    
    if (error.name === 'AuthenticationError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'authentication_error',
              message: 'Authentication failed. Please re-authenticate using the authenticate tool.',
            }
          })
        }]
      };
    }
    
    if (error.name === 'GraphAPIError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'api_error',
              message: `Microsoft Graph API error: ${error.message}`,
              code: error.code
            }
          })
        }]
      };
    }
    
    // Generic error fallback
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          error: {
            type: 'internal_error',
            message: 'An unexpected error occurred. Please try again or check the server logs.',
          }
        })
      }]
    };
  }
};

// Register tools directly with inline parameters
// Email tools
server.tool(
  "read_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    id: z.string().optional().describe('The ID of the email message to read'),
    messageId: z.string().optional().describe('Alternative parameter name for email ID'),
    message_id: z.string().optional().describe('Snake case alternative parameter name for email ID'),
    emailId: z.string().optional().describe('Alternative parameter name for email ID'),
    markAsRead: z.boolean().optional().describe('Whether to mark the email as read')
  },
  withErrorHandling(readEmailHandler),
  {
    description: 'Read a specific email by ID',
    usage: 'Use to get the full content of an email when you have its ID'
  }
);

server.tool(
  "list_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().optional().describe('Folder ID or well-known folder name (inbox, drafts, sentitems, deleteditems)'),
    limit: z.number().optional().describe('Maximum number of emails to return'),
    skip: z.number().optional().describe('Number of emails to skip (for pagination)'),
    orderBy: z.union([z.string(), z.array(z.string()), z.object({})]).optional()
      .describe('OData orderby specification (defaults to receivedDateTime desc)'),
    fields: z.union([z.string(), z.array(z.string())]).optional()
      .describe('Fields to include in the response'),
    search: z.string().optional().describe('Search query')
  },
  withErrorHandling(listEmailsHandler),
  {
    description: 'List emails from a mailbox folder',
    usage: 'Use to list emails from a specific folder'
  }
);

server.tool(
  "search_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    query: z.string().describe('Search query'),
    limit: z.number().optional().describe('Maximum number of emails to return'),
    fields: z.union([z.string(), z.array(z.string())]).optional()
      .describe('Fields to include in the response')
  },
  withErrorHandling(searchEmailsHandler),
  {
    description: 'Search for emails across all folders',
    usage: 'Use to search for emails matching specific criteria'
  }
);

// Authentication tools
server.tool(
  "authenticate",
  {
    userId: z.string().optional().describe('User identifier for multi-user support'),
    scopes: z.array(z.string()).optional().describe('OAuth scopes to request')
  },
  withErrorHandling(authenticateHandler),
  {
    description: 'Authenticate with Microsoft Outlook using OAuth',
    usage: 'Use to authenticate the user before accessing email data'
  }
);

// Add more auth tools
server.tool(
  "check_auth_status",
  {},
  withErrorHandling(checkAuthStatusHandler),
  {
    description: 'Check authentication status',
    usage: 'Use to verify if the user is authenticated'
  }
);

server.tool(
  "revoke_authentication",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")')
  },
  withErrorHandling(revokeAuthenticationHandler),
  {
    description: 'Revoke authentication and delete stored tokens',
    usage: 'Use to log out a user'
  }
);

server.tool(
  "list_authenticated_users",
  {},
  withErrorHandling(listAuthenticatedUsersHandler),
  {
    description: 'List all authenticated users',
    usage: 'Use to see which users are authenticated'
  }
);

// Add more email tools
server.tool(
  "mark_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('Email ID to mark'),
    isRead: z.boolean().optional().describe('Whether to mark as read (true) or unread (false)')
  },
  withErrorHandling(markEmailHandler),
  {
    description: 'Mark an email as read or unread',
    usage: 'Use to change the read status of an email'
  }
);

server.tool(
  "send_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type'),
    to: z.union([z.string(), z.array(z.string())]).describe('Recipient(s)'),
    cc: z.union([z.string(), z.array(z.string())]).optional().describe('CC recipient(s)'),
    bcc: z.union([z.string(), z.array(z.string())]).optional().describe('BCC recipient(s)')
  },
  withErrorHandling(sendEmailHandler),
  {
    description: 'Send a new email',
    usage: 'Use to send an email to one or more recipients'
  }
);

// Add some calendar tools
server.tool(
  "list_events",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    startDateTime: z.string().optional().describe('Start date and time in ISO format'),
    endDateTime: z.string().optional().describe('End date and time in ISO format'),
    limit: z.number().optional().describe('Maximum number of events to return')
  },
  withErrorHandling(listEventsHandler),
  {
    description: 'List calendar events within a date range',
    usage: 'Use to view upcoming calendar events'
  }
);

server.tool(
  "create_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    subject: z.string().describe('Event subject'),
    body: z.string().optional().describe('Event body content'),
    start: z.string().describe('Start time in ISO format'),
    end: z.string().optional().describe('End time in ISO format'),
    attendees: z.union([z.string(), z.array(z.string())]).optional().describe('Event attendees')
  },
  withErrorHandling(createEventHandler),
  {
    description: 'Create a new calendar event',
    usage: 'Use to schedule a new meeting or appointment'
  }
);

// Add some folder tools
server.tool(
  "list_folders",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    parentFolderId: z.string().optional().describe('Parent folder ID to list child folders')
  },
  withErrorHandling(listFoldersHandler),
  {
    description: 'List mail folders',
    usage: 'Use to view the folder structure in the mailbox'
  }
);

server.tool(
  "move_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailIds: z.array(z.string()).describe('Array of email IDs to move'),
    destinationFolderId: z.string().describe('Destination folder ID')
  },
  withErrorHandling(moveEmailsHandler),
  {
    description: 'Move emails to a folder',
    usage: 'Use to organize emails by moving them to different folders'
  }
);

// Add critical tools based on user needs and add more as required

const transport = new StdioServerTransport();

// Start the server
server.connect(transport).then(() => {
  logger.info(`${config.server.name} v${config.server.version} started on port ${config.server.port}`);
  
  if (config.testing.enabled) {
    logger.info('Server running in TEST MODE with mock data');
  }
  
  logger.info(`Authentication server should be running on port ${config.server.authPort}`);
}).catch(error => {
  logger.error('Failed to start MCP server:', error);
  process.exit(1);
});