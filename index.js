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
  toolMetadata: config.toolMetadata,
  
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
    
    // Handle authentication errors
    if (error.name === 'AuthenticationError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'authentication_error',
              message: 'Authentication failed. Please re-authenticate using the authenticate tool.',
              suggested_tool: 'authenticate',
              suggested_sequence: 'None - authentication is a prerequisite'
            }
          })
        }]
      };
    }
    
    if (error.name === 'TokenExpiredError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'token_expired',
              message: 'Your authentication token has expired. Please re-authenticate using the authenticate tool.',
              suggested_tool: 'authenticate',
              suggested_sequence: 'None - authentication is a prerequisite'
            }
          })
        }]
      };
    }
    
    // Handle specific resource errors with suggested workflows
    if (error.name === 'CalendarNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Calendar not found. Please call list_calendars first to get valid calendar IDs.',
              suggested_tool: 'list_calendars',
              suggested_sequence: 'view_calendar_events'
            }
          })
        }]
      };
    }
    
    if (error.name === 'EmailNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Email not found. Please call list_emails or search_emails first to get valid email IDs.',
              suggested_tool: 'list_emails',
              suggested_sequence: 'view_emails'
            }
          })
        }]
      };
    }
    
    if (error.name === 'FolderNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Folder not found. Please call list_folders first to get valid folder IDs.',
              suggested_tool: 'list_folders',
              suggested_sequence: 'organize_emails'
            }
          })
        }]
      };
    }
    
    if (error.name === 'AttachmentNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Attachment not found. Please call list_attachments first to get valid attachment IDs.',
              suggested_tool: 'list_attachments',
              suggested_workflow: {
                sequence: ['check_auth_status', 'authenticate', 'list_emails', 'list_attachments', 'get_attachment'],
                conditional_steps: {
                  'authenticate': 'Only if auth_needed is true'
                }
              }
            }
          })
        }]
      };
    }
    
    if (error.name === 'EventNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Event not found. Please call list_events first to get valid event IDs.',
              suggested_tool: 'list_events',
              suggested_sequence: 'view_calendar_events'
            }
          })
        }]
      };
    }
    
    if (error.name === 'RuleNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Rule not found. Please call list_rules first to get valid rule IDs.',
              suggested_tool: 'list_rules',
              suggested_workflow: {
                sequence: ['check_auth_status', 'authenticate', 'list_rules', 'get_rule'],
                conditional_steps: {
                  'authenticate': 'Only if auth_needed is true'
                }
              }
            }
          })
        }]
      };
    }
    
    if (error.name === 'DraftNotFoundError') {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'prerequisite_needed',
              message: 'Draft email not found. Please call create_draft first to create a draft email.',
              suggested_tool: 'create_draft',
              suggested_sequence: 'send_email_with_attachment'
            }
          })
        }]
      };
    }
    
    if (error.name === 'ParameterError') {
      // Get the current tool name from the error context or params
      const currentTool = error.toolName || 'current_tool';
      
      // Check if this tool has dependencies in the metadata
      const dependencies = config.toolMetadata[currentTool]?.dependencies || [];
      
      if (dependencies.length > 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: {
                type: 'parameter_error',
                message: `Invalid parameters: ${error.message}`,
                tool_dependencies: dependencies,
                suggestion: 'This tool requires information from prerequisite tools. Try following the suggested workflow.',
                suggested_workflow: {
                  sequence: ['check_auth_status', ...dependencies, currentTool],
                  conditional_steps: {
                    'authenticate': 'Only if auth_needed is true'
                  }
                }
              }
            })
          }]
        };
      }
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'parameter_error',
              message: `Invalid parameters: ${error.message}`
            }
          })
        }]
      };
    }
    
    if (error.name === 'GraphAPIError') {
      let errorMessage = `Microsoft Graph API error: ${error.message}`;
      let suggestedTool = null;
      let suggestedSequence = null;
      
      // Add specific guidance for common API errors
      if (error.code === 'InvalidAuthenticationToken') {
        errorMessage = 'Your authentication token is invalid or expired. Please re-authenticate.';
        suggestedTool = 'authenticate';
      } else if (error.code === 'AccessDenied') {
        errorMessage = 'You do not have permission to perform this operation. Please check your Microsoft 365 permissions.';
        suggestedTool = 'authenticate';
      } else if (error.code === 'MailboxNotEnabledForRESTAPI') {
        errorMessage = 'This mailbox is not enabled for Microsoft Graph API access. Please contact your administrator.';
      }
      
      // Check if the error relates to a specific resource
      if (error.message.includes('calendar')) {
        suggestedSequence = 'view_calendar_events';
      } else if (error.message.includes('mail') || error.message.includes('message')) {
        suggestedSequence = 'view_emails';
      }
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'api_error',
              message: errorMessage,
              code: error.code,
              suggested_tool: suggestedTool,
              suggested_sequence: suggestedSequence
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
    id: z.string().optional().describe('The ID of the email message to read - obtained from list_emails or search_emails'),
    messageId: z.string().optional().describe('Alternative parameter name for email ID - obtained from list_emails or search_emails'),
    message_id: z.string().optional().describe('Snake case alternative parameter name for email ID - obtained from list_emails or search_emails'),
    emailId: z.string().optional().describe('Alternative parameter name for email ID - obtained from list_emails or search_emails'),
    markAsRead: z.boolean().optional().describe('Whether to mark the email as read when retrieving it')
  },
  withErrorHandling(readEmailHandler),
  {
    description: 'Read a specific email by ID',
    usage: `Use to get the full content of an email when you have its ID. Call check_auth_status first to determine if authentication is needed, then list_emails or search_emails to get email IDs before using this tool.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_emails to get a list of emails with their IDs
4) Then read_email with parameters:
   {
     "id": "AAMkADE1...",
     "markAsRead": true
   }
   
For keeping the email as unread:
1) Call check_auth_status
2) Call authenticate if needed
3) Call list_emails or search_emails to find the email
4) Call read_email with markAsRead=false`
  }
);

server.tool(
  "list_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().optional().describe('Folder ID or well-known folder name (inbox, drafts, sentitems, deleteditems) - obtain specific folder IDs from list_folders'),
    limit: z.number().optional().describe('Maximum number of emails to return (use for pagination)'),
    skip: z.number().optional().describe('Number of emails to skip (use with limit for pagination)'),
    orderBy: z.union([z.string(), z.array(z.string()), z.object({})]).optional()
      .describe('OData orderby specification (defaults to receivedDateTime desc)'),
    fields: z.union([z.string(), z.array(z.string())]).optional()
      .describe('Fields to include in the response - comma-separated list or array of field names'),
    search: z.string().optional().describe('Search query to filter results (uses server-side search)')
  },
  withErrorHandling(listEmailsHandler),
  {
    description: 'List emails from a mailbox folder',
    usage: `Use to list emails from a specific folder. Call check_auth_status first to determine if authentication is needed, then optionally list_folders if you need to work with a custom folder rather than default ones like "inbox" or "sentitems".

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then list_emails with parameters:
   {
     "folderId": "inbox", 
     "limit": 10,
     "orderBy": "receivedDateTime desc"
   }
   
For custom folders:
1) Call check_auth_status
2) Call authenticate if needed
3) Call list_folders to get folder IDs
4) Call list_emails with the specific folderId`
  }
);

server.tool(
  "search_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    query: z.string().describe('Search query to find specific emails across folders'),
    limit: z.number().optional().describe('Maximum number of emails to return (use for pagination)'),
    fields: z.union([z.string(), z.array(z.string())]).optional()
      .describe('Fields to include in the response - comma-separated list or array of field names')
  },
  withErrorHandling(searchEmailsHandler),
  {
    description: 'Search for emails across all folders',
    usage: `Use to search for emails matching specific criteria across the entire mailbox. Call check_auth_status first to determine if authentication is needed. For folder-specific searches, use list_emails with the search parameter instead.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then search_emails with parameters:
   {
     "query": "monthly report",
     "limit": 10,
     "fields": ["id", "subject", "receivedDateTime", "from", "bodyPreview"]
   }
   
For narrowing results:
1) Call check_auth_status
2) Call authenticate if needed
3) Call search_emails with a specific query
4) Call read_email with IDs from the search results`
  }
);

// Authentication tools
server.tool(
  "authenticate",
  {
    userId: z.string().optional().describe('User identifier for multi-user scenarios (defaults to "default")'),
    forceNewAuth: z.boolean().optional().describe('Force new authentication flow even if valid tokens exist')
  },
  withErrorHandling(authenticateHandler),
  {
    description: 'Authenticate with Microsoft Graph API',
    usage: `Use this tool only when check_auth_status indicates authentication is needed. Authentication tokens are cached, so this doesn't need to be called every time.

Example:
1) First call check_auth_status to determine if authentication is needed:
   {
     "userId": "default"
   }
   
2) If auth_needed is true, then call authenticate:
   {
     "userId": "default"
   }
   
3) Follow the authentication URL provided in the response
4) Complete the sign-in process in your browser
5) Return to continue working with other tools

For force re-authentication (rarely needed):
1) Call authenticate with parameters:
   {
     "forceNewAuth": true
   }`
  }
);

// Add more auth tools
server.tool(
  "check_auth_status",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")')
  },
  async (params) => {
    try {
      const result = await checkAuthStatusHandler(params);
      
      // Enhance the response with more information
      const responseData = JSON.parse(result.content[0].text);
      
      // Add additional fields to help Claude understand auth status better
      if (responseData.authenticated) {
        responseData.auth_needed = false;
        responseData.message = "Authentication already complete. No need to authenticate again.";
        responseData.token_expiry = responseData.tokenExpiresAt || null;
        
        // Calculate if token will expire soon (within 5 minutes)
        if (responseData.tokenExpiresAt) {
          const expiryDate = new Date(responseData.tokenExpiresAt);
          const now = new Date();
          const timeUntilExpiry = expiryDate - now;
          const fiveMinutes = 5 * 60 * 1000;
          
          if (timeUntilExpiry < fiveMinutes && timeUntilExpiry > 0) {
            responseData.auth_needed = true;
            responseData.message = "Authentication token will expire soon. Consider re-authenticating.";
          }
        }
      } else {
        responseData.auth_needed = true;
        responseData.message = "Authentication required. Please call authenticate tool.";
      }
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify(responseData)
        }]
      };
    } catch (error) {
      logger.error('Error checking auth status:', error);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            authenticated: false,
            auth_needed: true,
            message: "Error checking authentication status. Please authenticate.",
            error: error.message
          })
        }]
      };
    }
  },
  {
    description: 'Check authentication status',
    usage: `Use this tool to check if authentication is needed before performing operations. It tells you whether authentication is required or if existing cached credentials can be used.

Example:
1) Call check_auth_status:
   {
     "userId": "default"
   }

This will return information about whether authentication is needed, along with details about token validity and expiration.

If auth_needed is false, you can proceed directly to other tools without calling authenticate first.`
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
    usage: `Use to log out a user and remove their stored authentication tokens. Call this when you're done working with sensitive data or when you need to switch authentication contexts.

Example:
1) Call revoke_authentication with parameters:
   {
     "userId": "default"
   }
   
After revoking authentication, you'll need to call authenticate again before accessing protected resources.`
  }
);

server.tool(
  "list_authenticated_users",
  {},
  withErrorHandling(listAuthenticatedUsersHandler),
  {
    description: 'List all authenticated users',
    usage: `Use to see which users are authenticated in the system. This tool does not require authentication.

Example:
1) Call list_authenticated_users with no parameters
   {}
   
This returns a list of all user IDs that have stored authentication tokens.`
  }
);

// Add more email tools
server.tool(
  "mark_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    id: z.string().describe('Email ID to mark - obtained from list_emails or search_emails'),
    isRead: z.boolean().optional().describe('Whether to mark the email as read (true) or unread (false)'),
    isFlagged: z.boolean().optional().describe('Whether to flag (true) or unflag (false) the email'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Set the importance level of the email')
  },
  withErrorHandling(markEmailHandler),
  {
    description: 'Update email properties like read/unread status or importance',
    usage: 'Use to change the status of an email message. Call authenticate first, then list_emails or search_emails to get the email ID before using this tool. At least one of isRead, isFlagged, or importance must be provided.'
  }
);

server.tool(
  "send_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content - can be plain text or HTML depending on bodyType'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type (Text or HTML)'),
    to: z.union([z.string(), z.array(z.string())]).describe('Recipient email address(es) - string or array of strings'),
    cc: z.union([z.string(), z.array(z.string())]).optional().describe('CC recipient email address(es) - string or array of strings'),
    bcc: z.union([z.string(), z.array(z.string())]).optional().describe('BCC recipient email address(es) - string or array of strings')
  },
  withErrorHandling(sendEmailHandler),
  {
    description: 'Send a new email',
    usage: `Use to send an email to one or more recipients. Call check_auth_status first to determine if authentication is needed. For emails with attachments, create a draft with create_draft first, then add attachments with add_attachment before sending.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then send_email with parameters:
   {
     "subject": "Meeting Agenda",
     "body": "<p>Here's the agenda for our meeting tomorrow.</p><p>Looking forward to it!</p>",
     "bodyType": "HTML",
     "to": "recipient@example.com",
     "cc": ["cc1@example.com", "cc2@example.com"]
   }
   
For emails with attachments:
1) Call check_auth_status
2) Call authenticate if needed
3) Call create_draft to create a draft email
4) Call add_attachment for each attachment
5) Call send_email with the draft ID`
  }
);

// Add some calendar tools
server.tool(
  "list_events",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    startDateTime: z.string().optional().describe('Start date and time in ISO format (e.g., "2023-11-01T00:00:00Z")'),
    endDateTime: z.string().optional().describe('End date and time in ISO format (e.g., "2023-11-30T23:59:59Z")'),
    limit: z.number().optional().describe('Maximum number of events to return'),
    calendarId: z.string().optional().describe('Specific calendar ID (obtain from list_calendars first). Required when working with non-default calendars.')
  },
  withErrorHandling(listEventsHandler),
  {
    description: 'List calendar events within a date range',
    usage: 'Use to view upcoming calendar events. Call authenticate first, then list_calendars if you need to work with a specific calendar. Results can be used with get_event to view event details.'
  }
);

server.tool(
  "create_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    subject: z.string().describe('Event subject/title'),
    body: z.string().optional().describe('Event body/description content'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type (Text or HTML)'),
    start: z.string().describe('Start date and time in ISO format (e.g., "2023-11-15T09:00:00Z")'),
    end: z.string().describe('End date and time in ISO format (e.g., "2023-11-15T10:00:00Z")'),
    location: z.string().optional().describe('Event location'),
    attendees: z.array(z.object({
      email: z.string(),
      name: z.string().optional(),
      type: z.enum(['required', 'optional']).optional()
    })).optional().describe('List of attendees with their email addresses'),
    isOnlineMeeting: z.boolean().optional().describe('Whether this is an online meeting'),
    calendarId: z.string().optional().describe('Specific calendar ID (obtain from list_calendars first). Required when working with non-default calendars.')
  },
  withErrorHandling(createEventHandler),
  {
    description: 'Create a new calendar event',
    usage: `Use to schedule a new meeting or appointment. Call check_auth_status first to determine if authentication is needed, then list_calendars if you need to work with a specific calendar. For events with attendees, consider using find_meeting_times first to identify suitable time slots.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_calendars to get available calendars
4) Then create_event with parameters:
   {
     "calendarId": "AAMkADE1...", 
     "subject": "Team Meeting",
     "body": "Weekly team sync to discuss project status",
     "bodyType": "Text",
     "start": "2023-11-10T15:00:00Z",
     "end": "2023-11-10T16:00:00Z",
     "location": "Conference Room A",
     "attendees": [
       {
         "email": "colleague@example.com",
         "name": "John Doe",
         "type": "required"
       },
       {
         "email": "manager@example.com",
         "type": "optional"
       }
     ],
     "isOnlineMeeting": true
   }
   
For finding optimal meeting times:
1) Call check_auth_status
2) Call authenticate if needed
3) Call find_meeting_times to check attendee availability
4) Call create_event with the chosen time slot`
  }
);

// Add some folder tools
server.tool(
  "list_folders",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    parentFolderId: z.string().optional().describe('ID of the parent folder to list subfolders from. If not provided, lists top-level folders.')
  },
  withErrorHandling(listFoldersHandler),
  {
    description: 'List mail folders',
    usage: `Use to get available mail folders and their IDs. Call check_auth_status first to determine if authentication is needed. This tool provides folder IDs needed for other operations such as list_emails, move_emails, create_folder, etc. When called without parentFolderId, lists top-level folders; with parentFolderId lists subfolders.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then list_folders to get top-level folders:
   {
     "userId": "default"
   }
   
For listing subfolders:
1) Call check_auth_status
2) Call authenticate if needed
3) Call list_folders to get top-level folders
4) Call list_folders again with parameters:
   {
     "parentFolderId": "AAMkFOL1..."
   }`
  }
);

server.tool(
  "move_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailIds: z.array(z.string()).describe('Array of email IDs to move - obtained from list_emails or search_emails'),
    destinationFolderId: z.string().describe('Destination folder ID - obtained from list_folders or using well-known folder names like "inbox" or "archive"')
  },
  withErrorHandling(moveEmailsHandler),
  {
    description: 'Move emails to a different folder',
    usage: `Use to organize emails by moving them between folders. Call check_auth_status first to determine if authentication is needed, then list_emails or search_emails to get email IDs, and list_folders to get the destination folder ID before using this tool. Can move multiple emails at once for efficiency.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_folders to find or create destination folder
4) Call list_emails to find emails to move
5) Then move_emails with parameters:
   {
     "emailIds": ["AAMkADE1...", "AAMkADE2..."],
     "destinationFolderId": "AAMkFOL1..."
   }
   
For archiving emails:
1) Call check_auth_status
2) Call authenticate if needed
3) Call list_emails to find emails to archive
4) Call move_emails with destinationFolderId="archive"`
  }
);

// Add critical tools based on user needs and add more as required

// Add missing email tools
server.tool(
  "create_draft",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type'),
    to: z.union([z.string(), z.array(z.string())]).optional().describe('Recipient(s)'),
    cc: z.union([z.string(), z.array(z.string())]).optional().describe('CC recipient(s)'),
    bcc: z.union([z.string(), z.array(z.string())]).optional().describe('BCC recipient(s)')
  },
  withErrorHandling(createDraftHandler),
  {
    description: 'Create a draft email',
    usage: `Use to save an email as a draft without sending it. Call check_auth_status first to determine if authentication is needed. This is particularly useful when you need to add attachments before sending.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then create_draft with parameters:
   {
     "subject": "Project Proposal",
     "body": "Please find attached our proposal for the new project.",
     "bodyType": "Text",
     "to": ["recipient1@example.com", "recipient2@example.com"],
     "cc": "manager@example.com"
   }
   
After creating the draft, you can:
1) Add attachments with add_attachment
2) Send the draft using send_email with the draft ID`
  }
);

server.tool(
  "reply_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email to reply to - obtained from list_emails or search_emails'),
    body: z.string().describe('Reply content to add'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type (Text or HTML)'),
    replyAll: z.boolean().optional().describe('Whether to reply to all recipients (true) or just the sender (false)'),
    sendNow: z.boolean().optional().describe('Whether to send immediately (true) or create a draft (false)')
  },
  withErrorHandling(replyEmailHandler),
  {
    description: 'Reply to an email',
    usage: 'Use to respond to an existing email. Call authenticate first, then list_emails or search_emails to find the email, then read_email to view its content before replying. Set replyAll=true to include all original recipients, and sendNow=false to create a draft instead of sending immediately.'
  }
);

server.tool(
  "forward_email",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email to forward'),
    to: z.union([z.string(), z.array(z.string())]).describe('Recipient(s)'),
    cc: z.union([z.string(), z.array(z.string())]).optional().describe('CC recipient(s)'),
    bcc: z.union([z.string(), z.array(z.string())]).optional().describe('BCC recipient(s)'),
    body: z.string().optional().describe('Additional message to include'),
    bodyType: z.enum(['Text', 'HTML']).optional().describe('Body content type')
  },
  withErrorHandling(forwardEmailHandler),
  {
    description: 'Forward an email',
    usage: `Use to forward an existing email to new recipients. Call check_auth_status first to determine if authentication is needed, then list_emails or search_emails to find the email to forward.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_emails or search_emails to find the email
4) Then forward_email with parameters:
   {
     "emailId": "AAMkADE1...",
     "to": ["recipient1@example.com", "recipient2@example.com"],
     "cc": "manager@example.com",
     "body": "Please see the forwarded email below. I'd like your feedback on this.",
     "bodyType": "Text"
   }`
  }
);

server.tool(
  "get_attachment",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email - obtained from list_emails or search_emails'),
    attachmentId: z.string().describe('ID of the attachment - obtained from list_attachments')
  },
  withErrorHandling(getAttachmentHandler),
  {
    description: 'Get email attachment',
    usage: `Use to retrieve a specific attachment from an email. Call check_auth_status first to determine if authentication is needed, then list_emails or search_emails to find the email, then list_attachments to get attachment IDs before using this tool.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_emails or search_emails to find the email with attachments
4) Call list_attachments to get the attachment IDs
5) Then get_attachment with parameters:
   {
     "emailId": "AAMkADE1...",
     "attachmentId": "AAMkATT1..."
   }`
  }
);

server.tool(
  "list_attachments",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email')
  },
  withErrorHandling(listAttachmentsHandler),
  {
    description: 'List attachments for an email',
    usage: `Use to see all attachments on a specific email. Call check_auth_status first to determine if authentication is needed, then list_emails or search_emails to find the email before using this tool.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_emails to find emails with attachments
4) Then list_attachments with parameters:
   {
     "emailId": "AAMkADE1..."
   }
   
This will return a list of attachments with their IDs which can be used with get_attachment to download them.`
  }
);

server.tool(
  "add_attachment",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email or draft - obtained from create_draft or list_emails'),
    name: z.string().describe('Attachment filename including extension (e.g., "document.pdf")'),
    contentType: z.string().describe('MIME type of the attachment (e.g., "application/pdf")'),
    contentBytes: z.string().describe('Base64 encoded content of the attachment')
  },
  withErrorHandling(addAttachmentHandler),
  {
    description: 'Add attachment to an email or draft',
    usage: `Use to attach a file to an email draft before sending. Call check_auth_status first to determine if authentication is needed, then create_draft to create an email draft before adding attachments with this tool.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call create_draft to create a new draft email
4) Then add_attachment with parameters:
   {
     "emailId": "AAMkADE1...",
     "name": "document.pdf",
     "contentType": "application/pdf",
     "contentBytes": "JVBERi0xLjMKJcTl8uXrp/Og0MTGCjQgMCBvYmoKPDwgL0xlbmd0aCA1IDAgUg=="
   }
   
You can add multiple attachments by calling this tool multiple times with the same emailId.`
  }
);

server.tool(
  "delete_attachment",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailId: z.string().describe('ID of the email or draft'),
    attachmentId: z.string().describe('ID of the attachment to delete')
  },
  withErrorHandling(deleteAttachmentHandler),
  {
    description: 'Delete attachment from an email draft',
    usage: `Use to remove an attachment from an email draft. Call check_auth_status first to determine if authentication is needed, then list_attachments to get the attachment IDs.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_attachments to get IDs of attachments on the draft
4) Then delete_attachment with parameters:
   {
     "emailId": "AAMkADE1...",
     "attachmentId": "AAMkATT1..."
   }`
  }
);

// Add missing calendar tools
server.tool(
  "get_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    eventId: z.string().describe('ID of the calendar event')
  },
  withErrorHandling(getEventHandler),
  {
    description: 'Get details of a specific calendar event',
    usage: `Use to retrieve detailed information about a calendar event. Call check_auth_status first to determine if authentication is needed, then list_events to get event IDs.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_events to find the event you're interested in
4) Then get_event with parameters:
   {
     "eventId": "AAMkAEV1..."
   }`
  }
);

server.tool(
  "list_calendars",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")')
  },
  withErrorHandling(listCalendarsHandler),
  {
    description: 'List available calendars',
    usage: `Use to retrieve all calendars the user has access to. Call check_auth_status first to determine if authentication is needed. This tool provides calendar IDs needed for create_event, list_events, etc. when working with non-default calendars.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then list_calendars with parameters:
   {
     "userId": "default"
   }
   
This will return a list of calendars with their IDs which can be used with other calendar tools.`
  }
);

server.tool(
  "update_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    eventId: z.string().describe('ID of the calendar event to update'),
    subject: z.string().optional().describe('New event subject'),
    body: z.string().optional().describe('New event body content'),
    start: z.string().optional().describe('New start time in ISO format'),
    end: z.string().optional().describe('New end time in ISO format'),
    attendees: z.union([z.string(), z.array(z.string())]).optional().describe('New event attendees')
  },
  withErrorHandling(updateEventHandler),
  {
    description: 'Update an existing calendar event',
    usage: `Use to modify the details of an existing meeting or appointment. Call check_auth_status first to determine if authentication is needed, then list_events and get_event to find the event to update.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_events to find the event
4) Call get_event to view its current details
5) Then update_event with parameters:
   {
     "eventId": "AAMkAEV1...",
     "subject": "Updated Meeting Title",
     "start": "2023-11-10T16:00:00Z",
     "end": "2023-11-10T17:00:00Z"
   }
   
Only include the fields you want to change. Omitted fields will keep their current values.`
  }
);

server.tool(
  "respond_to_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    eventId: z.string().describe('ID of the calendar event'),
    response: z.enum(['accept', 'tentativelyAccept', 'decline']).describe('Response to the event invitation'),
    comment: z.string().optional().describe('Optional comment with the response')
  },
  withErrorHandling(respondToEventHandler),
  {
    description: 'Respond to a calendar event invitation',
    usage: `Use to accept, tentatively accept, or decline a meeting invitation. Call check_auth_status first to determine if authentication is needed, then list_events to find invitations.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_events to find invitation events
4) Then respond_to_event with parameters:
   {
     "eventId": "AAMkAEV1...",
     "response": "accept",
     "comment": "Looking forward to the meeting."
   }`
  }
);

server.tool(
  "delete_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    eventId: z.string().describe('ID of the calendar event to delete')
  },
  withErrorHandling(deleteEventHandler),
  {
    description: 'Delete a calendar event',
    usage: `Use to remove an event from the calendar without notifying attendees. Call check_auth_status first to determine if authentication is needed, then list_events to find the event to delete.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_events to find the event
4) Then delete_event with parameters:
   {
     "eventId": "AAMkAEV1..."
   }
   
Note: For meetings you've organized with attendees, consider using cancel_event instead to notify participants.`
  }
);

server.tool(
  "cancel_event",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    eventId: z.string().describe('ID of the calendar event to cancel'),
    comment: z.string().optional().describe('Optional cancellation message')
  },
  withErrorHandling(cancelEventHandler),
  {
    description: 'Cancel a calendar event and notify attendees',
    usage: `Use to cancel a meeting you organized and notify participants. Call check_auth_status first to determine if authentication is needed, then list_events to find the event to cancel.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_events to find the event
4) Then cancel_event with parameters:
   {
     "eventId": "AAMkAEV1...",
     "comment": "This meeting has been cancelled due to a scheduling conflict. We will reschedule soon."
   }`
  }
);

// Add missing folder tools
server.tool(
  "get_folder",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().describe('ID of the folder - obtained from list_folders')
  },
  withErrorHandling(getFolderHandler),
  {
    description: 'Get details of a specific folder',
    usage: `Use to retrieve information about a mail folder. Call check_auth_status first to determine if authentication is needed, then list_folders to find the folder IDs.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_folders to get folder IDs
4) Then get_folder with parameters:
   {
     "folderId": "AAMkFOL1..."
   }`
  }
);

server.tool(
  "create_folder",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    displayName: z.string().describe('Name of the new folder to create'),
    parentFolderId: z.string().optional().describe('ID of the parent folder where the new folder should be created - obtained from list_folders')
  },
  withErrorHandling(createFolderHandler),
  {
    description: 'Create a new mail folder',
    usage: `Use to organize emails by creating new folders. Call check_auth_status first to determine if authentication is needed, then optionally list_folders to get the parentFolderId if you want to create a subfolder.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then create_folder with parameters:
   {
     "displayName": "Project X"
   }
   
For creating a subfolder:
1) Call check_auth_status
2) Call authenticate if needed
3) Call list_folders to get parent folder ID
4) Then create_folder with parameters:
   {
     "displayName": "Meeting Notes",
     "parentFolderId": "AAMkFOL1..."
   }`
  }
);

server.tool(
  "update_folder",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().describe('ID of the folder to update'),
    displayName: z.string().describe('New name for the folder')
  },
  withErrorHandling(updateFolderHandler),
  {
    description: 'Update a mail folder',
    usage: `Use to rename an existing folder. Call check_auth_status first to determine if authentication is needed, then list_folders to get the folder ID.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_folders to find the folder to rename
4) Then update_folder with parameters:
   {
     "folderId": "AAMkFOL1...",
     "displayName": "New Folder Name"
   }`
  }
);

server.tool(
  "delete_folder",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().describe('ID of the folder to delete - obtained from list_folders')
  },
  withErrorHandling(deleteFolderHandler),
  {
    description: 'Delete a mail folder',
    usage: `Use to remove an existing folder. Call check_auth_status first to determine if authentication is needed, then list_folders to identify the folder to delete. Be careful as this permanently removes the folder and all its contents.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_folders to find the folder to delete
4) Then delete_folder with parameters:
   {
     "folderId": "AAMkFOL1..."
   }`
  }
);

server.tool(
  "move_folder",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    folderId: z.string().describe('ID of the folder to move - obtained from list_folders'),
    destinationFolderId: z.string().describe('ID of the destination parent folder - obtained from list_folders')
  },
  withErrorHandling(moveFolderHandler),
  {
    description: 'Move a folder to a new parent folder',
    usage: `Use to reorganize the folder structure. Call check_auth_status first to determine if authentication is needed, then list_folders to get the IDs of both the folder to move and the destination folder.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_folders to get folder IDs
4) Then move_folder with parameters:
   {
     "folderId": "AAMkFOL1...",
     "destinationFolderId": "AAMkFOL2..."
   }`
  }
);

server.tool(
  "copy_emails",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    emailIds: z.array(z.string()).describe('Array of email IDs to copy'),
    destinationFolderId: z.string().describe('Destination folder ID')
  },
  withErrorHandling(copyEmailsHandler),
  {
    description: 'Copy emails to a folder',
    usage: `Use to keep emails in the original location while also placing them in another folder. Call check_auth_status first to determine if authentication is needed, then list_emails to get email IDs and list_folders to get the destination folder ID.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_emails to find emails to copy
4) Call list_folders to find or create the destination folder
5) Then copy_emails with parameters:
   {
     "emailIds": ["AAMkADE1...", "AAMkADE2..."],
     "destinationFolderId": "AAMkFOL1..."
   }`
  }
);

// Add rules tools
server.tool(
  "list_rules",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")')
  },
  withErrorHandling(listRulesHandler),
  {
    description: 'List inbox rules',
    usage: `Use to see all inbox rules configured by the user. Call check_auth_status first to determine if authentication is needed.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then list_rules with parameters:
   {
     "userId": "default"
   }`
  }
);

server.tool(
  "get_rule",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    ruleId: z.string().describe('ID of the rule')
  },
  withErrorHandling(getRuleHandler),
  {
    description: 'Get details of a specific inbox rule',
    usage: `Use to retrieve detailed information about an inbox rule. Call check_auth_status first to determine if authentication is needed, then list_rules to get rule IDs.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_rules to get rule IDs
4) Then get_rule with parameters:
   {
     "ruleId": "RULE123..."
   }`
  }
);

server.tool(
  "create_rule",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    displayName: z.string().describe('Name for the rule'),
    conditions: z.object({}).describe('Conditions that trigger the rule'),
    actions: z.object({}).describe('Actions to take when conditions are met'),
    isEnabled: z.boolean().optional().describe('Whether the rule is enabled')
  },
  withErrorHandling(createRuleHandler),
  {
    description: 'Create a new inbox rule',
    usage: `Use to automate email organization by creating rules. Call check_auth_status first to determine if authentication is needed.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Then create_rule with parameters:
   {
     "displayName": "Move Project X Emails",
     "conditions": {
       "subjectContains": ["Project X"]
     },
     "actions": {
       "moveToFolder": "AAMkFOL1..."
     },
     "isEnabled": true
   }
   
Note: The format of conditions and actions depends on the type of rule you want to create.`
  }
);

server.tool(
  "update_rule",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    ruleId: z.string().describe('ID of the rule to update'),
    displayName: z.string().optional().describe('New name for the rule'),
    conditions: z.object({}).optional().describe('New conditions that trigger the rule'),
    actions: z.object({}).optional().describe('New actions to take when conditions are met'),
    isEnabled: z.boolean().optional().describe('Whether the rule is enabled')
  },
  withErrorHandling(updateRuleHandler),
  {
    description: 'Update an existing inbox rule',
    usage: `Use to modify the conditions or actions of an existing rule. Call check_auth_status first to determine if authentication is needed, then list_rules and get_rule to find the rule to update.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_rules to find the rule
4) Call get_rule to see its current configuration
5) Then update_rule with parameters:
   {
     "ruleId": "RULE123...",
     "displayName": "Updated Rule Name",
     "isEnabled": false
   }
   
Only include the fields you want to change. Omitted fields will keep their current values.`
  }
);

server.tool(
  "delete_rule",
  {
    userId: z.string().optional().describe('User identifier (optional, defaults to "default")'),
    ruleId: z.string().describe('ID of the rule to delete')
  },
  withErrorHandling(deleteRuleHandler),
  {
    description: 'Delete an inbox rule',
    usage: `Use to remove an existing inbox rule. Call check_auth_status first to determine if authentication is needed, then list_rules to find the rule to delete.

Example:
1) Call check_auth_status first
2) Call authenticate if auth_needed is true
3) Call list_rules to find the rule to delete
4) Then delete_rule with parameters:
   {
     "ruleId": "RULE123..."
   }`
  }
);

// Register resources about tool relationships
server.resource('tool-relationships', {
  type: 'json',
  content: JSON.stringify({
    categories: {
      'auth': 'Authentication tools for managing access to Microsoft Graph API',
      'email': 'Email management tools for reading, sending, and organizing messages',
      'folder': 'Folder management tools for organizing mailbox structure',
      'attachment': 'Attachment handling tools for working with email attachments',
      'calendar': 'Calendar management tools for events and appointments',
      'rule': 'Email rule management tools for automatic email processing'
    },
    dependencies: config.toolMetadata,
    commonWorkflows: {
      "email_management": {
        "description": "Basic email workflow",
        "steps": ["check_auth_status", "list_folders", "list_emails", "read_email"],
        "conditional_steps": {
          "authenticate": "Only if check_auth_status indicates auth_needed is true"
        }
      },
      "send_with_attachments": {
        "description": "Send email with attachments",
        "steps": ["check_auth_status", "create_draft", "add_attachment", "send_email"],
        "conditional_steps": {
          "authenticate": "Only if check_auth_status indicates auth_needed is true"
        }
      },
      "calendar_management": {
        "description": "Working with calendars",
        "steps": ["check_auth_status", "list_calendars", "list_events", "get_event"],
        "conditional_steps": {
          "authenticate": "Only if check_auth_status indicates auth_needed is true"
        }
      },
      "schedule_meeting": {
        "description": "Schedule a meeting with attendees",
        "steps": ["check_auth_status", "list_calendars", "find_meeting_times", "create_event"],
        "conditional_steps": {
          "authenticate": "Only if check_auth_status indicates auth_needed is true"
        }
      },
      "organize_emails": {
        "description": "Organize emails into folders",
        "steps": ["check_auth_status", "list_folders", "create_folder", "list_emails", "move_emails"],
        "conditional_steps": {
          "authenticate": "Only if check_auth_status indicates auth_needed is true"
        }
      }
    },
    bestPractices: [
      {
        "title": "Check Authentication Status First",
        "description": "Call check_auth_status before operations to determine if authentication is needed"
      },
      {
        "title": "Authenticate Only When Needed",
        "description": "Only call authenticate when check_auth_status indicates auth_needed is true"
      },
      {
        "title": "Work with Lists Before Items",
        "description": "Call list_* tools first to get IDs before working with specific items"
      },
      {
        "title": "Check Dependencies",
        "description": "Review the dependencies field for each tool to ensure prerequisites are met"
      },
      {
        "title": "Use Well-Known Folders",
        "description": "For common folders, use well-known names like 'inbox', 'drafts', 'sentitems' instead of IDs"
      },
      {
        "title": "Find Meeting Times Before Creating Events",
        "description": "Use find_meeting_times to check availability before scheduling events with attendees"
      }
    ],
    authenticationNotes: {
      "tokenCaching": "Authentication tokens are cached locally and remain valid until they expire",
      "checkBeforeAuth": "Always call check_auth_status first to determine if authentication is needed",
      "whenToAuthenticate": [
        "When starting a new session",
        "When check_auth_status indicates auth_needed is true",
        "When a tool returns an authentication error",
        "When tokens are about to expire"
      ]
    }
  })
});

// Add a tool helper for finding related tools
server.tool(
  "get_tool_info",
  {
    toolName: z.string().describe('Name of the tool to get information about'),
    includeRelated: z.boolean().optional().describe('Whether to include related tools in the response')
  },
  async (params) => {
    const { toolName, includeRelated = true } = params;
    
    if (!config.toolMetadata[toolName]) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'tool_not_found',
              message: `Tool "${toolName}" not found. Use one of the available tools.`,
              available_tools: Object.keys(config.toolMetadata)
            }
          })
        }]
      };
    }
    
    const toolInfo = config.toolMetadata[toolName];
    const result = {
      tool: toolName,
      category: toolInfo.category,
      dependencies: toolInfo.dependencies || []
    };
    
    // Add conditional authentication information for tools that depend on authenticate
    if (toolInfo.dependencies.includes('authenticate') && toolName !== 'authenticate') {
      result.authentication_note = "Authentication is required but may already be cached. Call check_auth_status first to determine if authenticate needs to be called.";
      
      // Replace authenticate with check_auth_status in the dependencies list for display
      const dependencies = [...result.dependencies];
      const authIndex = dependencies.indexOf('authenticate');
      if (authIndex !== -1) {
        dependencies[authIndex] = 'check_auth_status';
        result.dependency_workflow = ['check_auth_status', 'authenticate (if needed)', ...dependencies.filter(d => d !== 'check_auth_status')];
      }
    }
    
    if (includeRelated && toolInfo.related) {
      result.related_tools = toolInfo.related.map(relatedTool => {
        const relatedInfo = config.toolMetadata[relatedTool];
        return {
          name: relatedTool,
          category: relatedInfo?.category || 'unknown',
          recommended_sequence: toolInfo.dependencies.includes(relatedTool) ? 
            'before' : 
            (relatedInfo?.dependencies.includes(toolName) ? 'after' : 'any')
        };
      });
      
      // Find common workflows that include this tool
      const relevantWorkflows = [];
      const commonWorkflows = JSON.parse(server.resources['tool-relationships'].content).commonWorkflows;
      
      for (const [name, workflow] of Object.entries(commonWorkflows)) {
        if (workflow.steps.includes(toolName)) {
          relevantWorkflows.push({
            name,
            description: workflow.description,
            position: workflow.steps.indexOf(toolName) + 1,
            total_steps: workflow.steps.length,
            steps: workflow.steps,
            conditional_steps: workflow.conditional_steps || {}
          });
        }
      }
      
      if (relevantWorkflows.length > 0) {
        result.workflows = relevantWorkflows;
      }
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify(result)
      }]
    };
  },
  {
    description: 'Get information about a specific tool and its relationships',
    usage: `Use to understand how tools relate to each other and which tools should be called first. 
    
Example:
1) Call get_tool_info with parameters:
   {
     "toolName": "create_event",
     "includeRelated": true
   }

This will return information about the create_event tool, its dependencies, related tools, and workflows it's part of.

Important: For tools that have authenticate as a dependency, check_auth_status should be called first to determine if authentication is actually needed.`
  }
);

// Helper function for registering tool sequences
function registerToolSequence(name, toolSequence, metadata) {
  server.resource(`sequences/${name}`, {
    type: 'json',
    content: JSON.stringify({
      sequence: toolSequence,
      ...metadata
    })
  });
  
  logger.info(`Registered tool sequence: ${name}`);
}

// Register common tool sequences
// Email management sequences
registerToolSequence(
  "view_emails",
  ["check_auth_status", "authenticate", "list_folders", "list_emails", "read_email"],
  {
    description: "Complete workflow for browsing and reading emails",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Find and read recent emails from my inbox",
    parameters: {
      "list_emails": {
        "folderId": "inbox",
        "limit": 10
      }
    }
  }
);

registerToolSequence(
  "send_simple_email",
  ["check_auth_status", "authenticate", "send_email"],
  {
    description: "Workflow for sending a basic email",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Send an email to john@example.com with the subject 'Meeting Tomorrow'",
    parameters: {
      "send_email": {
        "to": "recipient@example.com",
        "subject": "Example Subject",
        "body": "Email content goes here",
        "bodyType": "Text"
      }
    }
  }
);

registerToolSequence(
  "send_email_with_attachment",
  ["check_auth_status", "authenticate", "create_draft", "add_attachment", "send_email"],
  {
    description: "Workflow for sending an email with attachments",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Send an email with the project proposal PDF attached",
    parameters: {
      "create_draft": {
        "to": "recipient@example.com",
        "subject": "Proposal Document",
        "body": "Please find the proposal attached."
      },
      "add_attachment": {
        "name": "proposal.pdf",
        "contentType": "application/pdf"
      }
    }
  }
);

registerToolSequence(
  "reply_to_email",
  ["check_auth_status", "authenticate", "list_emails", "read_email", "reply_email"],
  {
    description: "Workflow for replying to an existing email",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Reply to the email from Sarah about the project timeline",
    parameters: {
      "reply_email": {
        "replyAll": false,
        "body": "Here's my response..."
      }
    }
  }
);

// Calendar management sequences
registerToolSequence(
  "view_calendar_events",
  ["check_auth_status", "authenticate", "list_calendars", "list_events"],
  {
    description: "Workflow for viewing upcoming calendar events",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Show my calendar events for next week",
    parameters: {
      "list_events": {
        "startDateTime": "2023-11-01T00:00:00Z",
        "endDateTime": "2023-11-08T00:00:00Z"
      }
    }
  }
);

registerToolSequence(
  "schedule_meeting",
  ["check_auth_status", "authenticate", "list_calendars", "find_meeting_times", "create_event"],
  {
    description: "Complete workflow for scheduling a meeting with attendees",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Schedule a team meeting next Tuesday at 2pm",
    parameters: {
      "find_meeting_times": {
        "attendees": [
          { "email": "colleague1@example.com", "type": "required" },
          { "email": "colleague2@example.com", "type": "required" }
        ],
        "durationInMinutes": 60
      },
      "create_event": {
        "subject": "Team Meeting",
        "isOnlineMeeting": true
      }
    }
  }
);

registerToolSequence(
  "manage_event_response",
  ["check_auth_status", "authenticate", "list_events", "respond_to_event"],
  {
    description: "Workflow for responding to event invitations",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Accept the meeting invitation from the marketing team",
    parameters: {
      "respond_to_event": {
        "response": "accept"
      }
    }
  }
);

// Email organization sequences
registerToolSequence(
  "organize_emails",
  ["check_auth_status", "authenticate", "list_folders", "create_folder", "list_emails", "move_emails"],
  {
    description: "Workflow for organizing emails into folders",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Move all emails from John to a new Project X folder",
    parameters: {
      "create_folder": {
        "displayName": "Project X"
      },
      "list_emails": {
        "search": "from:john@example.com"
      }
    }
  }
);

registerToolSequence(
  "setup_email_rule",
  ["check_auth_status", "authenticate", "list_folders", "create_rule"],
  {
    description: "Workflow for setting up automatic email organization rules",
    conditional_steps: {
      "authenticate": "Only if check_auth_status indicates auth_needed is true"
    },
    example: "Create a rule to move all emails with 'Invoice' in the subject to the Finance folder",
    parameters: {
      "create_rule": {
        "displayName": "Move Invoice Emails",
        "conditions": {
          "subjectContains": ["Invoice"]
        },
        "actions": {
          "moveToFolder": "Finance Folder ID"
        }
      }
    }
  }
);

// Register a sequences resource that lists all available sequences
server.resource('available-sequences', {
  type: 'json',
  content: JSON.stringify({
    email: ["view_emails", "send_simple_email", "send_email_with_attachment", "reply_to_email"],
    calendar: ["view_calendar_events", "schedule_meeting", "manage_event_response"],
    organization: ["organize_emails", "setup_email_rule"]
  })
});

// Add a tool for getting sequence information
server.tool(
  "get_sequence",
  {
    sequenceName: z.string().optional().describe('Name of the sequence to get details for'),
    category: z.string().optional().describe('Category of sequences to list (email, calendar, organization)')
  },
  async (params) => {
    const { sequenceName, category } = params;
    
    // If a specific sequence is requested, return its details
    if (sequenceName) {
      try {
        const resourceName = `sequences/${sequenceName}`;
        if (!server.resources[resourceName]) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: {
                  type: 'sequence_not_found',
                  message: `Sequence "${sequenceName}" not found`,
                  available_sequences: Object.keys(server.resources)
                    .filter(key => key.startsWith('sequences/'))
                    .map(key => key.replace('sequences/', ''))
                }
              })
            }]
          };
        }
        
        const sequenceData = JSON.parse(server.resources[resourceName].content);
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              name: sequenceName,
              ...sequenceData
            })
          }]
        };
      } catch (error) {
        logger.error(`Error retrieving sequence ${sequenceName}:`, error);
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: {
                type: 'internal_error',
                message: `Error retrieving sequence: ${error.message}`
              }
            })
          }]
        };
      }
    }
    
    // If a category is provided, list sequences in that category
    if (category) {
      try {
        const availableSequences = JSON.parse(server.resources['available-sequences'].content);
        
        if (!availableSequences[category]) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                error: {
                  type: 'category_not_found',
                  message: `Category "${category}" not found`,
                  available_categories: Object.keys(availableSequences)
                }
              })
            }]
          };
        }
        
        const sequences = availableSequences[category];
        const sequenceDetails = sequences.map(name => {
          const resourceName = `sequences/${name}`;
          const data = JSON.parse(server.resources[resourceName].content);
          return {
            name,
            description: data.description,
            example: data.example
          };
        });
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              category,
              sequences: sequenceDetails
            })
          }]
        };
      } catch (error) {
        logger.error(`Error retrieving sequences for category ${category}:`, error);
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              error: {
                type: 'internal_error',
                message: `Error retrieving sequences: ${error.message}`
              }
            })
          }]
        };
      }
    }
    
    // If neither sequence name nor category provided, list all available categories
    try {
      const availableSequences = JSON.parse(server.resources['available-sequences'].content);
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            available_categories: Object.keys(availableSequences),
            message: "Provide a category or sequenceName to get more details"
          })
        }]
      };
    } catch (error) {
      logger.error('Error retrieving available sequences:', error);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'internal_error',
              message: `Error retrieving available sequences: ${error.message}`
            }
          })
        }]
      };
    }
  },
  {
    description: 'Get information about available tool sequences for common workflows',
    usage: `Use to discover predefined sequences of tools for common tasks. This helps understand the recommended order of tool calls for different workflows.

Example for listing all categories:
1) Call get_sequence with no parameters:
   {}

Example for listing sequences in a category:
1) Call get_sequence with category parameter:
   {
     "category": "email"
   }

Example for getting details of a specific sequence:
1) Call get_sequence with sequenceName parameter:
   {
     "sequenceName": "send_email_with_attachment"
   }

This will return the full sequence of tools to call, along with example parameters for each step.`
  }
);

// Add a tool for suggesting the appropriate workflow for a task
server.tool(
  "suggest_workflow",
  {
    task: z.string().describe('Description of the task you want to accomplish')
  },
  async (params) => {
    const { task } = params;
    
    if (!task) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'missing_parameter',
              message: 'Please provide a task description'
            }
          })
        }]
      };
    }
    
    try {
      // Get all available sequences
      const availableSequences = JSON.parse(server.resources['available-sequences'].content);
      const allSequences = Object.values(availableSequences).flat();
      
      // Create an array of sequence details for matching
      const sequenceDetails = allSequences.map(name => {
        const resourceName = `sequences/${name}`;
        const data = JSON.parse(server.resources[resourceName].content);
        return {
          name,
          description: data.description,
          example: data.example,
          keywords: `${data.description} ${data.example}`.toLowerCase(),
          sequence: data.sequence,
          parameters: data.parameters || {}
        };
      });
      
      // Simple keyword matching to find relevant sequences
      const taskLower = task.toLowerCase();
      const matchedSequences = sequenceDetails
        .map(seq => ({
          ...seq,
          score: calculateRelevanceScore(taskLower, seq.keywords)
        }))
        .filter(seq => seq.score > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, 3); // Return top 3 matches
      
      if (matchedSequences.length === 0) {
        // If no direct matches, suggest based on task type
        let suggestedSequence = null;
        
        if (taskLower.includes('email') && taskLower.includes('read')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'view_emails');
        } else if (taskLower.includes('email') && taskLower.includes('send')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'send_simple_email');
        } else if (taskLower.includes('attachment')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'send_email_with_attachment');
        } else if (taskLower.includes('calendar') || taskLower.includes('event')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'view_calendar_events');
        } else if (taskLower.includes('meeting') || taskLower.includes('schedule')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'schedule_meeting');
        } else if (taskLower.includes('folder') || taskLower.includes('organize')) {
          suggestedSequence = sequenceDetails.find(s => s.name === 'organize_emails');
        }
        
        if (suggestedSequence) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                suggested_sequence: suggestedSequence.name,
                description: suggestedSequence.description,
                tools: suggestedSequence.sequence,
                parameters: suggestedSequence.parameters,
                suggestion_confidence: "medium",
                message: "No exact match found. This is a suggested workflow based on your task description."
              })
            }]
          };
        }
        
        // No matches found
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              message: "No suitable workflow found for this task. Please use get_sequence to explore available workflows or break down your task into simpler steps.",
              available_categories: Object.keys(availableSequences)
            })
          }]
        };
      }
      
      // Return the best match with full details
      const bestMatch = matchedSequences[0];
      
      // Get the complete sequence details
      const sequenceData = JSON.parse(server.resources[`sequences/${bestMatch.name}`].content);
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            suggested_sequence: bestMatch.name,
            description: bestMatch.description,
            tools: sequenceData.sequence,
            conditional_steps: sequenceData.conditional_steps || {},
            parameters: sequenceData.parameters || {},
            suggestion_confidence: bestMatch.score > 5 ? "high" : "medium",
            alternatives: matchedSequences.slice(1).map(seq => ({
              name: seq.name,
              description: seq.description
            }))
          })
        }]
      };
    } catch (error) {
      logger.error('Error suggesting workflow:', error);
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: {
              type: 'internal_error',
              message: `Error suggesting workflow: ${error.message}`
            }
          })
        }]
      };
    }
  },
  {
    description: 'Suggest the appropriate workflow for a given task',
    usage: `Use this tool to find the recommended sequence of tools for a specific task. Provide a clear description of what you want to accomplish.

Example:
1) Call suggest_workflow with a task description:
   {
     "task": "I want to send an email with a PDF attachment"
   }

This will return a suggested workflow with the sequence of tools to call and example parameters for each step.`
  }
);

// Helper function to calculate relevance score for workflow suggestions
function calculateRelevanceScore(task, keywords) {
  let score = 0;
  
  // Split task into words for matching
  const taskWords = task.split(/\s+/);
  
  // Check for exact phrase matches
  const keyPhrases = [
    { phrase: "send email", score: 5 },
    { phrase: "read email", score: 5 },
    { phrase: "email with attachment", score: 7 },
    { phrase: "schedule meeting", score: 6 },
    { phrase: "organize emails", score: 5 },
    { phrase: "create folder", score: 4 },
    { phrase: "move emails", score: 4 },
    { phrase: "view calendar", score: 5 },
    { phrase: "respond to invitation", score: 6 },
    { phrase: "check events", score: 4 }
  ];
  
  keyPhrases.forEach(({ phrase, score: phraseScore }) => {
    if (task.includes(phrase)) {
      score += phraseScore;
    }
  });
  
  // Check for individual keyword matches
  const keywordScores = {
    "email": 2,
    "emails": 2,
    "send": 2,
    "read": 2,
    "calendar": 2,
    "event": 2,
    "events": 2,
    "meeting": 2,
    "schedule": 2,
    "attachment": 3,
    "attachments": 3,
    "pdf": 1,
    "document": 1,
    "folder": 2,
    "organize": 2,
    "move": 1,
    "create": 1,
    "rule": 2,
    "invitation": 2,
    "reply": 2,
    "forward": 2
  };
  
  taskWords.forEach(word => {
    if (keywordScores[word]) {
      score += keywordScores[word];
    }
  });
  
  // Check if keywords contain task words for partial matches
  taskWords.forEach(word => {
    if (word.length > 3 && keywords.includes(word)) {
      score += 1;
    }
  });
  
  return score;
}

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