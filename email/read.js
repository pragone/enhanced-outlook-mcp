const config = require('../config');
const logger = require('../utils/logger');
const { createGraphClient } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');

/**
 * Read a specific email by ID
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Email content
 */
async function readEmailHandler(params = {}) {
  try {
    // Super detailed debugging
    logger.info(`READ EMAIL HANDLER START ------------------------------------`);
    logger.info(`Raw params: ${JSON.stringify(params)}`);
    
    // CRITICAL FIX FOR CLAUDE DESKTOP:
    // Check if this is a direct JSON-RPC request with the tool call format
    if (global.__last_message?.method === 'tools/call' && 
        global.__last_message?.params?.name === 'read_email' &&
        global.__last_message?.params?.arguments) {
      
      const directArgs = global.__last_message.params.arguments;
      logger.info(`Found direct JSON-RPC arguments: ${JSON.stringify(directArgs)}`);
      
      // Use the arguments directly from the JSON-RPC request
      if (directArgs.id || directArgs.messageId || directArgs.message_id || directArgs.emailId) {
        params = directArgs;
        logger.info(`Using direct JSON-RPC arguments for read_email`);
      }
    }
    
    // Check for raw message in params
    const rawMessage = params.__raw_message;
    if (rawMessage) {
      logger.info(`Found raw message in params: ${JSON.stringify(rawMessage)}`);
    }
    
    // Try all possible sources of parameters
    let requestParams = {};
    
    // Order of preference for parameter sources:
    // 1. Direct params
    // 2. params.arguments
    // 3. Raw message params.arguments
    // 4. params.contextData (Claude Desktop might use this)
    // 5. Check for potential email ID saved by enhanced transport
    // 6. Global last message
    
    if (params.message_id || params.messageId || params.id || params.emailId) {
      // Use direct params
      requestParams = params;
      logger.info(`Using direct params`);
    } else if (params.arguments) {
      // Use params.arguments
      requestParams = typeof params.arguments === 'object' ? params.arguments : params;
      logger.info(`Using params.arguments`);
      
      // Check if arguments is a string containing JSON
      if (typeof params.arguments === 'string') {
        try {
          const parsedArgs = JSON.parse(params.arguments);
          if (parsedArgs && typeof parsedArgs === 'object') {
            requestParams = parsedArgs;
            logger.info(`Parsed string arguments into object: ${JSON.stringify(parsedArgs)}`);
          }
        } catch (e) {
          logger.info(`Arguments string is not valid JSON: ${params.arguments}`);
        }
      }
    } else if (rawMessage?.params?.arguments) {
      // Use raw message params
      requestParams = rawMessage.params.arguments;
      logger.info(`Using raw message params`);
    } else if (params.contextData) {
      // Try to use contextData (sometimes used by Claude Desktop)
      requestParams = params.contextData;
      logger.info(`Using params.contextData: ${JSON.stringify(params.contextData)}`);
    } else if (params.signal) {
      // Handle Claude Desktop format - try to extract parameters from request
      logger.info(`Detected Claude Desktop format with signal property`);
      
      // Check if params has any properties that might be the email ID
      // Loop through all properties in case the ID is stored in an unconventional property
      for (const key in params) {
        // Skip known non-ID properties
        if (['signal', 'tool'].includes(key)) continue;
        
        const value = params[key];
        if (typeof value === 'string' && value.length > 20) {
          // This could be an email ID
          logger.info(`Found potential email ID in property ${key}: ${value}`);
          requestParams.id = value;
          break;
        }
        
        // Check if value is an object that might contain the email ID
        if (typeof value === 'object' && value !== null) {
          for (const subKey in value) {
            if (['id', 'messageId', 'message_id', 'emailId'].includes(subKey)) {
              logger.info(`Found email ID in nested property ${key}.${subKey}: ${value[subKey]}`);
              requestParams.id = value[subKey];
              break;
            }
          }
        }
      }
      
      // Check for any ID stored by the enhanced transport in global vars
      if (!requestParams.id && global._claude_potential_email_id) {
        logger.info(`Using email ID stored by enhanced transport: ${global._claude_potential_email_id}`);
        requestParams.id = global._claude_potential_email_id;
      }
      
      // If all else fails, try to use last stored email ID
      if (!requestParams.id && global._claude_last_email_id) {
        logger.info(`Using stored email ID from global context: ${global._claude_last_email_id}`);
        requestParams.id = global._claude_last_email_id;
      }
    } else if (global.__last_message?.params?.arguments) {
      // Use global last message
      requestParams = global.__last_message.params.arguments;
      logger.info(`Using global last message params`);
    } else {
      // Last resort - check globals directly
      if (global._claude_potential_email_id) {
        requestParams.id = global._claude_potential_email_id;
        logger.info(`Using _claude_potential_email_id from global context: ${global._claude_potential_email_id}`);
      } else if (global._claude_last_email_id) {
        requestParams.id = global._claude_last_email_id;
        logger.info(`Using _claude_last_email_id from global context: ${global._claude_last_email_id}`);
      }
    }
    
    logger.info(`Request params extracted: ${JSON.stringify(requestParams)}`);
    
    // Extract userId
    let userId = requestParams.userId;
    logger.info(`Extracted userId: ${userId}`);
    
    // Extract message ID from various parameter names
    const message_id = requestParams.message_id;
    const messageId = requestParams.messageId;
    const id = requestParams.id;
    const emailId = requestParams.emailId;
    
    logger.info(`Found parameter values:
      - message_id: ${message_id}
      - messageId: ${messageId}
      - id: ${id}
      - emailId: ${emailId}
    `);
    
    // Try to extract the message ID
    let finalMessageId = message_id || messageId || id || emailId;
    logger.info(`Final extracted messageId: ${finalMessageId}`);
    
    // Store email ID in global context for potential future use
    if (finalMessageId) {
      global._claude_last_email_id = finalMessageId;
    }
    
    const markAsRead = requestParams.markAsRead || false;
    logger.info(`markAsRead: ${markAsRead}`);
    
    if (!userId) {
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text", 
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Please authenticate first.'
            })
          }]
        };
      }
      userId = users.length === 1 ? users[0] : requestParams.userId;
      if (!userId) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'Multiple users found. Please specify userId parameter.'
            })
          }]
        };
      }
    }
    
    if (!finalMessageId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required. Please provide either "id", "messageId", or "emailId" parameter.'
          })
        }]
      };
    }
    
    logger.info(`Reading email ${finalMessageId} for user ${userId}`);
    
    const graphClient = await createGraphClient(userId);
    
    // Get email with detailed content
    const email = await graphClient.get(`/me/messages/${finalMessageId}`, {
      $expand: 'attachments'
    });
    
    if (!email) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: `Email not found with ID: ${finalMessageId}`
          })
        }]
      };
    }
    
    // Mark as read if requested
    if (markAsRead && !email.isRead) {
      await graphClient.patch(`/me/messages/${finalMessageId}`, {
        isRead: true
      });
    }
    
    // Format the detailed email response
    const formattedEmail = formatDetailedEmail(email);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          email: formattedEmail
        })
      }]
    };
  } catch (error) {
    logger.error(`Error reading email: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to read email: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Format detailed email response
 * @param {Object} email - Raw email from Graph API
 * @returns {Object} - Formatted detailed email
 */
function formatDetailedEmail(email) {
  // Extract sender information
  let sender = null;
  if (email.from && email.from.emailAddress) {
    sender = {
      name: email.from.emailAddress.name,
      email: email.from.emailAddress.address
    };
  }
  
  // Extract recipients
  let toRecipients = [];
  if (email.toRecipients && Array.isArray(email.toRecipients)) {
    toRecipients = email.toRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Extract CC recipients
  let ccRecipients = [];
  if (email.ccRecipients && Array.isArray(email.ccRecipients)) {
    ccRecipients = email.ccRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Extract BCC recipients
  let bccRecipients = [];
  if (email.bccRecipients && Array.isArray(email.bccRecipients)) {
    bccRecipients = email.bccRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Format attachments
  let attachments = [];
  if (email.attachments && Array.isArray(email.attachments)) {
    attachments = email.attachments.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId
    }));
  }
  
  // Format body content
  let body = {
    contentType: email.body ? email.body.contentType : 'text',
    content: email.body ? email.body.content : ''
  };
  
  // Create formatted response
  return {
    id: email.id,
    subject: email.subject || '(No Subject)',
    sender,
    toRecipients,
    ccRecipients,
    bccRecipients,
    receivedDateTime: email.receivedDateTime,
    sentDateTime: email.sentDateTime,
    hasAttachments: !!email.hasAttachments,
    attachments,
    isRead: email.isRead,
    isDraft: email.isDraft,
    importance: email.importance,
    body,
    conversationId: email.conversationId,
    parentFolderId: email.parentFolderId,
    internetMessageId: email.internetMessageId,
    webLink: email.webLink,
    categories: email.categories || []
  };
}

/**
 * Mark an email as read or unread
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Operation result
 */
async function markEmailHandler(params = {}) {
  try {
    // Super detailed debugging
    logger.info(`MARK EMAIL HANDLER START ------------------------------------`);
    logger.info(`Raw params: ${JSON.stringify(params)}`);
    
    // CRITICAL FIX FOR CLAUDE DESKTOP:
    // Check if this is a direct JSON-RPC request with the tool call format
    if (global.__last_message?.method === 'tools/call' && 
        global.__last_message?.params?.name === 'mark_email' &&
        global.__last_message?.params?.arguments) {
      
      const directArgs = global.__last_message.params.arguments;
      logger.info(`Found direct JSON-RPC arguments: ${JSON.stringify(directArgs)}`);
      
      // Use the arguments directly from the JSON-RPC request
      if (directArgs.id || directArgs.messageId || directArgs.message_id || directArgs.emailId) {
        params = directArgs;
        logger.info(`Using direct JSON-RPC arguments for mark_email`);
      }
    }
    
    // Check for raw message in params
    const rawMessage = params.__raw_message;
    if (rawMessage) {
      logger.info(`Found raw message in params: ${JSON.stringify(rawMessage)}`);
    }
    
    // Try all possible sources of parameters
    let requestParams = {};
    
    // Order of preference for parameter sources:
    // 1. Direct params
    // 2. params.arguments
    // 3. Raw message params.arguments
    // 4. params.contextData (Claude Desktop might use this)
    // 5. Check for potential email ID saved by enhanced transport
    // 6. Global last message
    
    if (params.message_id || params.messageId || params.id || params.emailId) {
      // Use direct params
      requestParams = params;
      logger.info(`Using direct params`);
    } else if (params.arguments) {
      // Use params.arguments
      requestParams = typeof params.arguments === 'object' ? params.arguments : params;
      logger.info(`Using params.arguments`);
      
      // Check if arguments is a string containing JSON
      if (typeof params.arguments === 'string') {
        try {
          const parsedArgs = JSON.parse(params.arguments);
          if (parsedArgs && typeof parsedArgs === 'object') {
            requestParams = parsedArgs;
            logger.info(`Parsed string arguments into object: ${JSON.stringify(parsedArgs)}`);
          }
        } catch (e) {
          logger.info(`Arguments string is not valid JSON: ${params.arguments}`);
        }
      }
    } else if (rawMessage?.params?.arguments) {
      // Use raw message params
      requestParams = rawMessage.params.arguments;
      logger.info(`Using raw message params`);
    } else if (params.contextData) {
      // Try to use contextData (sometimes used by Claude Desktop)
      requestParams = params.contextData;
      logger.info(`Using params.contextData: ${JSON.stringify(params.contextData)}`);
    } else if (params.signal) {
      // Handle Claude Desktop format - try to extract parameters from request
      logger.info(`Detected Claude Desktop format with signal property`);
      
      // Check if params has any properties that might be the email ID
      // Loop through all properties in case the ID is stored in an unconventional property
      for (const key in params) {
        // Skip known non-ID properties
        if (['signal', 'tool'].includes(key)) continue;
        
        const value = params[key];
        if (typeof value === 'string' && value.length > 20) {
          // This could be an email ID
          logger.info(`Found potential email ID in property ${key}: ${value}`);
          requestParams.id = value;
          break;
        }
        
        // Check if value is an object that might contain the email ID
        if (typeof value === 'object' && value !== null) {
          for (const subKey in value) {
            if (['id', 'messageId', 'message_id', 'emailId'].includes(subKey)) {
              logger.info(`Found email ID in nested property ${key}.${subKey}: ${value[subKey]}`);
              requestParams.id = value[subKey];
              break;
            }
          }
        }
      }
      
      // Check for any ID stored by the enhanced transport in global vars
      if (!requestParams.id && global._claude_potential_email_id) {
        logger.info(`Using email ID stored by enhanced transport: ${global._claude_potential_email_id}`);
        requestParams.id = global._claude_potential_email_id;
      }
      
      // If all else fails, try to use last stored email ID
      if (!requestParams.id && global._claude_last_email_id) {
        logger.info(`Using stored email ID from global context: ${global._claude_last_email_id}`);
        requestParams.id = global._claude_last_email_id;
      }
    } else if (global.__last_message?.params?.arguments) {
      // Use global last message
      requestParams = global.__last_message.params.arguments;
      logger.info(`Using global last message params`);
    } else {
      // Last resort - check globals directly
      if (global._claude_potential_email_id) {
        requestParams.id = global._claude_potential_email_id;
        logger.info(`Using _claude_potential_email_id from global context: ${global._claude_potential_email_id}`);
      } else if (global._claude_last_email_id) {
        requestParams.id = global._claude_last_email_id;
        logger.info(`Using _claude_last_email_id from global context: ${global._claude_last_email_id}`);
      }
    }
    
    logger.info(`Request params extracted: ${JSON.stringify(requestParams)}`);
    
    // Extract userId
    let userId = requestParams.userId;
    logger.info(`Extracted userId: ${userId}`);
    
    // Extract message ID from various parameter names
    const message_id = requestParams.message_id;
    const messageId = requestParams.messageId;
    const id = requestParams.id;
    const emailId = requestParams.emailId;
    
    logger.info(`Found parameter values:
      - message_id: ${message_id}
      - messageId: ${messageId}
      - id: ${id}
      - emailId: ${emailId}
    `);
    
    // Try to extract the message ID
    let finalEmailId = message_id || messageId || id || emailId;
    logger.info(`Final extracted emailId: ${finalEmailId}`);
    
    // Store email ID in global context for potential future use
    if (finalEmailId) {
      global._claude_last_email_id = finalEmailId;
    }
    
    const isRead = requestParams.isRead === undefined ? true : !!requestParams.isRead;
    logger.info(`isRead: ${isRead}`);
    
    if (!userId) {
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Please authenticate first.'
            })
          }]
        };
      }
      userId = users.length === 1 ? users[0] : requestParams.userId;
      if (!userId) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'Multiple users found. Please specify userId parameter.'
            })
          }]
        };
      }
    }
    
    if (!finalEmailId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required. Please provide either "id", "messageId", or "emailId" parameter.'
          })
        }]
      };
    }
    
    logger.info(`Marking email ${finalEmailId} as ${isRead ? 'read' : 'unread'} for user ${userId}`);
    
    const graphClient = await createGraphClient(userId);
    
    // Update the read status
    await graphClient.patch(`/me/messages/${finalEmailId}`, {
      isRead
    });
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: `Email marked as ${isRead ? 'read' : 'unread'} successfully`,
          emailId: finalEmailId,
          isRead
        })
      }]
    };
  } catch (error) {
    logger.error(`Error marking email: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to mark email: ${error.message}`
        })
      }]
    };
  }
}

/**
 * MCP Schema definition for tools
 * This helps Claude understand the expected parameters for each tool
 */
const mcpSchema = {
  tools: [
    {
      name: "read_email",
      description: "Reads an email message by its ID. When sending from Claude Desktop, make sure to include the email ID in the request parameters.",
      parameters: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description: "The ID of the email message to read (one of id, messageId, emailId, or message_id is required)"
          },
          messageId: {
            type: "string",
            description: "Alternative parameter name for email ID (same as id)"
          },
          message_id: {
            type: "string",
            description: "Snake case alternative parameter name for email ID (same as id)"
          },
          emailId: {
            type: "string",
            description: "Alternative parameter name for email ID (same as id)"
          },
          userId: {
            type: "string",
            description: "Optional user ID to specify which account to use"
          },
          markAsRead: {
            type: "boolean",
            description: "Whether to mark the email as read"
          }
        },
        required: []
      }
    },
    {
      name: "mark_email",
      description: "Marks an email as read or unread. When sending from Claude Desktop, make sure to include the email ID in the request parameters.",
      parameters: {
        type: "object",
        properties: {
          id: {
            type: "string",
            description: "The ID of the email message to mark (one of id, messageId, emailId, or message_id is required)"
          },
          messageId: {
            type: "string",
            description: "Alternative parameter name for email ID (same as id)"
          },
          message_id: {
            type: "string",
            description: "Snake case alternative parameter name for email ID (same as id)"
          },
          emailId: {
            type: "string",
            description: "Alternative parameter name for email ID (same as id)"
          },
          userId: {
            type: "string",
            description: "Optional user ID to specify which account to use"
          },
          isRead: {
            type: "boolean",
            description: "Whether to mark as read (true) or unread (false). Default is true."
          }
        },
        required: []
      }
    }
  ]
};

module.exports = {
  readEmailHandler,
  markEmailHandler,
  mcpSchema
};