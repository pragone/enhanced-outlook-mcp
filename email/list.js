const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');
const { listUsers } = require('../auth/token-manager');
const { buildQueryParams } = require('../utils/odata-helpers');

/**
 * List emails from a mailbox
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of emails
 */
async function listEmailsHandler(params = {}) {
  try {
    // Super detailed debugging
    logger.info(`LIST EMAILS HANDLER START ------------------------------------`);
    logger.info(`Raw params: ${JSON.stringify(params)}`);
    
    // CRITICAL FIX FOR CLAUDE DESKTOP:
    // Check if this is a direct JSON-RPC request with the tool call format
    if (global.__last_message?.method === 'tools/call' && 
        global.__last_message?.params?.name === 'list_emails' &&
        global.__last_message?.params?.arguments) {
      
      const directArgs = global.__last_message.params.arguments;
      logger.info(`Found direct JSON-RPC arguments: ${JSON.stringify(directArgs)}`);
      
      // Use the arguments directly from the JSON-RPC request
      params = directArgs;
      logger.info(`Using direct JSON-RPC arguments for list_emails`);
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
    // 5. Global last message
    
    if (Object.keys(params).length > 1 || (params.folderId || params.userId)) {
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
      
      // Check if params has any properties that might be folder ID
      for (const key in params) {
        // Skip known non-parameters
        if (['signal', 'tool'].includes(key)) continue;
        
        const value = params[key];
        if (typeof value === 'string') {
          // This could be a folder ID or name
          if (['inbox', 'drafts', 'sentitems', 'deleteditems'].includes(value.toLowerCase())) {
            logger.info(`Found folder name in property ${key}: ${value}`);
            requestParams.folderId = value;
            break;
          }
        }
        
        // Check if value is an object that might contain parameters
        if (typeof value === 'object' && value !== null) {
          for (const subKey in value) {
            if (['folderId', 'folder', 'folderName'].includes(subKey)) {
              logger.info(`Found folder ID in nested property ${key}.${subKey}: ${value[subKey]}`);
              requestParams.folderId = value[subKey];
              break;
            }
          }
        }
      }
      
      // Check for folder ID saved in global context
      if (!requestParams.folderId && global._claude_last_folder_id) {
        logger.info(`Using folder ID from global context: ${global._claude_last_folder_id}`);
        requestParams.folderId = global._claude_last_folder_id;
      }
    } else if (global.__last_message?.params?.arguments) {
      // Use global last message
      requestParams = global.__last_message.params.arguments;
      logger.info(`Using global last message params`);
    } else {
      // Last resort - use inbox as default
      requestParams.folderId = 'inbox';
      logger.info(`Using default folder 'inbox'`);
    }
    
    logger.info(`Request params extracted: ${JSON.stringify(requestParams)}`);
    
    // Extract userId
    let userId = requestParams.userId;
    logger.info(`Extracted userId: ${userId}`);
    
    // Extract folder ID
    const folderId = requestParams.folderId || 'inbox';
    logger.info(`Extracted folderId: ${folderId}`);
    
    // Store folder ID in global context for potential future use
    if (folderId) {
      global._claude_last_folder_id = folderId;
    }
    
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
    
    const limit = Math.min(
      requestParams.limit || config.email.maxEmailsPerRequest, 
      config.email.maxEmailsPerRequest
    );
    
    logger.info(`Listing emails for user ${userId} in folder ${folderId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Determine endpoint based on folder ID
    let endpoint;
    if (folderId.toLowerCase() === 'inbox') {
      endpoint = '/me/mailFolders/inbox/messages';
    } else if (folderId.toLowerCase() === 'drafts') {
      endpoint = '/me/mailFolders/drafts/messages';
    } else if (folderId.toLowerCase() === 'sentitems') {
      endpoint = '/me/mailFolders/sentItems/messages';
    } else if (folderId.toLowerCase() === 'deleteditems') {
      endpoint = '/me/mailFolders/deletedItems/messages';
    } else {
      endpoint = `/me/mailFolders/${folderId}/messages`;
    }
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: requestParams.fields || config.email.defaultFields,
      top: limit,
      filter: requestParams.filter,
      orderBy: requestParams.orderBy || { receivedDateTime: 'desc' },
      skip: requestParams.skip || 0,
      search: requestParams.search
    });
    
    // Get emails
    const emails = await graphClient.getPaginated(endpoint, queryParams, {
      maxPages: requestParams.maxPages || 1
    });
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          folder: folderId,
          count: emails.length,
          emails: emails.map(email => formatEmailResponse(email)),
          hasMore: emails.length >= limit
        })
      }]
    };
  } catch (error) {
    logger.error(`Error listing emails: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list emails: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Format email response to clean up and improve readability
 * @param {Object} email - Raw email from Graph API
 * @returns {Object} - Formatted email
 */
function formatEmailResponse(email) {
  // Extract sender name and email
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
  
  // Build a cleaner email object
  return {
    id: email.id,
    subject: email.subject || '(No Subject)',
    sender,
    toRecipients,
    ccRecipients,
    receivedDateTime: email.receivedDateTime,
    sentDateTime: email.sentDateTime,
    preview: email.bodyPreview,
    hasAttachments: !!email.hasAttachments,
    isRead: email.isRead,
    importance: email.importance,
    isDraft: email.isDraft,
    webLink: email.webLink,
    categories: email.categories || []
  };
}

module.exports = {
  listEmailsHandler
};