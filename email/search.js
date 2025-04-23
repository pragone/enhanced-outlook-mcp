const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');
const { listUsers } = require('../auth/token-manager');
const { buildQueryParams } = require('../utils/odata-helpers');

/**
 * Search emails across mailbox
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Search results
 */
async function searchEmailsHandler(params = {}) {
  try {
    // Super detailed debugging
    logger.info(`SEARCH EMAILS HANDLER START ------------------------------------`);
    logger.info(`Raw params: ${JSON.stringify(params)}`);
    
    // CRITICAL FIX FOR CLAUDE DESKTOP:
    // Check if this is a direct JSON-RPC request with the tool call format
    if (global.__last_message?.method === 'tools/call' && 
        global.__last_message?.params?.name === 'search_emails' &&
        global.__last_message?.params?.arguments) {
      
      const directArgs = global.__last_message.params.arguments;
      logger.info(`Found direct JSON-RPC arguments: ${JSON.stringify(directArgs)}`);
      
      // Use the arguments directly from the JSON-RPC request
      if (directArgs.query) {
        params = directArgs;
        logger.info(`Using direct JSON-RPC arguments for search_emails`);
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
    // 5. Global last message
    
    if (params.query) {
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
      
      // Check if params has any properties that might be the search query
      for (const key in params) {
        // Skip known non-query properties
        if (['signal', 'tool'].includes(key)) continue;
        
        const value = params[key];
        if (typeof value === 'string' && value.length > 0) {
          // This could be a search query
          logger.info(`Found potential search query in property ${key}: ${value}`);
          requestParams.query = value;
          break;
        }
        
        // Check if value is an object that might contain the query
        if (typeof value === 'object' && value !== null) {
          for (const subKey in value) {
            if (['query', 'q', 'search'].includes(subKey)) {
              logger.info(`Found search query in nested property ${key}.${subKey}: ${value[subKey]}`);
              requestParams.query = value[subKey];
              break;
            }
          }
        }
      }
      
      // Check for query saved in global context
      if (!requestParams.query && global._claude_last_search_query) {
        logger.info(`Using search query from global context: ${global._claude_last_search_query}`);
        requestParams.query = global._claude_last_search_query;
      }
    } else if (global.__last_message?.params?.arguments) {
      // Use global last message
      requestParams = global.__last_message.params.arguments;
      logger.info(`Using global last message params`);
    } else {
      // Last resort - check globals directly for previous search query
      if (global._claude_last_search_query) {
        requestParams.query = global._claude_last_search_query;
        logger.info(`Using _claude_last_search_query from global context: ${global._claude_last_search_query}`);
      }
    }
    
    logger.info(`Request params extracted: ${JSON.stringify(requestParams)}`);
    
    // Extract userId
    let userId = requestParams.userId;
    logger.info(`Extracted userId: ${userId}`);
    
    // Extract query and ensure we have one
    const query = requestParams.query || '';
    logger.info(`Extracted query: ${query}`);
    
    // Store search query in global context for potential future use
    if (query) {
      global._claude_last_search_query = query;
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    
    if (!query) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'query parameter is required'
          })
        }]
      };
    }
    
    const limit = Math.min(
      requestParams.limit || config.email.maxEmailsPerRequest, 
      config.email.maxEmailsPerRequest
    );
    
    logger.info(`Searching emails for user ${userId} with query: ${query}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: requestParams.fields || config.email.defaultFields,
      top: limit,
      search: query,
      orderBy: requestParams.orderBy || { receivedDateTime: 'desc' }
    });
    
    // Perform email search
    // Microsoft Graph allows searching across all folders with /me/messages endpoint
    const endpoint = '/me/messages';
    const emails = await graphClient.getPaginated(endpoint, queryParams, {
      maxPages: requestParams.maxPages || 1
    });
    
    // Group results by folder for better organization
    const folderGroups = {};
    for (const email of emails) {
      const parentFolderId = email.parentFolderId || 'unknown';
      
      if (!folderGroups[parentFolderId]) {
        folderGroups[parentFolderId] = {
          folderId: parentFolderId,
          folderName: 'Unknown', // We'll populate this later
          emails: []
        };
      }
      
      folderGroups[parentFolderId].emails.push(formatEmailResult(email));
    }
    
    // Fetch folder names if we have results
    if (Object.keys(folderGroups).length > 0) {
      await populateFolderNames(graphClient, folderGroups);
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          query,
          totalResults: emails.length,
          folderResults: Object.values(folderGroups),
          hasMore: emails.length >= limit
        })
      }]
    };
  } catch (error) {
    logger.error(`Error searching emails: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to search emails: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Format email search result
 * @param {Object} email - Raw email from Graph API
 * @returns {Object} - Formatted email result
 */
function formatEmailResult(email) {
  // Extract sender name and email
  let sender = null;
  if (email.from && email.from.emailAddress) {
    sender = {
      name: email.from.emailAddress.name,
      email: email.from.emailAddress.address
    };
  }
  
  // Build a cleaner email object
  return {
    id: email.id,
    subject: email.subject || '(No Subject)',
    sender,
    receivedDateTime: email.receivedDateTime,
    preview: email.bodyPreview,
    hasAttachments: !!email.hasAttachments,
    isRead: email.isRead,
    importance: email.importance,
    isDraft: email.isDraft
  };
}

/**
 * Populate folder names for search results
 * @param {GraphApiClient} graphClient - Graph API client
 * @param {Object} folderGroups - Folder groups to populate
 * @returns {Promise<void>}
 */
async function populateFolderNames(graphClient, folderGroups) {
  const folderIds = Object.keys(folderGroups).filter(id => id !== 'unknown');
  
  if (folderIds.length === 0) {
    return;
  }
  
  try {
    // Get well-known folder names first
    const wellKnownFolders = await graphClient.get('/me/mailFolders', {
      $select: 'id,displayName',
      $filter: "wellKnownName ne null"
    });
    
    // Map well-known folders
    if (wellKnownFolders && wellKnownFolders.value) {
      for (const folder of wellKnownFolders.value) {
        if (folderGroups[folder.id]) {
          folderGroups[folder.id].folderName = folder.displayName;
        }
      }
    }
    
    // Get names for remaining folders
    const remainingFolderIds = folderIds.filter(id => 
      folderGroups[id].folderName === 'Unknown'
    );
    
    // Batch remaining folders in groups of 10 to avoid long URLs
    const batchSize = 10;
    for (let i = 0; i < remainingFolderIds.length; i += batchSize) {
      const batch = remainingFolderIds.slice(i, i + batchSize);
      
      // Build filter to get multiple folders by ID
      const filterClauses = batch.map(id => `id eq '${id}'`);
      const filter = filterClauses.join(' or ');
      
      const folderBatch = await graphClient.get('/me/mailFolders', {
        $select: 'id,displayName',
        $filter: filter
      });
      
      if (folderBatch && folderBatch.value) {
        for (const folder of folderBatch.value) {
          if (folderGroups[folder.id]) {
            folderGroups[folder.id].folderName = folder.displayName;
          }
        }
      }
    }
  } catch (error) {
    logger.warn(`Error populating folder names: ${error.message}`);
    // We'll just continue with Unknown for folder names that couldn't be populated
  }
}

module.exports = {
  searchEmailsHandler
};