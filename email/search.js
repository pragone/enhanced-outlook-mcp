const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');
const { buildQueryParams } = require('../utils/odata-helpers');

/**
 * Search emails across mailbox
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Search results
 */
async function searchEmailsHandler(params = {}) {
  const userId = params.userId || 'default';
  const query = params.query;
  
  if (!query) {
    return {
      status: 'error',
      message: 'Search query is required'
    };
  }
  
  const limit = Math.min(
    params.limit || config.email.maxEmailsPerRequest, 
    config.email.maxEmailsPerRequest
  );
  
  try {
    logger.info(`Searching emails for user ${userId} with query: ${query}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: params.fields || config.email.defaultFields,
      top: limit,
      search: query,
      orderBy: params.orderBy || { receivedDateTime: 'desc' }
    });
    
    // Perform email search
    // Microsoft Graph allows searching across all folders with /me/messages endpoint
    const endpoint = '/me/messages';
    const emails = await graphClient.getPaginated(endpoint, queryParams, {
      maxPages: params.maxPages || 1
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
      status: 'success',
      query,
      totalResults: emails.length,
      folderResults: Object.values(folderGroups),
      hasMore: emails.length >= limit
    };
  } catch (error) {
    logger.error(`Error searching emails: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to search emails: ${error.message}`
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