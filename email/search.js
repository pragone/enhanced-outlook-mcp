const config = require('../config');
const logger = require('../utils/logger');
const { email: emailApi, folder: folderApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const auth = require('../auth/index');
const { buildQueryParams } = require('../utils/odata-helpers');

/**
 * Search emails across mailbox
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Search results
 */
async function searchEmailsHandler(params = {}) {
  try {
    // Debug logging
    logger.info(`Search emails handler started`);
    
    // Process parameters
    const requestParams = params.arguments ? params.arguments : params;
    
    // Extract userId
    let userId = requestParams.userId;
    
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
    
    // Extract query
    const query = requestParams.query || '';
    
    if (!query) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Query parameter is required'
          })
        }]
      };
    }
    
    // Calculate limit
    const limit = Math.min(
      requestParams.limit || config.email.maxEmailsPerRequest,
      config.email.maxEmailsPerRequest
    );
    
    logger.info(`Searching emails for user ${userId} with query: ${query}`);
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: requestParams.fields || config.email.defaultFields,
      top: limit,
      search: query
    });
    
    // Use emailApi to search messages
    const emailsResponse = await emailApi.listMessages(userId, queryParams);
    const emails = emailsResponse.value || [];
    
    // Sort results if needed
    if (requestParams.orderBy) {
      const sortField = Object.keys(requestParams.orderBy)[0] || 'receivedDateTime';
      const sortDirection = requestParams.orderBy[sortField] || 'desc';
      
      emails.sort((a, b) => {
        const valueA = a[sortField];
        const valueB = b[sortField];
        
        // Compare based on type
        if (valueA instanceof Date && valueB instanceof Date) {
          return sortDirection === 'asc' 
            ? valueA.getTime() - valueB.getTime()
            : valueB.getTime() - valueA.getTime();
        } else if (typeof valueA === 'string' && typeof valueB === 'string') {
          return sortDirection === 'asc'
            ? valueA.localeCompare(valueB)
            : valueB.localeCompare(valueA);
        } else {
          return sortDirection === 'asc' ? valueA - valueB : valueB - valueA;
        }
      });
    } else {
      // Default sort by receivedDateTime desc
      emails.sort((a, b) => {
        const dateA = new Date(a.receivedDateTime || 0);
        const dateB = new Date(b.receivedDateTime || 0);
        return dateB.getTime() - dateA.getTime();
      });
    }
    
    // Group results by folder
    const folderGroups = {};
    for (const email of emails) {
      const parentFolderId = email.parentFolderId || 'unknown';
      
      if (!folderGroups[parentFolderId]) {
        folderGroups[parentFolderId] = {
          folderId: parentFolderId,
          folderName: 'Unknown',
          emails: []
        };
      }
      
      folderGroups[parentFolderId].emails.push(formatEmailResult(email));
    }
    
    // Fetch folder names
    await fetchFolderNames(userId, folderGroups);
    
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
 * Fetch folder names for search results
 * @param {string} userId - User ID
 * @param {Object} folderGroups - Folder groups to populate
 * @returns {Promise<void>}
 */
async function fetchFolderNames(userId, folderGroups) {
  const folderIds = Object.keys(folderGroups).filter(id => id !== 'unknown');
  
  if (folderIds.length === 0) {
    return;
  }
  
  try {
    // Get folder information using folderApi
    for (const folderId of folderIds) {
      try {
        const folderInfo = await folderApi.getFolder(userId, folderId);
        if (folderInfo && folderInfo.displayName) {
          folderGroups[folderId].folderName = folderInfo.displayName;
        }
      } catch (folderError) {
        logger.warn(`Could not get info for folder ${folderId}: ${folderError.message}`);
      }
    }
  } catch (error) {
    logger.warn(`Error fetching folder names: ${error.message}`);
  }
}

module.exports = {
  searchEmailsHandler
}; 