const config = require('../config');
const logger = require('../utils/logger');
const { folder: folderApi, email: emailApi } = require('../utils/graph-api-adapter');
const { buildQueryParams } = require('../utils/odata-helpers');
const { listUsers } = require('../auth/token-manager');
const auth = require('../auth/index');
const { normalizeParameters } = require('../utils/parameter-helpers');

/**
 * List mail folders
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of folders
 */
async function listFoldersHandler(params = {}) {
  let userId = params.userId;
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
    userId = users.length === 1 ? users[0] : params.userId;
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
  const parentFolderId = params.parentFolderId;
  
  try {
    logger.info(`Listing mail folders for user ${userId}`);
    
    // Get folders using folderApi to ensure proper authentication handling
    const response = await folderApi.listFolders(userId, {
      parentFolderId: parentFolderId,
      top: params.limit || 100
    });
    
    if (!response || !response.value) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Failed to retrieve folders'
          })
        }]
      };
    }
    
    const folders = response.value.map(folder => ({
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      childFolderCount: folder.childFolderCount,
      itemCount: folder.totalItemCount,
      unreadItemCount: folder.unreadItemCount
    }));
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          count: folders.length,
          parentFolderId: parentFolderId || 'root',
          folders
        })
      }]
    };
  } catch (error) {
    logger.error(`Error listing folders: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list folders: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Get information about a specific folder
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Folder info
 */
async function getFolderHandler(params = {}) {
  let userId = params.userId;
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
    userId = users.length === 1 ? users[0] : params.userId;
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
  const folderId = params.folderId;
  
  if (!folderId) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'Folder ID is required'
        })
      }]
    };
  }
  
  try {
    logger.info(`Getting folder ${folderId} for user ${userId}`);
    
    // Normalize well-known folder IDs if needed
    const normalizedFolderId = ['inbox', 'drafts', 'sentitems', 'deleteditems'].includes(folderId.toLowerCase())
      ? folderId.toLowerCase()
      : folderId;
    
    // Use folderApi to get the folder details
    const folder = await folderApi.getFolder(userId, normalizedFolderId);
    
    if (!folder) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: `Folder not found with ID: ${folderId}`
          })
        }]
      };
    }
    
    // Get child folders
    const childFoldersResponse = await folderApi.listFolders(userId, {
      parentFolderId: normalizedFolderId
    });
    
    const childFolders = childFoldersResponse?.value 
      ? childFoldersResponse.value.map(child => ({
          id: child.id,
          name: child.displayName,
          childFolderCount: child.childFolderCount,
          itemCount: child.totalItemCount,
          unreadItemCount: child.unreadItemCount
        }))
      : [];
    
    // Get recent messages
    const recentMessagesResponse = await emailApi.listMessages(userId, {
      folderId: normalizedFolderId,
      top: 5,
      orderBy: { receivedDateTime: 'desc' },
      select: ['id', 'subject', 'from', 'receivedDateTime', 'isRead']
    });
    
    const recentMessages = recentMessagesResponse?.value
      ? recentMessagesResponse.value.map(message => ({
          id: message.id,
          subject: message.subject || '(No Subject)',
          sender: message.from ? {
            name: message.from.emailAddress.name,
            email: message.from.emailAddress.address
          } : null,
          receivedDateTime: message.receivedDateTime,
          isRead: message.isRead
        }))
      : [];
    
    // Build folder info response
    const folderInfo = {
      id: folder.id,
      name: folder.displayName,
      parentFolderId: folder.parentFolderId,
      childFolderCount: folder.childFolderCount,
      itemCount: folder.totalItemCount,
      unreadItemCount: folder.unreadItemCount,
      childFolders,
      recentMessages
    };
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          folder: folderInfo
        })
      }]
    };
  } catch (error) {
    logger.error(`Error getting folder: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to get folder: ${error.message}`
        })
      }]
    };
  }
}

module.exports = {
  listFoldersHandler,
  getFolderHandler
};