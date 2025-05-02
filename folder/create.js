// folder/create.js
const config = require('../config');
const logger = require('../utils/logger');
const { folder: folderApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const auth = require('../auth/index');

async function createFolderHandler(params = {}) {
  const name = params.name;
  const userId = params.userId || 'default';

  if (!name) return formatMcpResponse({ status: 'error', message: 'Folder name required' });

  try {
    logger.info(`Creating mail folder '${name}' for user ${userId}`);
    
    // Create the folder using folderApi
    // This will correctly handle authentication and token reuse
    const folderData = { displayName: name };
    const parentFolderId = params.parentFolderId; // May be undefined for root-level folders
    
    const folder = await folderApi.createFolder(userId, folderData, parentFolderId);
    
    return formatMcpResponse({
      status: 'success', 
      message: 'Folder created',
      folder: {
        id: folder.id,
        name: folder.displayName
      }
    });
  } catch (error) {
    logger.error(`Error creating folder: ${error.message}`);
    return formatMcpResponse({ status: 'error', message: `Failed: ${error.message}` });
  }
}

function formatMcpResponse(data) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(data)
      }
    ]
  };
}

module.exports = { createFolderHandler };
