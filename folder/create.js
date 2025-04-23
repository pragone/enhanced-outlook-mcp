// folder/create.js
const { GraphApiClient } = require('../utils/graph-api');

async function createFolderHandler(params = {}) {
  const { userId = 'default', name } = params;
  if (!name) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'Folder name required'
        })
      }]
    };
  }

  try {
    const graphClient = new GraphApiClient(userId);
    const folder = await graphClient.post('/me/mailFolders', { displayName: name });
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Folder created',
          folder: { id: folder.id, name: folder.displayName }
        })
      }]
    };
  } catch (error) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed: ${error.message}`
        })
      }]
    };
  }
}

module.exports = { createFolderHandler };
