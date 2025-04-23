const { listFoldersHandler, getFolderHandler } = require('./list');
const { createFolderHandler, updateFolderHandler, deleteFolderHandler } = require('./create');
const { moveEmailsHandler, moveFolderHandler, copyEmailsHandler } = require('./move');

// Export all handlers directly
module.exports = {
  // List and Get
  listFoldersHandler,
  getFolderHandler,
  
  // Create, Update, Delete
  createFolderHandler,
  updateFolderHandler,
  deleteFolderHandler,
  
  // Move and Copy
  moveEmailsHandler,
  moveFolderHandler,
  copyEmailsHandler
};