const { listFoldersHandler, getFolderHandler } = require('./list');
const { createFolderHandler, updateFolderHandler, deleteFolderHandler } = require('./create');
const { moveEmailsHandler, moveFolderHandler, copyEmailsHandler } = require('./move');

// Folder tool definitions
const folderTools = [
  // List and Get
  {
    name: 'list_folders',
    description: 'List mail folders',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        parentFolderId: {
          type: 'string',
          description: 'Parent folder ID to list child folders (optional)'
        },
        limit: {
          type: 'number',
          description: 'Maximum number of folders to return (optional)'
        },
        filter: {
          type: 'object',
          description: 'OData filter criteria (optional)'
        },
        orderBy: {
          type: ['object', 'string', 'array'],
          description: 'OData orderby specification (optional, defaults to displayName asc)'
        }
      }
    },
    handler: listFoldersHandler
  },
  {
    name: 'get_folder',
    description: 'Get detailed information about a mail folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        folderId: {
          type: 'string',
          description: 'Folder ID or well-known folder name (inbox, drafts, sentitems, deleteditems)'
        }
      },
      required: ['folderId']
    },
    handler: getFolderHandler
  },
  
  // Create, Update, Delete
  {
    name: 'create_folder',
    description: 'Create a new mail folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        name: {
          type: 'string',
          description: 'Name for the new folder'
        },
        parentFolderId: {
          type: 'string',
          description: 'Parent folder ID (optional, defaults to root)'
        }
      },
      required: ['name']
    },
    handler: createFolderHandler
  },
  {
    name: 'update_folder',
    description: 'Update a mail folder name',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        folderId: {
          type: 'string',
          description: 'Folder ID to update'
        },
        name: {
          type: 'string',
          description: 'New name for the folder'
        }
      },
      required: ['folderId', 'name']
    },
    handler: updateFolderHandler
  },
  {
    name: 'delete_folder',
    description: 'Delete a mail folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        folderId: {
          type: 'string',
          description: 'Folder ID to delete'
        }
      },
      required: ['folderId']
    },
    handler: deleteFolderHandler
  },
  
  // Move and Copy
  {
    name: 'move_emails',
    description: 'Move emails to a folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailIds: {
          type: 'array',
          items: {
            type: 'string'
          },
          description: 'Array of email IDs to move'
        },
        destinationFolderId: {
          type: 'string',
          description: 'Destination folder ID'
        }
      },
      required: ['emailIds', 'destinationFolderId']
    },
    handler: moveEmailsHandler
  },
  {
    name: 'move_folder',
    description: 'Move a folder to another parent folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        folderId: {
          type: 'string',
          description: 'Folder ID to move'
        },
        destinationFolderId: {
          type: 'string',
          description: 'Destination parent folder ID'
        }
      },
      required: ['folderId', 'destinationFolderId']
    },
    handler: moveFolderHandler
  },
  {
    name: 'copy_emails',
    description: 'Copy emails to a folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailIds: {
          type: 'array',
          items: {
            type: 'string'
          },
          description: 'Array of email IDs to copy'
        },
        destinationFolderId: {
          type: 'string',
          description: 'Destination folder ID'
        }
      },
      required: ['emailIds', 'destinationFolderId']
    },
    handler: copyEmailsHandler
  }
];

module.exports = folderTools;