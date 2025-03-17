const { listEmailsHandler } = require('./list');
const { searchEmailsHandler } = require('./search');
const { readEmailHandler, markEmailHandler } = require('./read');
const { 
  sendEmailHandler, 
  createDraftHandler, 
  replyEmailHandler, 
  forwardEmailHandler 
} = require('./send');
const { 
  getAttachmentHandler, 
  listAttachmentsHandler, 
  addAttachmentHandler, 
  deleteAttachmentHandler 
} = require('./attachments');

// Email tool definitions
const emailTools = [
  // List and Search
  {
    name: 'list_emails',
    description: 'List emails from a mailbox folder',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        folderId: {
          type: 'string',
          description: 'Folder ID or well-known folder name (inbox, drafts, sentitems, deleteditems) (optional, defaults to "inbox")'
        },
        limit: {
          type: 'number',
          description: 'Maximum number of emails to return (optional)'
        },
        skip: {
          type: 'number',
          description: 'Number of emails to skip (for pagination) (optional)'
        },
        filter: {
          type: 'object',
          description: 'OData filter criteria (optional)'
        },
        orderBy: {
          type: ['object', 'string', 'array'],
          description: 'OData orderby specification (optional, defaults to receivedDateTime desc)'
        },
        fields: {
          type: ['array', 'string'],
          description: 'Fields to include in the response (optional)'
        },
        search: {
          type: 'string',
          description: 'Search query (optional)'
        },
        maxPages: {
          type: 'number',
          description: 'Maximum number of pages to fetch (optional)'
        }
      }
    },
    handler: listEmailsHandler
  },
  {
    name: 'search_emails',
    description: 'Search for emails across all folders',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        query: {
          type: 'string',
          description: 'Search query'
        },
        limit: {
          type: 'number',
          description: 'Maximum number of emails to return (optional)'
        },
        fields: {
          type: ['array', 'string'],
          description: 'Fields to include in the response (optional)'
        },
        orderBy: {
          type: ['object', 'string', 'array'],
          description: 'OData orderby specification (optional, defaults to receivedDateTime desc)'
        },
        maxPages: {
          type: 'number',
          description: 'Maximum number of pages to fetch (optional)'
        }
      },
      required: ['query']
    },
    handler: searchEmailsHandler
  },
  
  // Read
  {
    name: 'read_email',
    description: 'Read a specific email by ID',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID to read'
        },
        markAsRead: {
          type: 'boolean',
          description: 'Whether to mark the email as read (optional, defaults to false)'
        }
      },
      required: ['emailId']
    },
    handler: readEmailHandler
  },
  {
    name: 'mark_email',
    description: 'Mark an email as read or unread',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID to mark'
        },
        isRead: {
          type: 'boolean',
          description: 'Whether to mark as read (true) or unread (false) (optional, defaults to true)'
        }
      },
      required: ['emailId']
    },
    handler: markEmailHandler
  },
  
  // Send and Reply
  {
    name: 'send_email',
    description: 'Send a new email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        subject: {
          type: 'string',
          description: 'Email subject'
        },
        body: {
          type: 'string',
          description: 'Email body content'
        },
        bodyType: {
          type: 'string',
          enum: ['Text', 'HTML'],
          description: 'Body content type (optional, defaults to HTML)'
        },
        to: {
          type: ['string', 'array'],
          description: 'Recipient(s) in format "name <email>" or just "email"'
        },
        cc: {
          type: ['string', 'array'],
          description: 'CC recipient(s) (optional)'
        },
        bcc: {
          type: ['string', 'array'],
          description: 'BCC recipient(s) (optional)'
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Email importance (optional, defaults to normal)'
        },
        saveToSentItems: {
          type: 'boolean',
          description: 'Whether to save to sent items (optional, defaults to true)'
        }
      },
      required: ['subject', 'body']
    },
    handler: sendEmailHandler
  },
  {
    name: 'create_draft',
    description: 'Create a draft email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        subject: {
          type: 'string',
          description: 'Email subject (optional)'
        },
        body: {
          type: 'string',
          description: 'Email body content (optional)'
        },
        bodyType: {
          type: 'string',
          enum: ['Text', 'HTML'],
          description: 'Body content type (optional, defaults to HTML)'
        },
        to: {
          type: ['string', 'array'],
          description: 'Recipient(s) (optional)'
        },
        cc: {
          type: ['string', 'array'],
          description: 'CC recipient(s) (optional)'
        },
        bcc: {
          type: ['string', 'array'],
          description: 'BCC recipient(s) (optional)'
        },
        importance: {
          type: 'string',
          enum: ['low', 'normal', 'high'],
          description: 'Email importance (optional, defaults to normal)'
        }
      }
    },
    handler: createDraftHandler
  },
  {
    name: 'reply_email',
    description: 'Reply to an email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID to reply to'
        },
        body: {
          type: 'string',
          description: 'Reply body content'
        },
        replyAll: {
          type: 'boolean',
          description: 'Whether to reply to all recipients (optional, defaults to false)'
        }
      },
      required: ['emailId', 'body']
    },
    handler: replyEmailHandler
  },
  {
    name: 'forward_email',
    description: 'Forward an email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID to forward'
        },
        to: {
          type: ['string', 'array'],
          description: 'Recipient(s) to forward to'
        },
        comment: {
          type: 'string',
          description: 'Additional comment to include (optional)'
        }
      },
      required: ['emailId', 'to']
    },
    handler: forwardEmailHandler
  },
  
  // Attachments
  {
    name: 'get_attachment',
    description: 'Get attachment content from an email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID containing the attachment'
        },
        attachmentId: {
          type: 'string',
          description: 'Attachment ID to get'
        }
      },
      required: ['emailId', 'attachmentId']
    },
    handler: getAttachmentHandler
  },
  {
    name: 'list_attachments',
    description: 'List attachments for an email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID to list attachments from'
        }
      },
      required: ['emailId']
    },
    handler: listAttachmentsHandler
  },
  {
    name: 'add_attachment',
    description: 'Add attachment to a draft email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID (must be a draft)'
        },
        name: {
          type: 'string',
          description: 'Attachment filename'
        },
        contentType: {
          type: 'string',
          description: 'MIME type of the attachment (optional, defaults to application/octet-stream)'
        },
        contentBytes: {
          type: 'string',
          description: 'Base64-encoded content of the attachment'
        },
        contentUrl: {
          type: 'string',
          description: 'URL to the attachment content (for reference attachments)'
        },
        providerType: {
          type: 'string',
          description: 'Provider type for reference attachments (optional)'
        },
        isInline: {
          type: 'boolean',
          description: 'Whether this is an inline attachment (optional, defaults to false)'
        }
      },
      required: ['emailId', 'name']
    },
    handler: addAttachmentHandler
  },
  {
    name: 'delete_attachment',
    description: 'Delete an attachment from a draft email',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        emailId: {
          type: 'string',
          description: 'Email ID (must be a draft)'
        },
        attachmentId: {
          type: 'string',
          description: 'Attachment ID to delete'
        }
      },
      required: ['emailId', 'attachmentId']
    },
    handler: deleteAttachmentHandler
  }
];

module.exports = emailTools;