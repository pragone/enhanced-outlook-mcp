const config = require('../config');
const logger = require('../utils/logger');
const { email: emailApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const auth = require('../auth/index');

/**
 * Get attachment content from an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Attachment content
 */
async function getAttachmentHandler(params = {}) {
  try {
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    const emailId = params.emailId;
    const attachmentId = params.attachmentId;
    
    if (!emailId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required'
          })
        }]
      };
    }
    
    if (!attachmentId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Attachment ID is required'
          })
        }]
      };
    }
    
    logger.info(`Getting attachment ${attachmentId} from email ${emailId} for user ${userId}`);
    
    // Get attachment using emailApi
    // This will correctly handle authentication and token reuse
    const attachment = await emailApi.getAttachment(userId, emailId, attachmentId);
    
    if (!attachment) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Attachment not found'
          })
        }]
      };
    }
    
    // Format the response based on attachment type
    const formattedAttachment = {
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline
    };
    
    // For file attachments, include the content
    if (attachment.contentBytes) {
      formattedAttachment.contentBytes = attachment.contentBytes;
    }
    
    // For item attachments, include item details
    if (attachment['@odata.type'] === '#microsoft.graph.itemAttachment') {
      formattedAttachment.itemType = attachment.itemType;
      formattedAttachment.item = attachment.item;
    }
    
    // For reference attachments, include reference details
    if (attachment['@odata.type'] === '#microsoft.graph.referenceAttachment') {
      formattedAttachment.sourceUrl = attachment.sourceUrl;
      formattedAttachment.providerType = attachment.providerType;
      formattedAttachment.thumbnailUrl = attachment.thumbnailUrl;
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          attachment: formattedAttachment
        })
      }]
    };
  } catch (error) {
    logger.error(`Error getting attachment: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to get attachment: ${error.message}`
        })
      }]
    };
  }
}

/**
 * List attachments for an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of attachments
 */
async function listAttachmentsHandler(params = {}) {
  try {
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    const emailId = params.emailId;
    
    if (!emailId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required'
          })
        }]
      };
    }
    
    logger.info(`Listing attachments for email ${emailId} for user ${userId}`);
    
    // Get attachments using emailApi
    // This will correctly handle authentication and token reuse
    const response = await emailApi.listAttachments(userId, emailId);
    
    if (!response || !response.value) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Failed to retrieve attachments'
          })
        }]
      };
    }
    
    // Format the attachments
    const attachments = response.value.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      lastModifiedDateTime: attachment.lastModifiedDateTime,
      attachmentType: attachment['@odata.type']?.replace('#microsoft.graph.', '')
    }));
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          emailId,
          count: attachments.length,
          attachments
        })
      }]
    };
  } catch (error) {
    logger.error(`Error listing attachments: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list attachments: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Add attachment to a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Attachment result
 */
async function addAttachmentHandler(params = {}) {
  try {
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    const emailId = params.emailId;
    
    if (!emailId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required'
          })
        }]
      };
    }
    
    if (!params.name) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Attachment name is required'
          })
        }]
      };
    }
    
    if (!params.contentBytes && !params.contentUrl) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Either contentBytes or contentUrl is required'
          })
        }]
      };
    }
    
    logger.info(`Adding attachment to email ${emailId} for user ${userId}`);
    
    // Prepare attachment data
    const attachmentData = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: params.name,
      contentType: params.contentType || 'application/octet-stream',
      isInline: params.isInline === true
    };
    
    // Add content bytes or URL
    if (params.contentBytes) {
      attachmentData.contentBytes = params.contentBytes;
    } else if (params.contentUrl) {
      attachmentData['@odata.type'] = '#microsoft.graph.referenceAttachment';
      attachmentData.sourceUrl = params.contentUrl;
      attachmentData.providerType = params.providerType || 'other';
    }
    
    // Use emailApi.addAttachment to add the attachment
    // This will correctly handle authentication and token reuse
    const attachment = await emailApi.addAttachment(userId, emailId, attachmentData);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Attachment added successfully',
          emailId,
          attachmentId: attachment.id
        })
      }]
    };
  } catch (error) {
    logger.error(`Error adding attachment: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to add attachment: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Delete an attachment from a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Deletion result
 */
async function deleteAttachmentHandler(params = {}) {
  try {
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    const emailId = params.emailId;
    const attachmentId = params.attachmentId;
    
    if (!emailId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email ID is required'
          })
        }]
      };
    }
    
    if (!attachmentId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Attachment ID is required'
          })
        }]
      };
    }
    
    logger.info(`Deleting attachment ${attachmentId} from email ${emailId} for user ${userId}`);
    
    // Delete attachment using emailApi
    // This will correctly handle authentication and token reuse
    await emailApi.deleteAttachment(userId, emailId, attachmentId);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Attachment deleted successfully',
          emailId,
          attachmentId
        })
      }]
    };
  } catch (error) {
    logger.error(`Error deleting attachment: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to delete attachment: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Handler to get email attachments
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Response with status and attachment data
 */
async function getAttachmentsHandler(params = {}) {
  try {
    let userId = params.userId;
    const messageId = params.messageId;
    
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
              message: 'Multiple users found. Please specify userId parameter to indicate which account to use.'
            })
          }]
        };
      }
    }
    
    if (!messageId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'messageId parameter is required'
          })
        }]
      };
    }
    
    logger.info(`Getting attachments for message ${messageId} for user ${userId}`);
    
    // Get attachments using emailApi
    // This will correctly handle authentication and token reuse
    const response = await emailApi.listAttachments(userId, messageId);
    
    if (!response || !response.value) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Failed to retrieve attachments'
          })
        }]
      };
    }
    
    // Format the attachments
    const attachments = response.value.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      lastModifiedDateTime: attachment.lastModifiedDateTime,
      attachmentType: attachment['@odata.type']?.replace('#microsoft.graph.', '')
    }));
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          messageId,
          count: attachments.length,
          attachments
        })
      }]
    };
  } catch (error) {
    logger.error(`Error getting attachments: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to get attachments: ${error.message}`
        })
      }]
    };
  }
}

module.exports = {
  getAttachmentHandler,
  listAttachmentsHandler,
  addAttachmentHandler,
  deleteAttachmentHandler,
  getAttachmentsHandler
};