const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');

/**
 * Get attachment content from an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Attachment content
 */
async function getAttachmentHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  const attachmentId = params.attachmentId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  if (!attachmentId) {
    return {
      status: 'error',
      message: 'Attachment ID is required'
    };
  }
  
  try {
    logger.info(`Getting attachment ${attachmentId} from email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Get attachment
    const attachment = await graphClient.get(`/me/messages/${emailId}/attachments/${attachmentId}`);
    
    if (!attachment) {
      return {
        status: 'error',
        message: 'Attachment not found'
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
      status: 'success',
      attachment: formattedAttachment
    };
  } catch (error) {
    logger.error(`Error getting attachment: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to get attachment: ${error.message}`
    };
  }
}

/**
 * List attachments for an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of attachments
 */
async function listAttachmentsHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  try {
    logger.info(`Listing attachments for email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Get attachments
    const response = await graphClient.get(`/me/messages/${emailId}/attachments`);
    
    if (!response || !response.value) {
      return {
        status: 'error',
        message: 'Failed to retrieve attachments'
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
      status: 'success',
      emailId,
      count: attachments.length,
      attachments
    };
  } catch (error) {
    logger.error(`Error listing attachments: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to list attachments: ${error.message}`
    };
  }
}

/**
 * Add attachment to a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Attachment result
 */
async function addAttachmentHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  if (!params.name) {
    return {
      status: 'error',
      message: 'Attachment name is required'
    };
  }
  
  if (!params.contentBytes && !params.contentUrl) {
    return {
      status: 'error',
      message: 'Either contentBytes or contentUrl is required'
    };
  }
  
  try {
    logger.info(`Adding attachment to email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
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
    
    // Add attachment to email
    const attachment = await graphClient.post(`/me/messages/${emailId}/attachments`, attachmentData);
    
    return {
      status: 'success',
      message: 'Attachment added successfully',
      emailId,
      attachmentId: attachment.id,
      attachmentName: params.name
    };
  } catch (error) {
    logger.error(`Error adding attachment: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to add attachment: ${error.message}`
    };
  }
}

/**
 * Delete an attachment from a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Deletion result
 */
async function deleteAttachmentHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  const attachmentId = params.attachmentId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  if (!attachmentId) {
    return {
      status: 'error',
      message: 'Attachment ID is required'
    };
  }
  
  try {
    logger.info(`Deleting attachment ${attachmentId} from email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Delete the attachment
    await graphClient.delete(`/me/messages/${emailId}/attachments/${attachmentId}`);
    
    return {
      status: 'success',
      message: 'Attachment deleted successfully',
      emailId,
      attachmentId
    };
  } catch (error) {
    logger.error(`Error deleting attachment: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to delete attachment: ${error.message}`
    };
  }
}

module.exports = {
  getAttachmentHandler,
  listAttachmentsHandler,
  addAttachmentHandler,
  deleteAttachmentHandler
};