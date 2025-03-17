const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');

/**
 * Read a specific email by ID
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Email content
 */
async function readEmailHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  try {
    logger.info(`Reading email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Get email with detailed content
    const email = await graphClient.get(`/me/messages/${emailId}`, {
      $expand: 'attachments'
    });
    
    if (!email) {
      return {
        status: 'error',
        message: `Email not found with ID: ${emailId}`
      };
    }
    
    // Mark as read if requested
    if (params.markAsRead && !email.isRead) {
      await graphClient.patch(`/me/messages/${emailId}`, {
        isRead: true
      });
    }
    
    // Format the detailed email response
    const formattedEmail = formatDetailedEmail(email);
    
    return {
      status: 'success',
      email: formattedEmail
    };
  } catch (error) {
    logger.error(`Error reading email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to read email: ${error.message}`
    };
  }
}

/**
 * Format detailed email response
 * @param {Object} email - Raw email from Graph API
 * @returns {Object} - Formatted detailed email
 */
function formatDetailedEmail(email) {
  // Extract sender information
  let sender = null;
  if (email.from && email.from.emailAddress) {
    sender = {
      name: email.from.emailAddress.name,
      email: email.from.emailAddress.address
    };
  }
  
  // Extract recipients
  let toRecipients = [];
  if (email.toRecipients && Array.isArray(email.toRecipients)) {
    toRecipients = email.toRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Extract CC recipients
  let ccRecipients = [];
  if (email.ccRecipients && Array.isArray(email.ccRecipients)) {
    ccRecipients = email.ccRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Extract BCC recipients
  let bccRecipients = [];
  if (email.bccRecipients && Array.isArray(email.bccRecipients)) {
    bccRecipients = email.bccRecipients.map(recipient => ({
      name: recipient.emailAddress.name,
      email: recipient.emailAddress.address
    }));
  }
  
  // Format attachments
  let attachments = [];
  if (email.attachments && Array.isArray(email.attachments)) {
    attachments = email.attachments.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId
    }));
  }
  
  // Format body content
  let body = {
    contentType: email.body ? email.body.contentType : 'text',
    content: email.body ? email.body.content : ''
  };
  
  // Create formatted response
  return {
    id: email.id,
    subject: email.subject || '(No Subject)',
    sender,
    toRecipients,
    ccRecipients,
    bccRecipients,
    receivedDateTime: email.receivedDateTime,
    sentDateTime: email.sentDateTime,
    hasAttachments: !!email.hasAttachments,
    attachments,
    isRead: email.isRead,
    isDraft: email.isDraft,
    importance: email.importance,
    body,
    conversationId: email.conversationId,
    parentFolderId: email.parentFolderId,
    internetMessageId: email.internetMessageId,
    webLink: email.webLink,
    categories: email.categories || []
  };
}

/**
 * Mark an email as read or unread
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Operation result
 */
async function markEmailHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  const isRead = params.isRead === undefined ? true : !!params.isRead;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  try {
    logger.info(`Marking email ${emailId} as ${isRead ? 'read' : 'unread'} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Update the read status
    await graphClient.patch(`/me/messages/${emailId}`, {
      isRead
    });
    
    return {
      status: 'success',
      message: `Email marked as ${isRead ? 'read' : 'unread'} successfully`,
      emailId,
      isRead
    };
  } catch (error) {
    logger.error(`Error marking email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to mark email: ${error.message}`
    };
  }
}

module.exports = {
  readEmailHandler,
  markEmailHandler
};