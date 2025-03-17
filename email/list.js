const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');
const { buildQueryParams } = require('../utils/odata-helpers');

/**
 * List emails from a mailbox
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of emails
 */
async function listEmailsHandler(params = {}) {
  const userId = params.userId || 'default';
  const folderId = params.folderId || 'inbox';
  const limit = Math.min(
    params.limit || config.email.maxEmailsPerRequest, 
    config.email.maxEmailsPerRequest
  );
  
  try {
    logger.info(`Listing emails for user ${userId} in folder ${folderId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Determine endpoint based on folder ID
    let endpoint;
    if (folderId.toLowerCase() === 'inbox') {
      endpoint = '/me/mailFolders/inbox/messages';
    } else if (folderId.toLowerCase() === 'drafts') {
      endpoint = '/me/mailFolders/drafts/messages';
    } else if (folderId.toLowerCase() === 'sentitems') {
      endpoint = '/me/mailFolders/sentItems/messages';
    } else if (folderId.toLowerCase() === 'deleteditems') {
      endpoint = '/me/mailFolders/deletedItems/messages';
    } else {
      endpoint = `/me/mailFolders/${folderId}/messages`;
    }
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: params.fields || config.email.defaultFields,
      top: limit,
      filter: params.filter,
      orderBy: params.orderBy || { receivedDateTime: 'desc' },
      skip: params.skip || 0,
      search: params.search
    });
    
    // Get emails
    const emails = await graphClient.getPaginated(endpoint, queryParams, {
      maxPages: params.maxPages || 1
    });
    
    return {
      status: 'success',
      folder: folderId,
      count: emails.length,
      emails: emails.map(email => formatEmailResponse(email)),
      hasMore: emails.length >= limit
    };
  } catch (error) {
    logger.error(`Error listing emails: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to list emails: ${error.message}`
    };
  }
}

/**
 * Format email response to clean up and improve readability
 * @param {Object} email - Raw email from Graph API
 * @returns {Object} - Formatted email
 */
function formatEmailResponse(email) {
  // Extract sender name and email
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
  
  // Build a cleaner email object
  return {
    id: email.id,
    subject: email.subject || '(No Subject)',
    sender,
    toRecipients,
    ccRecipients,
    receivedDateTime: email.receivedDateTime,
    sentDateTime: email.sentDateTime,
    preview: email.bodyPreview,
    hasAttachments: !!email.hasAttachments,
    isRead: email.isRead,
    importance: email.importance,
    isDraft: email.isDraft,
    webLink: email.webLink,
    categories: email.categories || []
  };
}

module.exports = {
  listEmailsHandler
};