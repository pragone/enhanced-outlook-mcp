const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');

/**
 * Send a new email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Send result
 */
async function sendEmailHandler(params = {}) {
  const userId = params.userId || 'default';
  
  // Check required parameters
  if (!params.subject) {
    return {
      status: 'error',
      message: 'Email subject is required'
    };
  }
  
  if (!params.body) {
    return {
      status: 'error',
      message: 'Email body is required'
    };
  }
  
  // At least one recipient is required (to, cc, or bcc)
  if (!params.to && !params.cc && !params.bcc) {
    return {
      status: 'error',
      message: 'At least one recipient (to, cc, or bcc) is required'
    };
  }
  
  try {
    logger.info(`Sending email for user ${userId} with subject: ${params.subject}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Prepare recipients
    const toRecipients = formatRecipients(params.to);
    const ccRecipients = formatRecipients(params.cc);
    const bccRecipients = formatRecipients(params.bcc);
    
    // Prepare email message
    const message = {
      subject: params.subject,
      body: {
        contentType: params.bodyType || 'HTML',
        content: params.body
      },
      toRecipients,
      ccRecipients,
      bccRecipients
    };
    
    // Set importance if provided
    if (params.importance) {
      message.importance = params.importance.toUpperCase();
    }
    
    // Send email using sendMail endpoint
    await graphClient.post('/me/sendMail', {
      message,
      saveToSentItems: params.saveToSentItems !== false
    });
    
    return {
      status: 'success',
      message: 'Email sent successfully',
      subject: params.subject,
      recipientCount: {
        to: toRecipients.length,
        cc: ccRecipients.length,
        bcc: bccRecipients.length
      }
    };
  } catch (error) {
    logger.error(`Error sending email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to send email: ${error.message}`
    };
  }
}

/**
 * Create a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Draft creation result
 */
async function createDraftHandler(params = {}) {
  const userId = params.userId || 'default';
  
  try {
    logger.info(`Creating email draft for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Prepare recipients
    const toRecipients = formatRecipients(params.to);
    const ccRecipients = formatRecipients(params.cc);
    const bccRecipients = formatRecipients(params.bcc);
    
    // Prepare draft message
    const message = {
      subject: params.subject || '',
      body: {
        contentType: params.bodyType || 'HTML',
        content: params.body || ''
      },
      toRecipients,
      ccRecipients,
      bccRecipients,
      isDraft: true
    };
    
    // Set importance if provided
    if (params.importance) {
      message.importance = params.importance.toUpperCase();
    }
    
    // Create draft by saving to drafts folder
    const draftEmail = await graphClient.post('/me/messages', message);
    
    return {
      status: 'success',
      message: 'Draft email created successfully',
      draftId: draftEmail.id,
      subject: params.subject,
      webLink: draftEmail.webLink
    };
  } catch (error) {
    logger.error(`Error creating draft email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to create draft email: ${error.message}`
    };
  }
}

/**
 * Reply to an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Reply result
 */
async function replyEmailHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  if (!params.body) {
    return {
      status: 'error',
      message: 'Reply body is required'
    };
  }
  
  try {
    logger.info(`Replying to email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Determine if it's a reply or reply all
    const endpoint = params.replyAll 
      ? `/me/messages/${emailId}/replyAll`
      : `/me/messages/${emailId}/reply`;
    
    // Send reply
    await graphClient.post(endpoint, {
      comment: params.body
    });
    
    return {
      status: 'success',
      message: `${params.replyAll ? 'Reply all' : 'Reply'} sent successfully`,
      emailId
    };
  } catch (error) {
    logger.error(`Error replying to email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to reply to email: ${error.message}`
    };
  }
}

/**
 * Forward an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Forward result
 */
async function forwardEmailHandler(params = {}) {
  const userId = params.userId || 'default';
  const emailId = params.emailId;
  
  if (!emailId) {
    return {
      status: 'error',
      message: 'Email ID is required'
    };
  }
  
  if (!params.to) {
    return {
      status: 'error',
      message: 'At least one recipient is required'
    };
  }
  
  try {
    logger.info(`Forwarding email ${emailId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Format recipients
    const toRecipients = formatRecipients(params.to);
    
    // Forward the email
    await graphClient.post(`/me/messages/${emailId}/forward`, {
      comment: params.comment || '',
      toRecipients
    });
    
    return {
      status: 'success',
      message: 'Email forwarded successfully',
      emailId,
      recipientCount: toRecipients.length
    };
  } catch (error) {
    logger.error(`Error forwarding email: ${error.message}`);
    
    return {
      status: 'error',
      message: `Failed to forward email: ${error.message}`
    };
  }
}

/**
 * Format recipients for API request
 * @param {string|Array} recipients - Recipients as string or array
 * @returns {Array} - Formatted recipients
 */
function formatRecipients(recipients) {
  if (!recipients) {
    return [];
  }
  
  // Handle string with comma or semicolon separators
  if (typeof recipients === 'string') {
    recipients = recipients.split(/[,;]/).map(r => r.trim()).filter(Boolean);
  }
  
  // Ensure it's an array
  if (!Array.isArray(recipients)) {
    recipients = [recipients];
  }
  
  // Format each recipient
  return recipients.map(recipient => {
    // If already in the correct format
    if (typeof recipient === 'object' && recipient.emailAddress) {
      return recipient;
    }
    
    // Handle string in format "Name <email@example.com>"
    if (typeof recipient === 'string') {
      const match = recipient.match(/^(.*?)\s*<([^>]+)>$/);
      if (match) {
        return {
          emailAddress: {
            name: match[1].trim(),
            address: match[2].trim()
          }
        };
      }
      
      // Just an email address
      return {
        emailAddress: {
          address: recipient.trim()
        }
      };
    }
    
    // Handle object with name and email properties
    if (typeof recipient === 'object' && recipient.email) {
      return {
        emailAddress: {
          name: recipient.name || '',
          address: recipient.email
        }
      };
    }
    
    // Default case
    return {
      emailAddress: {
        address: String(recipient)
      }
    };
  });
}

module.exports = {
  sendEmailHandler,
  createDraftHandler,
  replyEmailHandler,
  forwardEmailHandler
};