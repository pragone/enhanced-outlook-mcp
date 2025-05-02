const config = require('../config');
const logger = require('../utils/logger');
const { email: emailApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const { normalizeParameters } = require('../utils/parameter-helpers');
const auth = require('../auth/index');

/**
 * Send a new email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Send result
 */
async function sendEmailHandler(params = {}) {
  try {
    const { to, subject, body, contentType, cc, bcc, attachments } = params;
    
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
    
    // Check required parameters
    if (!subject) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email subject is required'
          })
        }]
      };
    }
    
    if (!body) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Email body is required'
          })
        }]
      };
    }
    
    // At least one recipient is required (to, cc, or bcc)
    if (!to && !cc && !bcc) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'At least one recipient (to, cc, or bcc) is required'
          })
        }]
      };
    }
    
    logger.info(`Sending email for user ${userId} with subject: ${subject}`);
    
    // Prepare recipients
    const toRecipients = formatRecipients(to);
    const ccRecipients = formatRecipients(cc);
    const bccRecipients = formatRecipients(bcc);
    
    // Prepare email message
    const message = {
      subject: subject,
      body: {
        contentType: contentType || 'HTML',
        content: body
      },
      toRecipients,
      ccRecipients,
      bccRecipients
    };
    
    // Set importance if provided
    if (params.importance) {
      message.importance = params.importance.toUpperCase();
    }
    
    // Send email using emailApi.sendMessage
    // This will correctly handle authentication and token reuse
    await emailApi.sendMessage(userId, message, {
      saveToSentItems: params.saveToSentItems !== false
    });
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Email sent successfully',
          subject: subject,
          recipientCount: {
            to: toRecipients.length,
            cc: ccRecipients.length,
            bcc: bccRecipients.length
          }
        })
      }]
    };
  } catch (error) {
    logger.error(`Error sending email: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to send email: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Create a draft email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Draft creation result
 */
async function createDraftHandler(params = {}) {
  let userId = params.userId;
  if (!userId) {
    const users = await listUsers();
    if (users.length === 0) {
      return formatMcpResponse({
        status: 'error',
        message: 'No authenticated users found. Please authenticate first.'
      });
    }
    userId = users.length === 1 ? users[0] : params.userId;
    if (!userId) {
      return formatMcpResponse({
        status: 'error',
        message: 'Multiple users found. Please specify userId parameter.'
      });
    }
  }
  
  try {
    logger.info(`Creating email draft for user ${userId}`);
    
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
    
    // Create draft using emailApi.createDraft
    // This will correctly handle authentication and token reuse
    const draftEmail = await emailApi.createDraft(userId, message);
    
    return formatMcpResponse({
      status: 'success',
      message: 'Draft email created successfully',
      emailId: draftEmail.id,
      subject: draftEmail.subject
    });
  } catch (error) {
    logger.error(`Error creating draft email: ${error.message}`);
    
    return formatMcpResponse({
      status: 'error',
      message: `Failed to create draft email: ${error.message}`
    });
  }
}

/**
 * Reply to an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Reply result
 */
async function replyEmailHandler(params = {}) {
  let userId = params.userId;
  if (!userId) {
    const users = await listUsers();
    if (users.length === 0) {
      return formatMcpResponse({
        status: 'error',
        message: 'No authenticated users found. Please authenticate first.'
      });
    }
    userId = users.length === 1 ? users[0] : params.userId;
    if (!userId) {
      return formatMcpResponse({
        status: 'error',
        message: 'Multiple users found. Please specify userId parameter.'
      });
    }
  }
  const messageId = params.messageId || params.emailId;
  
  if (!messageId) {
    return formatMcpResponse({
      status: 'error',
      message: 'Message ID is required'
    });
  }
  
  if (!params.message && !params.body) {
    return formatMcpResponse({
      status: 'error',
      message: 'Reply message or body is required'
    });
  }
  
  try {
    logger.info(`Replying to email ${messageId} for user ${userId}`);
    
    // Prepare the reply message or comment
    let replyMessage = null;
    let comment = null;
    
    if (params.message) {
      // Use full message object if provided
      replyMessage = params.message;
    } else {
      // Or just use the comment parameter
      comment = params.body;
    }
    
    // Determine if it's a reply or reply all
    const replyAll = params.replyAll === true;
    
    // Send the reply
    // Use emailApi.replyToMessage which will handle auth properly
    await emailApi.replyToMessage(userId, messageId, replyMessage, {
      replyAll,
      comment
    });
    
    return formatMcpResponse({
      status: 'success',
      message: `Email ${replyAll ? 'replied all' : 'replied'} successfully`,
      messageId: messageId
    });
  } catch (error) {
    logger.error(`Error replying to email: ${error.message}`);
    
    return formatMcpResponse({
      status: 'error',
      message: `Failed to reply to email: ${error.message}`
    });
  }
}

/**
 * Forward an email
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Forward result
 */
async function forwardEmailHandler(params = {}) {
  let userId = params.userId;
  if (!userId) {
    const users = await listUsers();
    if (users.length === 0) {
      return formatMcpResponse({
        status: 'error',
        message: 'No authenticated users found. Please authenticate first.'
      });
    }
    userId = users.length === 1 ? users[0] : params.userId;
    if (!userId) {
      return formatMcpResponse({
        status: 'error',
        message: 'Multiple users found. Please specify userId parameter.'
      });
    }
  }
  const messageId = params.messageId || params.emailId;
  const to = params.to;
  
  if (!messageId) {
    return formatMcpResponse({
      status: 'error',
      message: 'Message ID is required'
    });
  }
  
  if (!to) {
    return formatMcpResponse({
      status: 'error',
      message: 'At least one recipient (to) is required'
    });
  }
  
  try {
    logger.info(`Forwarding email ${messageId} for user ${userId}`);
    
    // Format recipients
    const toRecipients = formatRecipients(to);
    
    // Forward the email
    // Use emailApi.forwardMessage which will handle auth properly
    await emailApi.forwardMessage(userId, messageId, toRecipients, {
      comment: params.comment || params.body
    });
    
    return formatMcpResponse({
      status: 'success',
      message: 'Email forwarded successfully',
      messageId: messageId,
      recipientCount: toRecipients.length
    });
  } catch (error) {
    logger.error(`Error forwarding email: ${error.message}`);
    
    return formatMcpResponse({
      status: 'error',
      message: `Failed to forward email: ${error.message}`
    });
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

/**
 * Format response for MCP
 * @param {Object} data - Response data
 * @returns {Object} - MCP formatted response
 */
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

module.exports = {
  sendEmailHandler,
  createDraftHandler,
  replyEmailHandler,
  forwardEmailHandler
};