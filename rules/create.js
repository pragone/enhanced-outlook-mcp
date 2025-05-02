const config = require('../config');
const logger = require('../utils/logger');
const { rules: rulesApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const auth = require('../auth/index');

/**
 * Create a new mail rule
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Creation result
 */
async function createRuleHandler(params = {}) {
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
            message: 'Multiple users found. Please specify userId parameter.'
          })
        }]
      };
    }
  }
  
  if (!params.displayName) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'Rule display name is required'
        })
      }]
    };
  }
  
  // Validate that at least one condition is provided
  if (!params.conditions || Object.keys(params.conditions).length === 0) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'At least one condition is required for the rule'
        })
      }]
    };
  }
  
  // Validate that at least one action is provided
  if (!params.actions || Object.keys(params.actions).length === 0) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'At least one action is required for the rule'
        })
      }]
    };
  }
  
  try {
    logger.info(`Creating mail rule "${params.displayName}" for user ${userId}`);
    
    // Format conditions object
    const conditions = formatRuleConditions(params.conditions);
    
    // Format actions object
    const actions = formatRuleActions(params.actions);
    
    // Create rule data
    const ruleData = {
      displayName: params.displayName,
      sequence: params.sequence || 0,
      isEnabled: params.isEnabled !== false,
      conditions,
      actions
    };
    
    // Create the rule using rulesApi
    const rule = await rulesApi.createRule(userId, ruleData);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Rule created successfully',
          ruleId: rule.id,
          displayName: rule.displayName
        })
      }]
    };
  } catch (error) {
    logger.error(`Error creating mail rule: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to create mail rule: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Update an existing mail rule
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Update result
 */
async function updateRuleHandler(params = {}) {
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
            message: 'Multiple users found. Please specify userId parameter.'
          })
        }]
      };
    }
  }
  const ruleId = params.ruleId;
  
  if (!ruleId) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'Rule ID is required'
        })
      }]
    };
  }
  
  try {
    logger.info(`Updating mail rule ${ruleId} for user ${userId}`);
    
    // Prepare update data
    const updateData = {};
    
    if (params.displayName) {
      updateData.displayName = params.displayName;
    }
    
    if (params.sequence !== undefined) {
      updateData.sequence = params.sequence;
    }
    
    if (params.isEnabled !== undefined) {
      updateData.isEnabled = params.isEnabled;
    }
    
    if (params.conditions) {
      updateData.conditions = formatRuleConditions(params.conditions);
    }
    
    if (params.actions) {
      updateData.actions = formatRuleActions(params.actions);
    }
    
    // Update the rule using rulesApi
    await rulesApi.updateRule(userId, ruleId, updateData);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          message: 'Rule updated successfully',
          ruleId
        })
      }]
    };
  } catch (error) {
    logger.error(`Error updating mail rule: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to update mail rule: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Format rule conditions for API
 * @param {Object} conditions - Rule conditions from user
 * @returns {Object} - Formatted conditions for Graph API
 */
function formatRuleConditions(conditions) {
  const formattedConditions = {};
  
  // Process each condition type
  if (conditions.bodyContains) {
    formattedConditions.bodyContains = Array.isArray(conditions.bodyContains) 
      ? conditions.bodyContains 
      : [conditions.bodyContains];
  }
  
  if (conditions.bodyOrSubjectContains) {
    formattedConditions.bodyOrSubjectContains = Array.isArray(conditions.bodyOrSubjectContains) 
      ? conditions.bodyOrSubjectContains 
      : [conditions.bodyOrSubjectContains];
  }
  
  if (conditions.categories) {
    formattedConditions.categories = Array.isArray(conditions.categories) 
      ? conditions.categories 
      : [conditions.categories];
  }
  
  if (conditions.fromAddresses) {
    formattedConditions.fromAddresses = formatEmailAddresses(conditions.fromAddresses);
  }
  
  if (conditions.hasAttachments !== undefined) {
    formattedConditions.hasAttachments = conditions.hasAttachments;
  }
  
  if (conditions.headerContains) {
    formattedConditions.headerContains = Array.isArray(conditions.headerContains) 
      ? conditions.headerContains 
      : [conditions.headerContains];
  }
  
  if (conditions.importance) {
    formattedConditions.importance = conditions.importance;
  }
  
  if (conditions.isApprovalRequest !== undefined) {
    formattedConditions.isApprovalRequest = conditions.isApprovalRequest;
  }
  
  if (conditions.isAutomaticForward !== undefined) {
    formattedConditions.isAutomaticForward = conditions.isAutomaticForward;
  }
  
  if (conditions.isAutomaticReply !== undefined) {
    formattedConditions.isAutomaticReply = conditions.isAutomaticReply;
  }
  
  if (conditions.isEncrypted !== undefined) {
    formattedConditions.isEncrypted = conditions.isEncrypted;
  }
  
  if (conditions.isMeetingRequest !== undefined) {
    formattedConditions.isMeetingRequest = conditions.isMeetingRequest;
  }
  
  if (conditions.isMeetingResponse !== undefined) {
    formattedConditions.isMeetingResponse = conditions.isMeetingResponse;
  }
  
  if (conditions.isReadReceipt !== undefined) {
    formattedConditions.isReadReceipt = conditions.isReadReceipt;
  }
  
  if (conditions.messageActionFlag) {
    formattedConditions.messageActionFlag = conditions.messageActionFlag;
  }
  
  if (conditions.notSentToMe !== undefined) {
    formattedConditions.notSentToMe = conditions.notSentToMe;
  }
  
  if (conditions.recipientContains) {
    formattedConditions.recipientContains = Array.isArray(conditions.recipientContains) 
      ? conditions.recipientContains 
      : [conditions.recipientContains];
  }
  
  if (conditions.senderContains) {
    formattedConditions.senderContains = Array.isArray(conditions.senderContains) 
      ? conditions.senderContains 
      : [conditions.senderContains];
  }
  
  if (conditions.sensitivity) {
    formattedConditions.sensitivity = conditions.sensitivity;
  }
  
  if (conditions.sentCcMe !== undefined) {
    formattedConditions.sentCcMe = conditions.sentCcMe;
  }
  
  if (conditions.sentOnlyToMe !== undefined) {
    formattedConditions.sentOnlyToMe = conditions.sentOnlyToMe;
  }
  
  if (conditions.sentToAddresses) {
    formattedConditions.sentToAddresses = formatEmailAddresses(conditions.sentToAddresses);
  }
  
  if (conditions.sentToMe !== undefined) {
    formattedConditions.sentToMe = conditions.sentToMe;
  }
  
  if (conditions.sentToOrCcMe !== undefined) {
    formattedConditions.sentToOrCcMe = conditions.sentToOrCcMe;
  }
  
  if (conditions.subjectContains) {
    formattedConditions.subjectContains = Array.isArray(conditions.subjectContains) 
      ? conditions.subjectContains 
      : [conditions.subjectContains];
  }
  
  if (conditions.withinSizeRange) {
    formattedConditions.withinSizeRange = conditions.withinSizeRange;
  }
  
  return formattedConditions;
}

/**
 * Format rule actions for API
 * @param {Object} actions - Rule actions from user
 * @returns {Object} - Formatted actions for Graph API
 */
function formatRuleActions(actions) {
  const formattedActions = {};
  
  // Process each action type
  if (actions.assignCategories) {
    formattedActions.assignCategories = Array.isArray(actions.assignCategories) 
      ? actions.assignCategories 
      : [actions.assignCategories];
  }
  
  if (actions.copyToFolder) {
    formattedActions.copyToFolder = actions.copyToFolder;
  }
  
  if (actions.delete !== undefined) {
    formattedActions.delete = actions.delete;
  }
  
  if (actions.forwardAsAttachmentTo) {
    formattedActions.forwardAsAttachmentTo = formatEmailAddresses(actions.forwardAsAttachmentTo);
  }
  
  if (actions.forwardTo) {
    formattedActions.forwardTo = formatEmailAddresses(actions.forwardTo);
  }
  
  if (actions.markAsRead !== undefined) {
    formattedActions.markAsRead = actions.markAsRead;
  }
  
  if (actions.markImportance) {
    formattedActions.markImportance = actions.markImportance;
  }
  
  if (actions.moveToFolder) {
    formattedActions.moveToFolder = actions.moveToFolder;
  }
  
  if (actions.permanentDelete !== undefined) {
    formattedActions.permanentDelete = actions.permanentDelete;
  }
  
  if (actions.redirectTo) {
    formattedActions.redirectTo = formatEmailAddresses(actions.redirectTo);
  }
  
  if (actions.stopProcessingRules !== undefined) {
    formattedActions.stopProcessingRules = actions.stopProcessingRules;
  }
  
  return formattedActions;
}

/**
 * Format email addresses for API
 * @param {Array|string|Object} addresses - Email addresses in various formats
 * @returns {Array} - Formatted email addresses
 */
function formatEmailAddresses(addresses) {
  if (!addresses) {
    return [];
  }
  
  // Handle string with comma or semicolon separators
  if (typeof addresses === 'string') {
    addresses = addresses.split(/[,;]/).map(a => a.trim()).filter(Boolean);
  }
  
  // Ensure it's an array
  if (!Array.isArray(addresses)) {
    addresses = [addresses];
  }
  
  // Format each address
  return addresses.map(address => {
    // If already in the correct format
    if (typeof address === 'object' && address.emailAddress) {
      return address;
    }
    
    // Handle string in format "Name <email@example.com>"
    if (typeof address === 'string') {
      const match = address.match(/^(.*?)\s*<([^>]+)>$/);
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
          address: address.trim()
        }
      };
    }
    
    // Handle object with name and email properties
    if (typeof address === 'object' && address.email) {
      return {
        emailAddress: {
          name: address.name || '',
          address: address.email
        }
      };
    }
    
    // Default case
    return {
      emailAddress: {
        address: String(address)
      }
    };
  });
}

module.exports = {
  createRuleHandler,
  updateRuleHandler
};