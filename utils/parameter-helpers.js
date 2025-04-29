const logger = require('./logger');
const { listUsers } = require('../auth/token-manager');

/**
 * Looks up an actual userId when 'default' is specified
 * This helps with compatibility for clients that use 'default' as their userId
 * 
 * @returns {Promise<string|null>} The first available userId from storage or null if none
 */
async function lookupDefaultUser() {
  try {
    const users = await listUsers();
    if (users && users.length > 0) {
      logger.debug(`Using first available user '${users[0]}' for 'default' userId`);
      return users[0];
    }
    return null;
  } catch (error) {
    logger.error(`Error looking up default user: ${error.message}`);
    return null;
  }
}

/**
 * Normalizes userId from different parameter formats
 * Handles various naming conventions and formats from different clients including Claude Desktop
 * 
 * @param {Object} params - The raw parameters object
 * @param {boolean} allowDefault - Whether to use 'default' as fallback if no userId is provided
 * @returns {string|null} The normalized userId or null if not found
 */
function normalizeUserId(params = {}, allowDefault = true) {
  // Check for direct parameter names first
  const userId = params.userId || params.user_id || params.id;
  
  if (userId) {
    return userId;
  }
  
  // Check if this might be a Claude Desktop format with nested params
  if (params.arguments && typeof params.arguments === 'object') {
    const args = params.arguments;
    return args.userId || args.user_id || args.id;
  }
  
  // Check if the signal property exists (Claude Desktop indicator)
  if (params.signal && params.contextData) {
    const contextData = params.contextData;
    return contextData.userId || contextData.user_id || contextData.id;
  }
  
  // Check if using global last message format
  if (global.__last_message?.params?.arguments) {
    const args = global.__last_message.params.arguments;
    return args.userId || args.user_id || args.id;
  }

  // Return default if allowed
  return allowDefault ? 'default' : null;
}

/**
 * Normalize folder ID from different parameter formats
 * 
 * @param {Object} params - The raw parameters object
 * @param {string} defaultValue - Default value if not found
 * @returns {string} The normalized folderId
 */
function normalizeFolderId(params = {}, defaultValue = 'inbox') {
  return params.folderId || params.folder_id || getNestedParam(params, 'folderId') || defaultValue;
}

/**
 * Normalize email ID from different parameter formats
 * 
 * @param {Object} params - The raw parameters object
 * @returns {string|null} The normalized emailId or null if not found
 */
function normalizeEmailId(params = {}) {
  return params.emailId || params.email_id || params.id || params.messageId || params.message_id || getNestedParam(params, 'emailId');
}

/**
 * Try to get a parameter from nested objects
 * 
 * @param {Object} params - The parameters object
 * @param {string} paramName - The parameter name to look for
 * @returns {any|null} The parameter value or null if not found
 */
function getNestedParam(params, paramName) {
  // Check in arguments
  if (params.arguments && typeof params.arguments === 'object') {
    if (params.arguments[paramName]) {
      return params.arguments[paramName];
    }
  }
  
  // Check in contextData (Claude Desktop)
  if (params.contextData && typeof params.contextData === 'object') {
    if (params.contextData[paramName]) {
      return params.contextData[paramName];
    }
  }
  
  // Check in global last message
  if (global.__last_message?.params?.arguments) {
    if (global.__last_message.params.arguments[paramName]) {
      return global.__last_message.params.arguments[paramName];
    }
  }
  
  return null;
}

/**
 * Extract all normalized parameters from a request
 * Handles various parameter formats and naming conventions
 * 
 * @param {Object} params - The raw parameters object 
 * @param {Object} options - Options for normalization
 * @returns {Object} Normalized parameters object
 */
function normalizeParameters(params = {}, options = {}) {
  const normalized = {};
  
  // Add userId if available
  const userId = normalizeUserId(params, options.allowDefaultUserId !== false);
  if (userId) {
    normalized.userId = userId;
  }
  
  // Handle common parameters with their aliases
  const paramMappings = {
    folderId: ['folder_id', 'folder'],
    emailId: ['email_id', 'messageId', 'message_id'],
    calendarId: ['calendar_id', 'calendar'],
    eventId: ['event_id', 'event'],
    attachmentId: ['attachment_id'],
    ruleId: ['rule_id'],
    limit: ['count', 'max'],
    skip: ['offset'],
    query: ['search', 'filter'],
    startDateTime: ['start_date', 'start'],
    endDateTime: ['end_date', 'end']
  };
  
  // Process each parameter mapping
  for (const [normalName, aliases] of Object.entries(paramMappings)) {
    // Direct parameter
    if (params[normalName] !== undefined) {
      normalized[normalName] = params[normalName];
      continue;
    }
    
    // Check aliases
    for (const alias of aliases) {
      if (params[alias] !== undefined) {
        normalized[normalName] = params[alias];
        break;
      }
    }
    
    // Check nested objects if not found directly
    if (normalized[normalName] === undefined) {
      const nestedValue = getNestedParam(params, normalName);
      if (nestedValue !== null) {
        normalized[normalName] = nestedValue;
      } else {
        // Check aliases in nested objects
        for (const alias of aliases) {
          const aliasValue = getNestedParam(params, alias);
          if (aliasValue !== null) {
            normalized[normalName] = aliasValue;
            break;
          }
        }
      }
    }
  }
  
  // Copy any other parameters that weren't specifically mapped
  for (const key in params) {
    if (!normalized[key] && !['arguments', 'contextData', 'signal'].includes(key)) {
      normalized[key] = params[key];
    }
  }
  
  return normalized;
}

module.exports = {
  normalizeUserId,
  normalizeFolderId,
  normalizeEmailId,
  normalizeParameters,
  lookupDefaultUser
}; 