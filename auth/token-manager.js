const fs = require('fs').promises;
const path = require('path');
const axios = require('axios');
const config = require('../config');
const logger = require('../utils/logger');

// Token storage path
const TOKEN_STORAGE_PATH = config.server.tokenStoragePath;

/**
 * Get token data from storage
 * @returns {Promise<Object>} - Token storage data
 */
async function getTokenStorage() {
  try {
    const data = await fs.readFile(TOKEN_STORAGE_PATH, 'utf8');
    return JSON.parse(data);
  } catch (error) {
    // If file doesn't exist or can't be parsed, return empty object
    if (error.code === 'ENOENT' || error.name === 'SyntaxError') {
      return {};
    }
    
    logger.error(`Error reading token storage: ${error.message}`);
    throw error;
  }
}

/**
 * Save token storage data
 * @param {Object} tokenData - Token storage data to save
 * @returns {Promise<void>}
 */
async function saveTokenStorage(tokenData) {
  try {
    // Ensure directory exists
    const dir = path.dirname(TOKEN_STORAGE_PATH);
    await fs.mkdir(dir, { recursive: true });
    
    // Write token data
    await fs.writeFile(
      TOKEN_STORAGE_PATH, 
      JSON.stringify(tokenData, null, 2),
      { encoding: 'utf8', mode: 0o600 } // File permissions: user read/write only
    );
  } catch (error) {
    logger.error(`Error saving token storage: ${error.message}`);
    throw error;
  }
}

/**
 * Get token for a specific user
 * @param {string} userId - User identifier
 * @returns {Promise<Object|null>} - Token data or null if not found
 */
async function getToken(userId) {
  if (!userId) {
    throw new Error('User ID is required');
  }
  
  const tokenStorage = await getTokenStorage();
  const tokenData = tokenStorage[userId];
  
  if (!tokenData) {
    logger.info(`No token found for user ${userId}`);
    return null;
  }
  
  // Debug log the token data without revealing the full token
  logger.info(`Token found for user ${userId} with scopes: ${tokenData.scope || 'none specified'}`);
  
  // Check if token has scopes based on the API being accessed (inferred from the module importing this)
  // Include all possible scopes to avoid refresh loops
  const requiredScopes = [];
  const callerStack = new Error().stack;
  
  // Check which module is calling this function and set appropriate required scopes
  if (callerStack.includes('/calendar/')) {
    requiredScopes.push('Calendars.ReadWrite');
  } else if (callerStack.includes('/mail/')) {
    requiredScopes.push('Mail.Read');
  }
  
  // Only check required scopes if we've determined that specific ones are needed
  const hasRequiredScopes = requiredScopes.length === 0 || 
    (tokenData.scope && requiredScopes.every(scope => tokenData.scope.split(' ').includes(scope)));
  
  if (!hasRequiredScopes) {
    logger.warn(`Token for user ${userId} is missing required scopes. Found: ${tokenData.scope || 'none'}`);
    logger.info('Will attempt to reauthenticate with proper scopes');
    
    try {
      if (tokenData.refresh_token) {
        logger.info(`Attempting to refresh token with proper scopes for user ${userId}`);
        const refreshedToken = await refreshToken(userId, tokenData.refresh_token);
        return refreshedToken;
      } else {
        logger.error('No refresh token available, cannot upgrade scopes');
        return null;
      }
    } catch (error) {
      logger.error(`Failed to refresh token with proper scopes: ${error.message}`);
      return null;
    }
  }
  
  // Check if token is expired or about to expire (within 5 minutes)
  const now = Date.now();
  const expiresAt = tokenData.expires_at || 0;
  
  if (now >= expiresAt - 5 * 60 * 1000) {
    logger.info(`Token for user ${userId} is expired or about to expire, attempting refresh`);
    
    // Attempt to refresh the token
    try {
      const refreshedToken = await refreshToken(userId, tokenData.refresh_token);
      return refreshedToken;
    } catch (error) {
      logger.error(`Failed to refresh token for user ${userId}: ${error.message}`);
      return null;
    }
  }
  
  return tokenData;
}

/**
 * Refresh an expired token
 * @param {string} userId - User identifier
 * @param {string} refreshToken - Refresh token
 * @returns {Promise<Object>} - New token data
 */
async function refreshToken(userId, refreshToken) {
  if (!refreshToken) {
    throw new Error('Refresh token is required');
  }
  
  try {
    logger.info(`Refreshing token for user ${userId}`);
    logger.info(`Using scopes: ${config.microsoft.scopes.join(' ')}`);
    
    // Make request to Microsoft identity platform
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams({
        client_id: config.microsoft.clientId,
        scope: config.microsoft.scopes.join(' '),
        refresh_token: refreshToken,
        grant_type: 'refresh_token'
      }),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      }
    );
    
    const tokenData = response.data;
    
    // Log the received scopes
    logger.info(`Token refreshed with scopes: ${tokenData.scope || 'none specified'}`);
    
    // Calculate expiration time
    const expiresIn = tokenData.expires_in || 3600;
    const expiresAt = Date.now() + expiresIn * 1000;
    
    // Create updated token data
    const updatedTokenData = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token || refreshToken, // Use new refresh token if provided
      id_token: tokenData.id_token,
      expires_in: expiresIn,
      expires_at: expiresAt,
      scope: tokenData.scope,
      token_type: tokenData.token_type || 'Bearer'
    };
    
    // Save updated token
    await saveToken(userId, updatedTokenData);
    
    return updatedTokenData;
  } catch (error) {
    logger.error(`Token refresh error: ${error.message}`);
    if (error.response) {
      logger.error(`Token refresh response: ${JSON.stringify(error.response.data)}`);
    }
    throw new Error(`Failed to refresh token: ${error.message}`);
  }
}

/**
 * Save token for a specific user
 * @param {string} userId - User identifier
 * @param {Object} tokenData - Token data to save
 * @returns {Promise<void>}
 */
async function saveToken(userId, tokenData) {
  if (!userId) {
    throw new Error('User ID is required');
  }
  
  if (!tokenData || !tokenData.access_token) {
    throw new Error('Valid token data is required');
  }
  
  // Calculate expiration time if not provided
  if (!tokenData.expires_at && tokenData.expires_in) {
    tokenData.expires_at = Date.now() + tokenData.expires_in * 1000;
  }
  
  // Get existing storage
  const tokenStorage = await getTokenStorage();
  
  // Update token for user
  tokenStorage[userId] = tokenData;
  
  // Save updated storage
  await saveTokenStorage(tokenStorage);
  
  logger.info(`Token saved for user ${userId}`);
}

/**
 * Delete token for a specific user
 * @param {string} userId - User identifier
 * @returns {Promise<boolean>} - True if token was deleted, false if not found
 */
async function deleteToken(userId) {
  if (!userId) {
    throw new Error('User ID is required');
  }
  
  // Get existing storage
  const tokenStorage = await getTokenStorage();
  
  // Check if user has a token
  if (!tokenStorage[userId]) {
    return false;
  }
  
  // Delete token
  delete tokenStorage[userId];
  
  // Save updated storage
  await saveTokenStorage(tokenStorage);
  
  logger.info(`Token deleted for user ${userId}`);
  
  return true;
}

/**
 * List all users with stored tokens
 * @returns {Promise<Array>} - Array of user IDs
 */
async function listUsers() {
  const tokenStorage = await getTokenStorage();
  return Object.keys(tokenStorage);
}

/**
 * Validate and fix token storage if needed
 * @returns {Promise<boolean>} - True if validated successfully
 */
async function validateTokenStorage() {
  try {
    logger.info('Validating token storage');
    
    // Read the current token storage
    const tokenStoragePath = TOKEN_STORAGE_PATH;
    let tokenStorage = {};
    
    try {
      const data = await fs.readFile(tokenStoragePath, 'utf8');
      try {
        tokenStorage = JSON.parse(data);
        logger.info(`Token storage parsed successfully. Contains ${Object.keys(tokenStorage).length} users.`);
      } catch (parseError) {
        logger.error(`Token storage file exists but contains invalid JSON: ${parseError.message}`);
        logger.info('Creating new token storage file');
        tokenStorage = {};
        await saveTokenStorage(tokenStorage);
        return false;
      }
    } catch (readError) {
      if (readError.code === 'ENOENT') {
        logger.info('Token storage file does not exist. Creating new file.');
        await saveTokenStorage({});
        return true;
      }
      throw readError;
    }
    
    // Validate the token structure
    for (const [userId, token] of Object.entries(tokenStorage)) {
      logger.info(`Validating token for user: ${userId}`);
      
      if (!token || typeof token !== 'object') {
        logger.error(`Invalid token data for user ${userId}: Not an object`);
        delete tokenStorage[userId];
        continue;
      }
      
      if (!token.access_token) {
        logger.error(`Invalid token data for user ${userId}: Missing access_token`);
        delete tokenStorage[userId];
        continue;
      }
      
      if (!token.refresh_token) {
        logger.warn(`Token data for user ${userId} is missing refresh_token. Token will need to be renewed manually.`);
      }
      
      if (!token.scope) {
        logger.warn(`Token data for user ${userId} is missing scope information.`);
      } else {
        logger.info(`Token for user ${userId} has scopes: ${token.scope}`);
      }
    }
    
    // Save cleaned-up token storage
    await saveTokenStorage(tokenStorage);
    logger.info('Token storage validation complete');
    
    return true;
  } catch (error) {
    logger.error(`Error validating token storage: ${error.message}`);
    return false;
  }
}

module.exports = {
  getToken,
  saveToken,
  refreshToken,
  deleteToken,
  listUsers,
  validateTokenStorage
};