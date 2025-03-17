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
    return null;
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

module.exports = {
  getToken,
  saveToken,
  refreshToken,
  deleteToken,
  listUsers
};