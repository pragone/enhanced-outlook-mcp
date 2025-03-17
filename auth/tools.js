const axios = require('axios');
const config = require('../config');
const logger = require('../utils/logger');
const { saveToken, deleteToken, listUsers } = require('./token-manager');

/**
 * Tool handler for authenticating with Microsoft Graph API
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Authentication result
 */
async function authenticateHandler(params = {}) {
  const authServerUrl = `http://localhost:${config.server.authPort}`;
  
  try {
    logger.info('Starting authentication flow');
    
    // Prepare authentication request
    const authRequest = {
      clientId: config.microsoft.clientId,
      scopes: params.scopes || config.microsoft.scopes,
      redirectUri: config.microsoft.redirectUri,
      state: params.userId || 'default'
    };
    
    // Request authentication URL from auth server
    const response = await axios.post(`${authServerUrl}/auth/start`, authRequest);
    
    // Check if authentication was initiated successfully
    if (response.data && response.data.status === 'authentication_started') {
      // Poll for authentication status
      const maxRetries = 60; // 5 minutes (assuming 5s interval)
      const pollInterval = 5000; // 5 seconds
      
      logger.info(`Authentication URL created successfully, waiting for completion: ${response.data.authUrl}`);
      
      // Return the auth URL to the user
      return {
        status: 'authentication_started',
        message: 'Authentication started. Please complete the authentication in your browser.',
        authUrl: response.data.authUrl,
        userId: params.userId || 'default',
        instruction: 'Please complete the authentication process in your browser. You will be redirected back once authenticated.'
      };
    } else {
      throw new Error('Failed to start authentication process');
    }
  } catch (error) {
    logger.error(`Authentication error: ${error.message}`);
    
    // Return error response
    return {
      status: 'error',
      message: `Authentication failed: ${error.message}`,
      instruction: 'Please try again. If the problem persists, check server logs for more details.'
    };
  }
}

/**
 * Tool handler for checking authentication status
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Authentication status
 */
async function checkAuthStatusHandler(params = {}) {
  const authServerUrl = `http://localhost:${config.server.authPort}`;
  
  try {
    // Request authentication status from auth server
    const response = await axios.get(`${authServerUrl}/auth/status`);
    
    return {
      ...response.data,
      // Add clear instructions based on status
      instruction: response.data.isAuthenticating 
        ? 'Authentication in progress. Please complete the authentication in your browser.'
        : response.data.userId
          ? 'You are authenticated. You can now use other tools that require authentication.'
          : 'Not authenticated. Please use the authenticate tool to start the authentication process.'
    };
  } catch (error) {
    logger.error(`Check auth status error: ${error.message}`);
    
    // Return error response
    return {
      status: 'error',
      message: `Failed to check authentication status: ${error.message}`,
      isAuthenticating: false,
      userId: null
    };
  }
}

/**
 * Tool handler for revoking authentication
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Revocation result
 */
async function revokeAuthenticationHandler(params = {}) {
  try {
    const userId = params.userId || 'default';
    
    // Delete token for user
    const wasDeleted = await deleteToken(userId);
    
    if (wasDeleted) {
      logger.info(`Authentication revoked for user ${userId}`);
      
      return {
        status: 'success',
        message: `Authentication revoked successfully for user ${userId}`,
        instruction: 'You will need to authenticate again to use tools that require authentication.'
      };
    } else {
      return {
        status: 'warning',
        message: `No authentication found for user ${userId}`,
        instruction: 'No action was needed as you were not authenticated.'
      };
    }
  } catch (error) {
    logger.error(`Revoke authentication error: ${error.message}`);
    
    // Return error response
    return {
      status: 'error',
      message: `Failed to revoke authentication: ${error.message}`
    };
  }
}

/**
 * Tool handler for listing authenticated users
 * @returns {Promise<Object>} - List of authenticated users
 */
async function listAuthenticatedUsersHandler() {
  try {
    // Get list of users with stored tokens
    const users = await listUsers();
    
    return {
      status: 'success',
      users,
      count: users.length,
      instruction: users.length > 0
        ? 'These are the currently authenticated users. You can specify userId when using other tools to act on behalf of a specific user.'
        : 'No authenticated users found. Please use the authenticate tool to authenticate.'
    };
  } catch (error) {
    logger.error(`List authenticated users error: ${error.message}`);
    
    // Return error response
    return {
      status: 'error',
      message: `Failed to list authenticated users: ${error.message}`,
      users: []
    };
  }
}

module.exports = {
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler
};