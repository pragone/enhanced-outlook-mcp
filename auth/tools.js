const axios = require('axios');
const config = require('../config');
const logger = require('../utils/logger');
const { saveToken, deleteToken, listUsers } = require('./token-manager');
const { normalizeParameters, lookupDefaultUser } = require('../utils/parameter-helpers');

/**
 * Tool handler for authenticating with Microsoft Graph API
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Authentication result
 */
async function authenticateHandler(params = {}) {
  const authServerUrl = `http://localhost:${config.server.authPort}`;
  
  // Debug logging of the full configuration
  logger.info('==== AUTHENTICATION DEBUG START ====');
  logger.info(`Authentication server URL: ${authServerUrl}`);
  logger.info(`Client ID: ${config.microsoft.clientId}`);
  logger.info(`Client Secret present: ${!!process.env.MS_CLIENT_SECRET}`);
  logger.info(`Redirect URI: ${config.microsoft.redirectUri}`);
  logger.info(`Authority: ${config.microsoft.authority}`);
  logger.info(`Scopes: ${JSON.stringify(config.microsoft.scopes)}`);
  
  try {
    logger.info('Starting authentication flow');
    logger.info(`Authentication parameters: clientId: ${config.microsoft.clientId}, redirectUri: ${config.microsoft.redirectUri}`);
    
    // Use the full set of scopes from configuration
    const authRequest = {
      clientId: config.microsoft.clientId,
      scopes: config.microsoft.scopes,
      redirectUri: config.microsoft.redirectUri,
      state: 'default'
    };
    
    logger.info(`Sending auth request to ${authServerUrl}/auth/start`);
    logger.info(`Full auth request: ${JSON.stringify(authRequest)}`);
    
    // Request authentication URL from auth server
    try {
      logger.info('Making axios post request to auth server...');
      const response = await axios.post(`${authServerUrl}/auth/start`, authRequest);
      logger.info(`Auth server response received: ${JSON.stringify(response.data)}`);
      
      // Check if authentication was initiated successfully
      if (response.data && response.data.status === 'authentication_started') {
        logger.info(`Authentication URL created successfully: ${response.data.authUrl}`);
        
        // Return the auth URL to the user
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'authentication_started',
              message: 'Authentication started. Please complete the authentication in your browser.',
              authUrl: response.data.authUrl,
              userId: 'default',
              instruction: 'Please complete the authentication process in your browser. You will be redirected back once authenticated.'
            })
          }]
        };
      } else {
        logger.info(`Unexpected response from auth server: ${JSON.stringify(response.data)}`);
        throw new Error('Failed to start authentication process');
      }
    } catch (axiosError) {
      logger.error('Auth server communication error:');
      if (axiosError.response) {
        logger.error(`Status: ${axiosError.response.status}`);
        logger.error(`Headers: ${JSON.stringify(axiosError.response.headers)}`);
        logger.error(`Data: ${JSON.stringify(axiosError.response.data)}`);
      } else if (axiosError.request) {
        logger.error(`No response received. Request: ${JSON.stringify(axiosError.request._header || 'no request header')}`);
      } else {
        logger.error(`Error message: ${axiosError.message}`);
        logger.error(`Error stack: ${axiosError.stack}`);
      }
      throw axiosError;
    }
  } catch (error) {
    logger.error(`Authentication error: ${error.message}`);
    logger.info('==== AUTHENTICATION DEBUG END ====');
    
    // Return error response
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Authentication failed: ${error.message}`,
          instruction: 'Please try again. If the problem persists, check server logs for more details.'
        })
      }]
    };
  }
}

/**
 * Tool handler for checking authentication status
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Authentication status
 */
async function checkAuthStatusHandler(params = {}) {
  const normalizedParams = normalizeParameters(params);
  const authServerUrl = `http://localhost:${config.server.authPort}`;
  
  try {
    // Request authentication status from auth server
    const response = await axios.get(`${authServerUrl}/auth/status`);
    
    const responseData = { ...response.data };
    
    // For Claude Desktop compatibility, support 'default' user ID
    if (normalizedParams.userId === 'default') {
      // Get the first actual user if available
      const actualUserId = await lookupDefaultUser();
      
      if (actualUserId) {
        logger.info(`Mapping 'default' userId to actual user: ${actualUserId}`);
        responseData.userId = actualUserId;
        responseData.authenticated = true;
      }
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          ...responseData,
          // Add clear instructions based on status
          instruction: responseData.isAuthenticating 
            ? 'Authentication in progress. Please complete the authentication in your browser.'
            : responseData.userId
              ? 'You are authenticated. You can now use other tools that require authentication.'
              : 'Not authenticated. Please use the authenticate tool to start the authentication process.'
        })
      }]
    };
  } catch (error) {
    logger.error(`Check auth status error: ${error.message}`);
    
    // Return error response
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to check authentication status: ${error.message}`,
          isAuthenticating: false,
          userId: null
        })
      }]
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
    let userId = params.userId;
    if (!userId) {
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Nothing to revoke.'
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
              message: 'Multiple users found. Please specify userId parameter to indicate which authentication to revoke.'
            })
          }]
        };
      }
    }
    
    // Delete token for user
    const wasDeleted = await deleteToken(userId);
    
    if (wasDeleted) {
      logger.info(`Authentication revoked for user ${userId}`);
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'success',
            message: `Authentication revoked successfully for user ${userId}`,
            instruction: 'You will need to authenticate again to use tools that require authentication.'
          })
        }]
      };
    } else {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'warning',
            message: `No authentication found for user ${userId}`,
            instruction: 'No action was needed as you were not authenticated.'
          })
        }]
      };
    }
  } catch (error) {
    logger.error(`Revoke authentication error: ${error.message}`);
    
    // Return error response
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to revoke authentication: ${error.message}`
        })
      }]
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
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          users,
          count: users.length,
          instruction: users.length > 0
            ? 'These are the currently authenticated users. You can specify userId when using other tools to act on behalf of a specific user.'
            : 'No authenticated users found. Please use the authenticate tool to authenticate.'
        })
      }]
    };
  } catch (error) {
    logger.error(`List authenticated users error: ${error.message}`);
    
    // Return error response
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list authenticated users: ${error.message}`,
          users: []
        })
      }]
    };
  }
}

module.exports = {
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler
};