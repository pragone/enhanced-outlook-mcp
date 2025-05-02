const { getAuthService } = require('./auth-service');
const logger = require('../utils/logger');

/**
 * Tool handler for authenticating with Microsoft Graph API
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Authentication result
 */
async function authenticateHandler(params = {}) {
  try {
    logger.info('Starting web-based authentication flow');
    
    // Get auth service and initialize
    const authService = getAuthService();
    await authService.initialize();
    
    // Check if already authenticated
    const isAuthenticated = await authService.isAuthenticated();
    
    if (isAuthenticated) {
      logger.info('User is already authenticated');
      
      try {
        // Get user info
        const client = await authService.getGraphClient();
        const user = await client.api('/me').select('displayName,mail,userPrincipalName').get();
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'authenticated',
              message: `You are already authenticated as ${user.displayName} (${user.mail || user.userPrincipalName}).`,
              user: {
                displayName: user.displayName,
                email: user.mail || user.userPrincipalName
              },
              instruction: 'You can continue using Microsoft Graph API tools.'
            })
          }]
        };
      } catch (graphError) {
        logger.error(`Failed to get user info despite being authenticated: ${graphError.message}`);
        // Token may be invalid, proceed with new authentication
      }
    }
    
    // Start interactive authentication
    try {
      logger.info('Starting interactive authentication flow');
      const authCodeUrl = await authService.getAuthUrl();
      
      if (!authCodeUrl) {
        throw new Error('Failed to generate authentication URL');
      }
      
      // Return the auth URL before completing the authentication
      // The user will need to open this URL in a browser
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'pending',
            auth_url: authCodeUrl,
            message: 'Please authenticate in your browser using the provided URL.',
            instruction: 'After authenticating, run check_auth_status to verify your authentication status.'
          })
        }]
      };
      
    } catch (authError) {
      logger.error(`Authentication failed: ${authError.message}`);
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: `Authentication failed: ${authError.message}`,
            isAuthenticated: false,
            instruction: 'Please try again. If the problem persists, check server logs for more details.'
          })
        }]
      };
    }
  } catch (error) {
    logger.error(`Authentication error: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Authentication error: ${error.message}`,
          isAuthenticated: false,
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
  try {
    const authService = getAuthService();
    await authService.initialize();
    
    const isAuthenticated = await authService.isAuthenticated();
    
    if (isAuthenticated) {
      // Get user info if authenticated
      try {
        const client = await authService.getGraphClient();
        const user = await client.api('/me').select('displayName,mail,userPrincipalName').get();
        
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'authenticated',
              message: `You are authenticated as ${user.displayName} (${user.mail || user.userPrincipalName}).`,
              user: {
                displayName: user.displayName,
                email: user.mail || user.userPrincipalName
              },
              instruction: 'You can use all tools that require authentication.'
            })
          }]
        };
      } catch (graphError) {
        logger.error(`Failed to get user info: ${graphError.message}`);
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'partial',
              message: 'You are authenticated, but retrieving user details failed.',
              error: graphError.message,
              instruction: 'You can use tools that require authentication, but you may need to reauthenticate if problems persist.'
            })
          }]
        };
      }
    } else {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'unauthenticated',
            message: 'You are not authenticated.',
            instruction: 'Use the authenticate tool to start the authentication process.'
          })
        }]
      };
    }
  } catch (error) {
    logger.error(`Check auth status error: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to check authentication status: ${error.message}`,
          instruction: 'An error occurred. Please try again or check server logs for more details.'
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
    const authService = getAuthService();
    await authService.initialize();
    
    // Check if authenticated before attempting to sign out
    const isAuthenticated = await authService.isAuthenticated();
    
    if (!isAuthenticated) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'warning',
            message: 'No active authentication to revoke.',
            instruction: 'No action was needed as you were not authenticated.'
          })
        }]
      };
    }
    
    // Sign out and clean up resources
    const wasSignedOut = await authService.signOut();
    await authService.cleanup();
    
    if (wasSignedOut) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'success',
            message: 'Authentication revoked successfully.',
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
            message: 'Failed to revoke authentication completely.',
            instruction: 'Please try again or check server logs for more details.'
          })
        }]
      };
    }
  } catch (error) {
    logger.error(`Revoke authentication error: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to revoke authentication: ${error.message}`,
          instruction: 'An error occurred. Please try again or check server logs for more details.'
        })
      }]
    };
  }
}

// Export the handlers
module.exports = {
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  
  // Helper function to get the Graph client for use in other modules
  getGraphClientForMCP: async () => {
    const authService = getAuthService();
    await authService.initialize();
    return authService.getGraphClient();
  },
}; 