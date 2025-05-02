const { 
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  getGraphClientForMCP
} = require('./tools-api');

const {
  AuthService,
  getAuthService
} = require('./auth-service');

const logger = require('../utils/logger');

/**
 * Enhanced Outlook MCP Authentication
 * 
 * This module provides a streamlined authentication approach for Outlook MCP,
 * using browser-based authentication with MSAL and integrating directly with
 * the Microsoft Graph API.
 * 
 * Features:
 * - Self-contained authentication flow (no external auth server)
 * - Browser-based user authentication
 * - Token caching and silent refresh
 * - Direct Microsoft Graph integration
 */

// Export all handlers and utility functions from the new auth implementation
module.exports = {
  // Auth tool handlers
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  
  // Auth service
  AuthService,
  getAuthService,
  getGraphClient: getGraphClientForMCP,
  
  // Tool handler for listing authenticated users
  listAuthenticatedUsersHandler: async (params) => {
    const result = await checkAuthStatusHandler();
    const results = [];
    
    if (result?.content?.[0]?.text) {
      const data = JSON.parse(result.content[0].text);
      if (data.status === 'authenticated' && data.user) {
        results.push(data.user);
      }
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          users: results
        })
      }]
    };
  },
  
  // MCP service methods that match the old API for compatibility
  refreshTokenHandler: async (userId) => {
    try {
      const authService = getAuthService();
      await authService.initialize();
      const token = await authService.getAccessToken(true); // Force refresh
      
      if (token) {
        logger.info(`Successfully refreshed token for user ${userId || 'unknown'}`);
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'success',
              message: 'Token refreshed successfully',
            })
          }]
        };
      }
      
      throw new Error('Failed to refresh token');
    } catch (error) {
      logger.error(`Error refreshing token: ${error.message}`);
      
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: `Failed to refresh token: ${error.message}`,
          })
        }]
      };
    }
  },
  
  tokenInfoHandler: async (userId) => {
    try {
      const authService = getAuthService();
      await authService.initialize();
      if (await authService.isAuthenticated()) {
        const client = await authService.getGraphClient();
        const user = await client.api('/me').select('displayName,mail,userPrincipalName').get();
        return {
          email: user.mail || user.userPrincipalName,
          displayName: user.displayName
        };
      }
      return null;
    } catch (error) {
      logger.error(`Error getting token info: ${error.message}`);
      return null;
    }
  }
}; 