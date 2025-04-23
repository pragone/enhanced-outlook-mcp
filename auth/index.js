const { 
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler
} = require('./tools');
const { 
  listUsers, 
  getUserTokenData, 
  refreshAccessToken 
} = require('./token-manager');

// Export all handlers directly
module.exports = {
  // Auth tools
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler,
  
  // Token utilities
  listUsers,
  getUserTokenData,
  refreshAccessToken,
  refreshTokenHandler: refreshAccessToken,
  tokenInfoHandler: getUserTokenData
};