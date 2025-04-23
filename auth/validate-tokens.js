const { validateTokenStorage, listUsers, getToken } = require('./token-manager');
const logger = require('../utils/logger');

/**
 * Main function to validate and debug tokens
 */
async function main() {
  try {
    logger.info('=== TOKEN VALIDATION UTILITY ===');
    
    // First validate the token storage structure
    logger.info('Step 1: Validating token storage structure');
    const isValid = await validateTokenStorage();
    
    if (!isValid) {
      logger.error('Token storage validation failed. A new empty storage has been created.');
      logger.info('Please re-authenticate to create new tokens.');
      return;
    }
    
    // List all users with tokens
    logger.info('Step 2: Checking all authenticated users');
    const users = await listUsers();
    
    if (users.length === 0) {
      logger.info('No authenticated users found. Please authenticate first.');
      return;
    }
    
    logger.info(`Found ${users.length} authenticated users: ${users.join(', ')}`);
    
    // Check each user's token
    logger.info('Step 3: Checking token details for each user');
    for (const userId of users) {
      logger.info(`Checking token for user: ${userId}`);
      
      const tokenData = await getToken(userId);
      
      if (!tokenData) {
        logger.error(`Failed to retrieve valid token for user ${userId}`);
        continue;
      }
      
      // Display token info without revealing sensitive data
      logger.info(`Token type: ${tokenData.token_type || 'unknown'}`);
      logger.info(`Scopes: ${tokenData.scope || 'none specified'}`);
      logger.info(`Expires at: ${new Date(tokenData.expires_at).toLocaleString()}`);
      logger.info(`Has refresh token: ${!!tokenData.refresh_token}`);
      
      // Check if scopes are sufficient for Mail operations
      const hasMailRead = tokenData.scope && tokenData.scope.includes('Mail.Read');
      const hasMailReadWrite = tokenData.scope && tokenData.scope.includes('Mail.ReadWrite');
      const hasMailSend = tokenData.scope && tokenData.scope.includes('Mail.Send');
      
      if (!hasMailRead) {
        logger.error(`Token for user ${userId} is missing Mail.Read scope which is required for listing emails`);
      }
      
      if (!hasMailReadWrite) {
        logger.warn(`Token for user ${userId} is missing Mail.ReadWrite scope which may be required for some operations`);
      }
      
      if (!hasMailSend) {
        logger.warn(`Token for user ${userId} is missing Mail.Send scope which is required for sending emails`);
      }
    }
    
    logger.info('Token validation complete.');
    logger.info('If any issues were found, please re-authenticate to refresh the tokens with proper scopes.');
    
  } catch (error) {
    logger.error(`Token validation failed: ${error.message}`);
    logger.error(error.stack);
  }
}

// Execute the main function
main().catch(error => {
  logger.error(`Unhandled error: ${error.message}`);
  logger.error(error.stack);
  process.exit(1);
}); 