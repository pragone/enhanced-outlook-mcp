require('dotenv').config();
const { getAuthService } = require('./auth-service');
const logger = require('../utils/logger');

/**
 * Test the new authentication service with web-based auth flow
 */
async function testNewAuth() {
  // Set a global timeout for the entire test
  const testTimeout = setTimeout(() => {
    console.log("\n\n========================================");
    console.log("TEST TIMEOUT: The test has been running for too long.");
    console.log("If you see this, the browser auth flow might be waiting for your interaction");
    console.log("or there could be an issue with the authentication process.");
    console.log("========================================\n");
    process.exit(1);
  }, 60000); // 60 second timeout
  
  try {
    console.log("\n========================================");
    console.log("STARTING AUTHENTICATION TEST");
    console.log("========================================\n");
    
    logger.info('Starting new auth service test');
    console.log("1. Creating and initializing auth service...");
    
    // Create and initialize auth service
    const authService = getAuthService();
    await authService.initialize();
    
    // Check if authenticated
    console.log("2. Checking if already authenticated...");
    const isAuthenticated = await authService.isAuthenticated();
    logger.info(`Is already authenticated: ${isAuthenticated}`);
    console.log(`   Result: ${isAuthenticated ? "Already authenticated" : "Not authenticated"}`);
    
    if (isAuthenticated) {
      console.log("3. Getting user profile with cached credentials...");
      // Get Graph client and make a test call
      const client = await authService.getGraphClient();
      const userResponse = await client.api('/me').get();
      
      logger.info('User profile retrieved:');
      logger.info(`- Display Name: ${userResponse.displayName}`);
      logger.info(`- Email: ${userResponse.mail || userResponse.userPrincipalName}`);
      
      console.log(`   Success! Logged in as: ${userResponse.displayName} (${userResponse.mail || userResponse.userPrincipalName})`);
      logger.info('Silent authentication test completed successfully');
    } else {
      console.log("3. Starting interactive browser authentication...");
      console.log("   NOTE: This will open a browser window. Please complete the authentication there.");
      logger.info('Not authenticated. Starting interactive authentication...');
      
      // Start web-based authentication
      try {
        console.log("   Launching browser auth flow...");
        const authResult = await authService.authenticate();
        
        // Get access token
        console.log("4. Authentication successful, getting user info...");
        const token = authResult.accessToken;
        logger.info(`Access token acquired: ${token.substring(0, 10)}...`);
        
        // Get Graph client and make a test call
        const client = await authService.getGraphClient();
        const userResponse = await client.api('/me').get();
        
        logger.info('User profile retrieved:');
        logger.info(`- Display Name: ${userResponse.displayName}`);
        logger.info(`- Email: ${userResponse.mail || userResponse.userPrincipalName}`);
        
        console.log(`   Success! Logged in as: ${userResponse.displayName} (${userResponse.mail || userResponse.userPrincipalName})`);
        logger.info('Interactive authentication test completed successfully');
      } catch (authError) {
        console.log(`   ERROR: Authentication failed: ${authError.message}`);
        logger.error(`Interactive authentication failed: ${authError.message}`);
        if (authError.stack) {
          logger.error(authError.stack);
        }
      }
    }
    
    // Clean up resources before exiting
    console.log("5. Cleaning up resources...");
    await authService.cleanup();
    console.log("   Cleanup complete.");
    
    console.log("\n========================================");
    console.log("TEST COMPLETED SUCCESSFULLY");
    console.log("========================================\n");
  } catch (error) {
    console.log(`\nERROR: ${error.message}`);
    logger.error(`Authentication test failed: ${error.message}`);
    if (error.stack) {
      logger.error(error.stack);
    }
    
    // Make sure to clean up even on error
    try {
      console.log("Attempting cleanup after error...");
      const authService = getAuthService();
      await authService.cleanup();
      console.log("Cleanup complete.");
    } catch (cleanupError) {
      console.log(`Cleanup error: ${cleanupError.message}`);
      logger.error(`Error during cleanup: ${cleanupError.message}`);
    }
    
    console.log("\n========================================");
    console.log("TEST FAILED");
    console.log("========================================\n");
  } finally {
    // Clear the timeout
    clearTimeout(testTimeout);
  }
}

// Run the test
testNewAuth().catch(error => {
  console.error('Unhandled error:', error);
  process.exit(1);
}); 