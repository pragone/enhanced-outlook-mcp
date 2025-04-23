const { authenticateHandler, checkAuthStatusHandler } = require('./auth/tools');
const logger = require('./utils/logger');
const config = require('./config');

// Set logger to debug level
logger.level = 'debug';

async function main() {
  console.log('=== Authentication Debug ===');
  console.log('Microsoft client ID:', config.microsoft.clientId);
  console.log('Auth server port:', config.server.authPort);
  console.log('Redirect URI:', config.microsoft.redirectUri);
  console.log('Token storage path:', config.server.tokenStoragePath);
  console.log('Scopes:', config.microsoft.scopes);
  
  try {
    console.log('\nStarting authentication flow...');
    console.log('Please follow the URL displayed in your browser to complete the authentication process.');
    console.log('After authenticating, you\'ll be redirected back to the application.\n');
    
    const result = await authenticateHandler();
    console.log('\nAuthentication initiated:', result);
    
    console.log('\nChecking authentication status...');
    // Wait a bit for auth to complete
    await new Promise(resolve => setTimeout(resolve, 5000));
    
    const statusResult = await checkAuthStatusHandler();
    console.log('Auth status result:', statusResult);
  } catch (error) {
    console.error('Authentication error:', error);
  }
}

main(); 