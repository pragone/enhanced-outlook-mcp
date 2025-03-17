const express = require('express');
const bodyParser = require('body-parser');
const http = require('http');
const socketIo = require('socket.io');
const open = require('open');
const config = require('./config');
const logger = require('./utils/logger');
const { saveToken } = require('./auth/token-manager');

// Create Express app
const app = express();
app.use(bodyParser.json());

// Create HTTP server
const server = http.createServer(app);

// Set up Socket.IO for real-time status updates
const io = socketIo(server, {
  cors: {
    origin: '*',
    methods: ['GET', 'POST']
  }
});

// Initialize authentication state
let authState = {
  isAuthenticating: false,
  userId: null,
  error: null
};

// Socket.IO connection handler
io.on('connection', (socket) => {
  logger.info('Client connected to auth status updates');
  
  // Send current auth state to newly connected client
  socket.emit('authStatus', authState);
  
  // Handle disconnection
  socket.on('disconnect', () => {
    logger.info('Client disconnected from auth status updates');
  });
});

// Authentication callback endpoint
app.get('/auth/callback', async (req, res) => {
  const { code, state, error, error_description } = req.query;
  
  if (error) {
    logger.error(`Authentication error: ${error} - ${error_description}`);
    authState = {
      isAuthenticating: false,
      userId: null,
      error: `${error}: ${error_description}`
    };
    io.emit('authStatus', authState);
    
    return res.send(`
      <html>
        <head><title>Authentication Failed</title></head>
        <body>
          <h1>Authentication Failed</h1>
          <p>Error: ${error}</p>
          <p>Description: ${error_description}</p>
          <p>You can close this window now.</p>
        </body>
      </html>
    `);
  }
  
  try {
    // Exchange code for tokens using the token manager
    const tokenResult = await exchangeCodeForToken(code, state);
    
    // Update authentication state
    authState = {
      isAuthenticating: false,
      userId: tokenResult.userId,
      error: null
    };
    io.emit('authStatus', authState);
    
    return res.send(`
      <html>
        <head>
          <title>Authentication Successful</title>
          <style>
            body { font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; }
            .success { color: green; }
          </style>
        </head>
        <body>
          <h1 class="success">Authentication Successful</h1>
          <p>You have successfully authenticated with Microsoft Outlook.</p>
          <p>You can now close this window and return to Claude.</p>
        </body>
      </html>
    `);
  } catch (err) {
    logger.error('Token exchange error:', err);
    
    authState = {
      isAuthenticating: false,
      userId: null,
      error: `Token exchange failed: ${err.message}`
    };
    io.emit('authStatus', authState);
    
    return res.status(500).send(`
      <html>
        <head><title>Authentication Error</title></head>
        <body>
          <h1>Authentication Error</h1>
          <p>An error occurred while exchanging the authentication code for tokens.</p>
          <p>Error: ${err.message}</p>
          <p>Please try again.</p>
        </body>
      </html>
    `);
  }
});

// Initiate authentication endpoint
app.post('/auth/start', async (req, res) => {
  const { clientId, scopes, redirectUri, state = 'default' } = req.body;
  
  // Validate required parameters
  if (!clientId || !scopes || !redirectUri) {
    return res.status(400).json({
      error: 'Missing required authentication parameters'
    });
  }
  
  // Update authentication state
  authState = {
    isAuthenticating: true,
    userId: null,
    error: null
  };
  io.emit('authStatus', authState);
  
  // Construct the Microsoft OAuth URL
  const scopesParam = encodeURIComponent(scopes.join(' '));
  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${scopesParam}&response_mode=query&state=${state}`;
  
  logger.info(`Starting authentication flow with state: ${state}`);
  
  // Open the browser for authentication
  try {
    await open(authUrl);
    res.json({ status: 'authentication_started', authUrl });
  } catch (err) {
    logger.error('Failed to open browser:', err);
    
    authState = {
      isAuthenticating: false,
      userId: null,
      error: `Failed to open browser: ${err.message}`
    };
    io.emit('authStatus', authState);
    
    res.status(500).json({
      error: 'Failed to open browser for authentication',
      authUrl // Return the URL so it can be manually opened
    });
  }
});

// Check authentication status endpoint
app.get('/auth/status', (req, res) => {
  res.json(authState);
});

// Start the authentication server
server.listen(config.server.authPort, () => {
  logger.info(`Authentication server running on port ${config.server.authPort}`);
});

// Mock function for token exchange - would need to be implemented
// with actual Microsoft Graph API token endpoints
async function exchangeCodeForToken(code, state) {
  // This would be implemented with actual token exchange logic
  // using the Microsoft Identity platform
  logger.info(`Exchanging code for token with state: ${state}`);
  
  // For the purposes of this example, we'll mock a successful token response
  const mockTokenResponse = {
    access_token: 'mock_access_token',
    refresh_token: 'mock_refresh_token',
    id_token: 'mock_id_token',
    expires_in: 3600
  };
  
  // Extract user information from ID token
  const userId = 'mock_user_id'; // Would be extracted from the decoded ID token
  
  // Save the token using the token manager
  await saveToken(userId, mockTokenResponse);
  
  return { userId };
}