const express = require('express');
const bodyParser = require('body-parser');
const http = require('http');
const socketIo = require('socket.io');
const open = require('open');
const config = require('./config');
const logger = require('./utils/logger');
const { saveToken } = require('./auth/token-manager');
const axios = require('axios');

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
  logger.info('==== AUTH SERVER DEBUG START ====');
  logger.info(`Auth start request received. Body: ${JSON.stringify(req.body)}`);
  logger.info(`Auth start request headers: ${JSON.stringify(req.headers)}`);
  
  const { clientId, scopes, redirectUri, state = 'default' } = req.body;
  
  // Log request details
  logger.info(`Auth start request received: clientId=${clientId}, redirectUri=${redirectUri}, state=${state}`);
  logger.info(`Requested scopes: ${JSON.stringify(scopes)}`);
  
  // More detailed validation
  const validationErrors = [];
  if (!clientId) validationErrors.push('Missing clientId');
  if (!scopes) validationErrors.push('Missing scopes');
  if (!redirectUri) validationErrors.push('Missing redirectUri');
  
  if (validationErrors.length > 0) {
    const errorMessage = `Missing required parameters: ${validationErrors.join(', ')}`;
    logger.error(errorMessage);
    logger.error(`clientId=${!!clientId}, scopes=${!!scopes}, redirectUri=${!!redirectUri}`);
    
    logger.info('==== AUTH SERVER DEBUG END ====');
    return res.status(400).json({
      error: errorMessage
    });
  }
  
  // Validate the redirectUri format
  try {
    new URL(redirectUri);
  } catch (e) {
    logger.error(`Invalid redirectUri format: ${redirectUri}`);
    logger.info('==== AUTH SERVER DEBUG END ====');
    return res.status(400).json({
      error: 'Invalid redirectUri format'
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
  try {
    const scopesParam = encodeURIComponent(Array.isArray(scopes) ? scopes.join(' ') : scopes);
    const encodedRedirectUri = encodeURIComponent(redirectUri);
    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodedRedirectUri}&scope=${scopesParam}&response_mode=query&state=${state}`;
    
    logger.info(`Starting authentication flow with state: ${state}`);
    logger.info(`Authentication URL: ${authUrl}`);
    
    // Open the browser for authentication
    try {
      await open(authUrl);
      logger.info('Browser opened successfully with auth URL');
      logger.info('==== AUTH SERVER DEBUG END ====');
      res.json({ status: 'authentication_started', authUrl });
    } catch (err) {
      logger.error('Failed to open browser:', err);
      
      authState = {
        isAuthenticating: false,
        userId: null,
        error: `Failed to open browser: ${err.message}`
      };
      io.emit('authStatus', authState);
      
      logger.info('==== AUTH SERVER DEBUG END ====');
      res.status(500).json({
        error: 'Failed to open browser for authentication',
        authUrl // Return the URL so it can be manually opened
      });
    }
  } catch (error) {
    logger.error(`Error constructing auth URL: ${error.message}`);
    logger.error(error.stack);
    
    authState = {
      isAuthenticating: false,
      userId: null,
      error: `Error constructing auth URL: ${error.message}`
    };
    io.emit('authStatus', authState);
    
    logger.info('==== AUTH SERVER DEBUG END ====');
    res.status(500).json({
      error: `Error constructing auth URL: ${error.message}`
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

// Exchange authorization code for token
async function exchangeCodeForToken(code, state) {
  if (!code) {
    logger.error('Authorization code is missing');
    throw new Error('Authorization code is required');
  }
  
  try {
    logger.info(`Exchanging code for token with state: ${state}`);
    logger.info(`Configured redirectUri: ${config.microsoft.redirectUri}`);
    
    // Log if client secret is present (without revealing it)
    if (!process.env.MS_CLIENT_SECRET) {
      logger.warn('MS_CLIENT_SECRET environment variable is missing');
    } else {
      logger.info('MS_CLIENT_SECRET is present');
    }
    
    // Prepare token request parameters
    const tokenParams = {
      client_id: config.microsoft.clientId,
      client_secret: process.env.MS_CLIENT_SECRET,
      code: code,
      redirect_uri: config.microsoft.redirectUri,
      grant_type: 'authorization_code'
    };
    
    logger.info(`Token request parameters: client_id=${tokenParams.client_id}, redirect_uri=${tokenParams.redirect_uri}, grant_type=${tokenParams.grant_type}`);
    
    // Make token request to Microsoft identity platform
    logger.info('Sending token request to Microsoft identity platform...');
    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      new URLSearchParams(tokenParams),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      }
    );
    
    logger.info('Token request successful');
    logger.info(`Token response status: ${response.status}`);
    logger.info(`Token response includes: ${Object.keys(response.data).join(', ')}`);
    
    const tokenData = response.data;
    
    // Calculate expiration time
    const expiresIn = tokenData.expires_in || 3600;
    const expiresAt = Date.now() + expiresIn * 1000;
    
    // Extract user information from ID token
    let userId;
    if (state !== 'default') {
      userId = state;
    } else {
      try {
        // Decode the JWT token to get user information
        const idTokenParts = tokenData.id_token.split('.');
        if (idTokenParts.length !== 3) {
          throw new Error('Invalid ID token format');
        }
        
        // Decode the payload (second part)
        const payload = JSON.parse(Buffer.from(idTokenParts[1], 'base64').toString());
        
        // Use the user's email as the ID if available, otherwise use preferred_username
        userId = payload.email || payload.preferred_username;
        
        if (!userId) {
          throw new Error('Could not extract user identifier from ID token');
        }
        
        logger.info(`Extracted user ID from token: ${userId}`);
      } catch (error) {
        logger.error(`Error decoding ID token: ${error.message}`);
        throw new Error(`Failed to extract user identifier: ${error.message}`);
      }
    }
    
    // Create complete token data
    const completeTokenData = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      id_token: tokenData.id_token,
      expires_in: expiresIn,
      expires_at: expiresAt,
      scope: tokenData.scope,
      token_type: tokenData.token_type || 'Bearer'
    };
    
    // Save token
    logger.info(`Saving token for user: ${userId}`);
    await saveToken(userId, completeTokenData);
    logger.info('Token saved successfully');
    
    return { userId };
  } catch (error) {
    logger.error(`Token exchange error: ${error.message}`);
    
    // Log more detailed error information
    if (error.response) {
      logger.error(`Error response status: ${error.response.status}`);
      logger.error(`Error response headers: ${JSON.stringify(error.response.headers)}`);
      logger.error(`Error response data: ${JSON.stringify(error.response.data)}`);
    } else if (error.request) {
      logger.error('No response received from server');
      logger.error(`Request details: ${error.request._header || JSON.stringify(error.request)}`);
    } else {
      logger.error(`Error details: ${error.stack}`);
    }
    
    throw new Error(`Failed to exchange code for token: ${error.message}`);
  }
}