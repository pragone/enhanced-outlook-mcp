const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');
const url = require('url');
const open = require('open');
const { PublicClientApplication, ConfidentialClientApplication } = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
const logger = require('../utils/logger');
const config = require('../config');

/**
 * Authentication service using MSAL for Microsoft Graph API
 * Combines the architectural benefits of the OneNote MCP approach
 * with the user experience of web-based authentication
 */
class AuthService {
  constructor(options = {}) {
    this.client = null;
    this.tokenCache = null;
    this.httpServer = null;
    this.authInProgress = false;
    this.authPromiseResolve = null;
    this.authPromiseReject = null;
    
    // Configuration options
    this.clientId = options.clientId || config.microsoft.clientId;
    this.clientSecret = options.clientSecret || process.env.MS_CLIENT_SECRET;
    this.redirectUri = options.redirectUri || config.microsoft.redirectUri;
    this.authority = options.authority || config.microsoft.authority;
    this.scopes = options.scopes || config.microsoft.scopes;
    this.cacheFile = options.cacheFile || path.join(os.homedir(), '.enhanced-outlook-mcp-token-cache.json');
    
    if (!this.clientId) {
      throw new Error('Client ID is required. Set MS_CLIENT_ID in your .env file.');
    }

    logger.debug(`Auth service created with client ID: ${this.clientId.substring(0, 8)}...`);
    logger.debug(`Redirect URI: ${this.redirectUri}`);
    logger.debug(`Scopes: ${Array.isArray(this.scopes) ? this.scopes.join(', ') : this.scopes}`);
    
    this.pca = this._createPca();
    logger.info('Auth service initialized with client ID: ' + this.clientId.substring(0, 5) + '...');
  }

  /**
   * Create and configure the MSAL application
   * @private
   * @returns {PublicClientApplication|ConfidentialClientApplication}
   */
  _createPca() {
    // Setup token cache persistence
    const beforeCacheAccess = async (cacheContext) => {
      try {
        if (fs.existsSync(this.cacheFile)) {
          const data = fs.readFileSync(this.cacheFile, 'utf-8');
          cacheContext.tokenCache.deserialize(data);
          logger.info('Token cache loaded from disk');
          logger.debug(`Token cache loaded from: ${this.cacheFile}`);
        } else {
          logger.debug(`No token cache file found at: ${this.cacheFile}`);
        }
      } catch (error) {
        logger.error(`Error loading token cache: ${error.message}`);
        logger.debug(`Error loading token cache: ${error.message}`);
      }
    };

    const afterCacheAccess = async (cacheContext) => {
      try {
        if (cacheContext.cacheHasChanged) {
          const data = cacheContext.tokenCache.serialize();
          // Ensure directory exists
          const dir = path.dirname(this.cacheFile);
          if (!fs.existsSync(dir)) {
            fs.mkdirSync(dir, { recursive: true });
          }
          fs.writeFileSync(this.cacheFile, data, {
            encoding: 'utf-8',
            mode: 0o600 // File permissions: user read/write only
          });
          logger.info('Token cache saved to disk');
          logger.debug(`Token cache saved to: ${this.cacheFile}`);
        }
      } catch (error) {
        logger.error(`Error saving token cache: ${error.message}`);
        logger.debug(`Error saving token cache: ${error.message}`);
      }
    };

    const cachePlugin = {
      beforeCacheAccess,
      afterCacheAccess,
    };

    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: this.authority,
      },
      cache: {
        cachePlugin,
      },
    };

    logger.debug("Creating MSAL application with config:", JSON.stringify({
      clientId: this.clientId.substring(0, 8) + '...',
      authority: this.authority,
      clientSecretProvided: !!this.clientSecret
    }));

    // If we have a client secret, use ConfidentialClientApplication
    if (this.clientSecret) {
      msalConfig.auth.clientSecret = this.clientSecret;
      logger.debug("Using ConfidentialClientApplication (with client secret)");
      return new ConfidentialClientApplication(msalConfig);
    }

    // Otherwise use PublicClientApplication
    logger.debug("Using PublicClientApplication (without client secret)");
    return new PublicClientApplication(msalConfig);
  }

  /**
   * Initialize the auth service and attempt to acquire a token
   * @returns {Promise<AuthService>} This instance
   */
  async initialize() {
    try {
      this.tokenCache = this.pca.getTokenCache();
      logger.info('Starting authentication initialization');
      logger.debug("Initializing authentication service...");
      
      let authResult = await this._acquireTokenSilently();
      
      if (authResult) {
        // Initialize Graph client with acquired token
        this._initializeGraphClient(authResult.accessToken);
        logger.info('Auth service initialized successfully with cached token');
        logger.debug("Auth service initialized successfully with cached token");
        return this;
      } else {
        // We don't want to auto-trigger interactive auth during initialization
        // Just inform that auth is needed and return the service
        logger.info('No cached token available. Interactive authentication required.');
        logger.debug("No cached token available. Interactive authentication will be required.");
        return this;
      }
    } catch (error) {
      logger.error(`Auth initialization error: ${error.message}`);
      logger.debug(`Error initializing auth service: ${error.message}`);
      throw error;
    }
  }

  /**
   * Attempt to acquire a token silently from the cache
   * @returns {Promise<Object|null>} The auth result or null if failed
   * @private
   */
  async _acquireTokenSilently() {
    try {
      const accounts = await this.tokenCache.getAllAccounts();
      
      if (accounts.length > 0) {
        logger.debug(`Found ${accounts.length} account(s) in cache, attempting silent token acquisition`);
        logger.info(`Found ${accounts.length} account(s) in cache, attempting silent token acquisition`);
        try {
          // Try to acquire token silently
          const silentRequest = {
            scopes: this.scopes,
            account: accounts[0],
            forceRefresh: false
          };
          
          const silentResult = await this.pca.acquireTokenSilent(silentRequest);
          logger.info('Token acquired silently');
          logger.debug("Token acquired silently");
          return silentResult;
        } catch (silentError) {
          logger.warn(`Silent token acquisition failed: ${silentError.message}`);
          logger.debug(`Silent token acquisition failed: ${silentError.message}`);
          return null;
        }
      } else {
        logger.info('No accounts found in cache');
        logger.debug("No accounts found in token cache");
        return null;
      }
    } catch (error) {
      logger.error(`Silent token acquisition error: ${error.message}`);
      logger.debug(`Error during silent token acquisition: ${error.message}`);
      return null;
    }
  }

  /**
   * Initialize a local HTTP server to handle the OAuth callback
   * @returns {Promise<http.Server>} The HTTP server
   * @private
   */
  _initAuthCallbackServer() {
    return new Promise((resolve, reject) => {
      if (this.httpServer) {
        // Server already exists
        logger.debug("Auth callback server already exists, reusing it");
        resolve(this.httpServer);
        return;
      }

      try {
        // Extract port from redirect URI
        const redirectUrl = new URL(this.redirectUri);
        const port = redirectUrl.port || (redirectUrl.protocol === 'https:' ? 443 : 80);
        const path = redirectUrl.pathname;
        
        logger.debug(`Initializing auth callback server on port ${port} with path "${path}"`);
        logger.debug(`Full redirect URI: ${this.redirectUri}`);

        this.httpServer = http.createServer((req, res) => {
          const reqUrl = url.parse(req.url, true);
          logger.debug(`Received request: ${req.method} ${req.url}`);

          // Check if this is the callback URL
          if (reqUrl.pathname === path) {
            logger.info('Received OAuth callback');
            logger.debug('Received OAuth callback!');

            // For security, use a generic response
            const htmlResponse = `
<!DOCTYPE html>
<html>
<head>
  <title>Authentication Complete</title>
  <style>
    body { font-family: Arial, sans-serif; text-align: center; margin-top: 50px; }
    .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
    .success { color: #4CAF50; }
    .error { color: #F44336; }
  </style>
</head>
<body>
  <div class="container">
    <h1 id="status-header"></h1>
    <p id="status-message"></p>
  </div>

  <script>
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get('code');
    const error = urlParams.get('error');
    
    if (code) {
      document.getElementById('status-header').textContent = 'Authentication Successful';
      document.getElementById('status-header').className = 'success';
      document.getElementById('status-message').textContent = 'You have been successfully authenticated. You can now close this window and return to your application.';
    } else if (error) {
      document.getElementById('status-header').textContent = 'Authentication Failed';
      document.getElementById('status-header').className = 'error';
      document.getElementById('status-message').textContent = 'Authentication failed: ' + error + ' - ' + urlParams.get('error_description');
    } else {
      document.getElementById('status-header').textContent = 'Authentication Status Unknown';
      document.getElementById('status-message').textContent = 'No authentication code or error was received.';
    }
  </script>
</body>
</html>`;

            // If we have an authorization code, process it
            if (reqUrl.query.code) {
              const code = reqUrl.query.code;
              const state = reqUrl.query.state;
              
              logger.info(`Received authorization code with state: ${state}`);
              logger.debug(`Received authorization code with state: ${state || 'none'}`);
              
              // Exchange the code for tokens - async but don't wait for it
              this._exchangeCodeForToken(code).then((tokenResponse) => {
                logger.debug("Successfully exchanged code for token!");
                if (this.authPromiseResolve) {
                  this.authPromiseResolve(tokenResponse);
                  this.authPromiseResolve = null;
                  this.authPromiseReject = null;
                }
              }).catch((error) => {
                logger.error(`Error exchanging code for token: ${error.message}`);
                logger.debug(`Error exchanging code for token: ${error.message}`);
                if (this.authPromiseReject) {
                  this.authPromiseReject(error);
                  this.authPromiseResolve = null;
                  this.authPromiseReject = null;
                }
              });
              
              // Send success response
              res.writeHead(200, { 'Content-Type': 'text/html' });
              res.end(htmlResponse);
            } else if (reqUrl.query.error) {
              // Handle authentication error
              const error = reqUrl.query.error;
              const errorDescription = reqUrl.query.error_description;
              
              logger.error(`Authentication error: ${error} - ${errorDescription}`);
              logger.debug(`Authentication error: ${error} - ${errorDescription}`);
              
              if (this.authPromiseReject) {
                this.authPromiseReject(new Error(`${error}: ${errorDescription}`));
                this.authPromiseResolve = null;
                this.authPromiseReject = null;
              }
              
              // Send error response
              res.writeHead(200, { 'Content-Type': 'text/html' });
              res.end(htmlResponse);
            } else {
              // Unknown callback format
              logger.error('Unknown OAuth callback format');
              logger.debug('Unknown OAuth callback format - no code or error');
              
              // Reject if we have a pending promise
              if (this.authPromiseReject) {
                this.authPromiseReject(new Error('Unknown OAuth callback format'));
                this.authPromiseResolve = null;
                this.authPromiseReject = null;
              }
              
              // Send generic response
              res.writeHead(200, { 'Content-Type': 'text/html' });
              res.end(htmlResponse);
            }
          } else {
            // Not the callback path, return 404
            logger.debug(`Received request for non-callback path: ${reqUrl.pathname}`);
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            res.end('Not found');
          }
        });

        // Start the server on the given port
        this.httpServer.listen(port, () => {
          logger.info(`Authentication callback server listening on port ${port}`);
          logger.debug(`Authentication callback server listening on port ${port}`);
          resolve(this.httpServer);
        });

        this.httpServer.on('error', (error) => {
          logger.error(`Authentication server error: ${error.message}`);
          logger.debug(`Authentication server error: ${error.message}`);
          reject(error);
        });
      } catch (error) {
        logger.error(`Failed to initialize auth callback server: ${error.message}`);
        logger.debug(`Failed to initialize auth callback server: ${error.message}`);
        reject(error);
      }
    });
  }

  /**
   * Exchange authorization code for tokens
   * @param {string} code - The authorization code
   * @returns {Promise<Object>} The token response
   * @private
   */
  async _exchangeCodeForToken(code) {
    try {
      logger.info('Exchanging authorization code for tokens');
      logger.debug("Exchanging authorization code for tokens...");
      
      let tokenResponse;
      if (this.clientSecret) {
        // Use confidential client flow (with client secret)
        logger.debug("Using confidential client flow to exchange code");
        tokenResponse = await this.pca.acquireTokenByCode({
          code,
          scopes: this.scopes,
          redirectUri: this.redirectUri
        });
      } else {
        // Use public client flow (without client secret)
        logger.debug("Using public client flow to exchange code");
        tokenResponse = await this.pca.acquireTokenByCode({
          code,
          scopes: this.scopes,
          redirectUri: this.redirectUri
        });
      }
      
      logger.info('Successfully acquired tokens');
      logger.debug("Successfully acquired tokens!");
      
      // Initialize Graph client with the new token
      this._initializeGraphClient(tokenResponse.accessToken);
      
      return tokenResponse;
    } catch (error) {
      logger.error(`Failed to exchange code for token: ${error.message}`);
      logger.debug(`Failed to exchange code for token: ${error.message}`);
      throw error;
    }
  }

  /**
   * Initiate interactive browser-based authentication
   * @returns {Promise<Object>} The authentication result
   */
  async authenticate() {
    try {
      // Try to acquire token silently first
      logger.debug("Trying silent token acquisition first...");
      const silentResult = await this._acquireTokenSilently();
      if (silentResult) {
        logger.debug("Silent authentication successful!");
        this._initializeGraphClient(silentResult.accessToken);
        return silentResult;
      }
      
      // Silent auth failed, proceed with interactive auth
      logger.info('Starting browser-based authentication flow');
      logger.debug("Silent authentication failed. Starting browser-based authentication flow...");
      
      // Make sure we don't have multiple auth flows happening
      if (this.authInProgress) {
        throw new Error('Authentication is already in progress');
      }
      
      this.authInProgress = true;
      
      // Initialize the callback server
      logger.debug("Initializing callback server...");
      await this._initAuthCallbackServer();
      
      // Generate the auth URL
      logger.debug("Generating auth URL...");
      const authCodeUrlParameters = {
        scopes: this.scopes,
        redirectUri: this.redirectUri,
        responseMode: 'query',
      };
      
      const authCodeUrl = await this.pca.getAuthCodeUrl(authCodeUrlParameters);
      
      logger.info(`Generated authentication URL: ${authCodeUrl}`);
      logger.debug(`Generated authentication URL: ${authCodeUrl}`);
      
      // Create a promise that will be resolved when the auth is complete
      logger.debug("Creating auth promise...");
      const authPromise = new Promise((resolve, reject) => {
        this.authPromiseResolve = resolve;
        this.authPromiseReject = reject;
        
        // Set a timeout to reject the promise after 5 minutes
        setTimeout(() => {
          if (this.authPromiseReject) {
            logger.debug("Authentication timed out after 5 minutes");
            this.authPromiseReject(new Error('Authentication timed out after 5 minutes'));
            this.authPromiseResolve = null;
            this.authPromiseReject = null;
            this.authInProgress = false;
          }
        }, 5 * 60 * 1000);
      });
      
      // Open the auth URL in the browser
      logger.info('Opening browser for authentication');
      logger.debug("Opening browser for authentication...");
      try {
        await open(authCodeUrl);
        logger.debug("Browser launched successfully!");
      } catch (openError) {
        logger.debug(`Failed to open browser: ${openError.message}`);
        logger.debug(`Please manually open this URL: ${authCodeUrl}`);
      }
      
      // Wait for the auth to complete via the callback
      logger.debug("Waiting for authentication callback...");
      const authResult = await authPromise;
      
      this.authInProgress = false;
      logger.debug("Authentication completed successfully!");
      
      return authResult;
    } catch (error) {
      this.authInProgress = false;
      logger.error(`Authentication error: ${error.message}`);
      logger.debug(`Authentication error: ${error.message}`);
      throw error;
    }
  }

  /**
   * Initialize the Microsoft Graph API client
   * @param {string} accessToken - The access token
   * @private
   */
  _initializeGraphClient(accessToken) {
    this.client = Client.init({
      authProvider: (done) => {
        done(null, accessToken);
      },
    });
    logger.info('Graph client initialized');
    logger.debug("Microsoft Graph client initialized");
  }

  /**
   * Get the Microsoft Graph client
   * @returns {Client} The Graph client
   */
  async getGraphClient() {
    try {
      if (!this.client) {
        // Try to initialize with a token first
        logger.debug("Graph client not initialized, attempting to get token...");
        const authResult = await this._acquireTokenSilently();
        if (authResult) {
          logger.debug("Got token silently, initializing Graph client");
          this._initializeGraphClient(authResult.accessToken);
        } else {
          logger.debug("No cached token available");
          throw new Error('Graph client not initialized and no cached token available. Call authenticate() first.');
        }
      }
      return this.client;
    } catch (error) {
      logger.error(`Error getting Graph client: ${error.message}`);
      logger.debug(`Error getting Graph client: ${error.message}`);
      throw error;
    }
  }

  /**
   * Get a new access token (refreshing if necessary)
   * @returns {Promise<string>} The access token
   */
  async getAccessToken() {
    try {
      logger.debug("Getting access token...");
      const authResult = await this._acquireTokenSilently();
      if (authResult) {
        logger.debug("Got access token silently");
        return authResult.accessToken;
      }
      
      // No valid token available, need interactive auth
      logger.debug("No valid token in cache, starting interactive auth...");
      const interactiveResult = await this.authenticate();
      logger.debug("Got access token via interactive auth");
      return interactiveResult.accessToken;
    } catch (error) {
      logger.error(`Error getting access token: ${error.message}`);
      logger.debug(`Error getting access token: ${error.message}`);
      throw error;
    }
  }

  /**
   * Check if the user is authenticated
   * @returns {Promise<boolean>} True if authenticated
   */
  async isAuthenticated() {
    try {
      logger.debug("Checking if user is authenticated...");
      const accounts = await this.tokenCache.getAllAccounts();
      const authenticated = accounts.length > 0;
      logger.debug(`User authenticated: ${authenticated} (${accounts.length} accounts in cache)`);
      return authenticated;
    } catch (error) {
      logger.error(`Error checking authentication status: ${error.message}`);
      logger.debug(`Error checking authentication status: ${error.message}`);
      return false;
    }
  }

  /**
   * Sign out the user and clear the token cache
   * @returns {Promise<boolean>} True if sign out was successful
   */
  async signOut() {
    try {
      logger.debug("Signing out user...");
      const accounts = await this.tokenCache.getAllAccounts();
      
      if (accounts.length === 0) {
        logger.info('No accounts to sign out');
        logger.debug("No accounts to sign out");
        return false;
      }
      
      // Remove all accounts from the cache
      logger.debug(`Removing ${accounts.length} account(s) from token cache`);
      for (const account of accounts) {
        await this.tokenCache.removeAccount(account);
      }
      
      // Clean up the client
      this.client = null;
      logger.debug("Graph client cleared");
      
      logger.info('User signed out successfully');
      logger.debug("User signed out successfully");
      return true;
    } catch (error) {
      logger.error(`Error signing out: ${error.message}`);
      logger.debug(`Error signing out: ${error.message}`);
      return false;
    }
  }

  /**
   * Clean up resources (close server, etc.)
   */
  async cleanup() {
    if (this.httpServer) {
      logger.debug("Closing authentication server...");
      return new Promise((resolve) => {
        this.httpServer.close(() => {
          logger.info('Authentication server closed');
          logger.debug("Authentication server closed");
          this.httpServer = null;
          resolve();
        });
      });
    }
    logger.debug("No authentication server to close");
    return Promise.resolve();
  }

  /**
   * Get the authentication URL for starting the OAuth flow
   * This does not initiate the full authentication flow, just generates the URL
   * @returns {Promise<string>} The authentication URL
   */
  async getAuthUrl() {
    try {
      if (!this.httpServer) {
        logger.info('Initializing auth callback server');
        this._initAuthCallbackServer();
      }
      
      const authCodeUrlParameters = {
        scopes: this.scopes,
        redirectUri: this.redirectUri,
      };
      
      logger.info('Generating auth URL...');
      const authCodeUrl = await this.pca.getAuthCodeUrl(authCodeUrlParameters);
      logger.info(`Auth URL generated: ${authCodeUrl.substring(0, 60)}...`);
      
      return authCodeUrl;
    } catch (error) {
      logger.error(`Error generating auth URL: ${error.message}`);
      logger.debug(`Error generating auth URL: ${error.message}`);
      throw error;
    }
  }
}

// Export a singleton instance and factory function
let authServiceInstance = null;

/**
 * Get the AuthService instance, creating it if necessary
 * @param {Object} options - Configuration options
 * @returns {AuthService} The AuthService instance
 */
function getAuthService(options = {}) {
  if (!authServiceInstance) {
    logger.debug("Creating new AuthService instance");
    authServiceInstance = new AuthService(options);
  } else {
    logger.debug("Reusing existing AuthService instance");
  }
  return authServiceInstance;
}

module.exports = {
  AuthService,
  getAuthService,
}; 