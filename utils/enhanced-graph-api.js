const axios = require('axios');
const config = require('../config');
const logger = require('./logger');
const { rateLimiter } = require('./rate-limiter');
const { lookupDefaultUser } = require('./parameter-helpers');

/**
 * Enhanced Microsoft Graph API client with integrated auth
 * This client works directly with the new AuthService
 */
class EnhancedGraphApiClient {
  /**
   * Create a new Graph API client
   * @param {Object} authService - The AuthService instance
   * @param {string} [userId='default'] - User ID for multi-user support
   */
  constructor(authService, userId = 'default') {
    if (!authService) {
      throw new Error('AuthService is required to create an EnhancedGraphApiClient');
    }
    this.authService = authService;
    this.userId = userId;
    this.baseUrl = config.microsoft.apiBaseUrl;
    this.requestCount = 0;
    this.msGraphClient = null;
  }

  /**
   * Initialize the Microsoft Graph client
   * @returns {Promise<Client>} Microsoft Graph client instance
   * @private
   */
  async _initializeGraphClient() {
    if (this.msGraphClient) {
      return this.msGraphClient;
    }

    try {
      // Get token directly from AuthService
      const token = await this.authService.getAccessToken();
      if (!token) {
        throw new Error('Failed to get access token');
      }

      // Create the client without initialization to avoid circular dependencies
      this.msGraphClient = await this.authService.getGraphClient();
      logger.info('Graph client initialized successfully');
      return this.msGraphClient;
    } catch (error) {
      logger.error(`Failed to initialize Graph client: ${error.message}`);
      throw new Error(`Graph client initialization failed: ${error.message}`);
    }
  }

  /**
   * Create an authenticated request config
   * @param {string} method - HTTP method
   * @param {string} endpoint - API endpoint (without base URL)
   * @param {Object} [data] - Request body for POST/PATCH/PUT
   * @param {Object} [params] - Query parameters
   * @param {Object} [headers] - Additional headers
   * @returns {Promise<Object>} - Axios request config
   */
  async createRequestConfig(method, endpoint, data = null, params = null, headers = {}) {
    // Get access token from AuthService
    const accessToken = await this.authService.getAccessToken();
    if (!accessToken) {
      logger.error('No valid access token found');
      logger.info('Please authenticate with the proper scopes before trying again');
      throw new Error('No valid access token found. Please authenticate first.');
    }
    
    // Log success but not the token itself
    logger.debug(`Successfully retrieved access token for user '${this.userId}'`);
    
    return {
      method,
      url: `${this.baseUrl}/${endpoint.replace(/^\//, '')}`,
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
        ...headers
      },
      data,
      params
    };
  }
  
  /**
   * Make a request to the Microsoft Graph API
   * @param {string} method - HTTP method
   * @param {string} endpoint - API endpoint (without base URL)
   * @param {Object} [data] - Request body for POST/PATCH/PUT
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async request(method, endpoint, data = null, params = null, options = {}) {
    try {
      // Check rate limits
      await rateLimiter.check(this.userId);
      
      const requestConfig = await this.createRequestConfig(
        method, 
        endpoint, 
        data, 
        params, 
        options.headers
      );
      
      this.requestCount++;
      const requestId = `${this.userId}-${this.requestCount}`;
      
      logger.debug(`Making Graph API request ${requestId}: ${method} ${endpoint}`);
      logger.debug(`Request params: ${JSON.stringify(params)}`);
      
      if (data) {
        logger.debug(`Request body: ${JSON.stringify(data)}`);
      }
      
      const response = await axios(requestConfig);
      
      logger.debug(`Graph API response ${requestId} status: ${response.status}`);
      
      if (options.returnFullResponse) {
        return response;
      }
      
      return response.data;
    } catch (error) {
      // Handle token expiration and auto-refresh
      if (error.response && error.response.status === 401) {
        logger.warn('Received 401 error, token may be expired. Attempting to refresh...');
        
        try {
          // Force token refresh and retry once
          await this.authService.getAccessToken(true); // Force refresh
          
          // Retry the request with the new token
          logger.info('Token refreshed, retrying request');
          
          const retryConfig = await this.createRequestConfig(
            method,
            endpoint,
            data,
            params,
            options.headers
          );
          
          const retryResponse = await axios(retryConfig);
          
          if (options.returnFullResponse) {
            return retryResponse;
          }
          
          return retryResponse.data;
        } catch (refreshError) {
          logger.error(`Token refresh failed: ${refreshError.message}`);
          this.handleRequestError(error, method, endpoint);
        }
      } else {
        this.handleRequestError(error, method, endpoint);
      }
    }
  }
  
  /**
   * Handle request errors
   * @param {Error} error - The error object
   * @param {string} method - HTTP method
   * @param {string} endpoint - API endpoint
   * @throws {Error} - Enhanced error with additional info
   */
  handleRequestError(error, method, endpoint) {
    if (error.response) {
      // The request was made and the server responded with a non-2xx status
      const status = error.response.status;
      const data = error.response.data;
      
      logger.error(`Graph API error (${status}) for ${method} ${endpoint}: ${JSON.stringify(data)}`);
      
      const graphError = new Error(data.error ? data.error.message : 'Unknown Graph API error');
      graphError.name = 'GraphAPIError';
      graphError.status = status;
      graphError.code = data.error ? data.error.code : 'unknown';
      graphError.data = data;
      
      // Handle authentication errors
      if (status === 401) {
        graphError.name = 'AuthenticationError';
        graphError.message = 'Authentication failed. Please re-authenticate.';
      }
      
      // Handle throttling
      if (status === 429) {
        graphError.name = 'ThrottlingError';
        const retryAfter = error.response.headers['retry-after'] || 30;
        graphError.retryAfter = parseInt(retryAfter, 10);
        graphError.message = `Request throttled. Try again in ${retryAfter} seconds.`;
      }
      
      throw graphError;
    } else if (error.request) {
      // The request was made but no response was received
      logger.error(`No response received for ${method} ${endpoint}: ${error.message}`);
      
      const networkError = new Error('No response received from Microsoft Graph API');
      networkError.name = 'NetworkError';
      networkError.request = error.request;
      throw networkError;
    } else {
      // Something happened in setting up the request
      logger.error(`Request setup error for ${method} ${endpoint}: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Make a GET request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async get(endpoint, params = null, options = {}) {
    return this.request('GET', endpoint, null, params, options);
  }
  
  /**
   * Make a POST request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} data - Request body
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async post(endpoint, data, params = null, options = {}) {
    return this.request('POST', endpoint, data, params, options);
  }
  
  /**
   * Make a PATCH request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} data - Request body
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async patch(endpoint, data, params = null, options = {}) {
    return this.request('PATCH', endpoint, data, params, options);
  }
  
  /**
   * Make a DELETE request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async delete(endpoint, params = null, options = {}) {
    return this.request('DELETE', endpoint, null, params, options);
  }
  
  /**
   * Handle paginated results from the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Array>} - Combined results from all pages
   */
  async getPaginated(endpoint, params = {}, options = {}) {
    let allResults = [];
    let nextLink = null;
    const maxPages = options.maxPages || 10; // Safety limit
    let pageCount = 0;
    
    // Make initial request
    const response = await this.get(endpoint, params, { ...options, returnFullResponse: true });
    
    if (response.data.value) {
      allResults = [...response.data.value];
    }
    
    nextLink = response.data['@odata.nextLink'];
    
    // Follow pagination links if they exist
    while (nextLink && pageCount < maxPages) {
      pageCount++;
      logger.debug(`Fetching next page (${pageCount}) from: ${nextLink}`);
      
      // Extract the relative path from the full URL
      const nextLinkPath = nextLink.replace(this.baseUrl, '');
      
      // Get the next page
      const nextPageResponse = await this.get(nextLinkPath, null, { ...options, returnFullResponse: true });
      
      if (nextPageResponse.data.value) {
        allResults = [...allResults, ...nextPageResponse.data.value];
      }
      
      nextLink = nextPageResponse.data['@odata.nextLink'];
    }
    
    if (nextLink && pageCount >= maxPages) {
      logger.warn(`Reached maximum page limit (${maxPages}), results may be incomplete`);
    }
    
    return allResults;
  }

  /**
   * Execute a Microsoft Graph SDK batch request with multiple operations
   * @param {Array<Object>} requests - Array of request objects with id, method, url, [headers], [body]
   * @returns {Promise<Object>} - Batch response
   */
  async batchRequest(requests) {
    if (!Array.isArray(requests) || requests.length === 0) {
      throw new Error('Batch requests must be a non-empty array');
    }

    const batchRequestBody = {
      requests: requests.map(req => ({
        id: req.id,
        method: req.method,
        url: req.url.replace(/^\/v1.0\//, ''), // Remove API version if present
        headers: req.headers || {},
        body: req.body
      }))
    };

    return this.post('$batch', batchRequestBody);
  }
}

/**
 * Create an enhanced Graph API client with the provided auth service
 * @param {Object} authService - AuthService instance
 * @param {string} [userId='default'] - User ID for multi-user support
 * @returns {EnhancedGraphApiClient} - Graph API client instance
 */
function createEnhancedGraphClient(authService, userId = 'default') {
  return new EnhancedGraphApiClient(authService, userId);
}

module.exports = {
  EnhancedGraphApiClient,
  createEnhancedGraphClient
}; 