const axios = require('axios');
const config = require('../config');
const logger = require('./logger');
const { getToken } = require('../auth/token-manager');
const { rateLimiter } = require('./rate-limiter');
const { lookupDefaultUser } = require('./parameter-helpers');

/**
 * Microsoft Graph API client wrapper
 */
class GraphApiClient {
  /**
   * Create a new Graph API client
   * @param {string} userId - User ID for token retrieval
   */
  constructor(userId) {
    this.userId = userId;
    this.baseUrl = config.microsoft.apiBaseUrl;
    this.requestCount = 0;
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
    // Handle 'default' userId special case for Claude Desktop compatibility
    let effectiveUserId = this.userId;
    
    if (this.userId === 'default') {
      logger.debug("'default' userId detected, looking up first available user");
      const actualUserId = await lookupDefaultUser();
      if (actualUserId) {
        effectiveUserId = actualUserId;
        logger.debug(`Mapped 'default' to actual userId: ${actualUserId}`);
      } else {
        logger.error("No available user found to map from 'default'");
      }
    }
    
    // Get access token
    const tokenInfo = await getToken(effectiveUserId);
    if (!tokenInfo || !tokenInfo.access_token) {
      logger.error(`No valid access token found for user '${effectiveUserId}'`);
      logger.info('Please authenticate with the proper scopes before trying again');
      throw new Error('No valid access token found. Please authenticate first.');
    }
    
    // Log success but not the token itself
    logger.debug(`Successfully retrieved access token for user '${effectiveUserId}'`);
    
    return {
      method,
      url: `${this.baseUrl}/${endpoint.replace(/^\//, '')}`,
      headers: {
        'Authorization': `Bearer ${tokenInfo.access_token}`,
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
      this.handleRequestError(error, method, endpoint);
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
      logger.warn(`Reached maximum page limit (${maxPages}) for paginated request to ${endpoint}`);
    }
    
    return allResults;
  }
}

module.exports = { GraphApiClient };