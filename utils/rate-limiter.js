const NodeCache = require('node-cache');
const config = require('../config');
const logger = require('./logger');

class RateLimiter {
  constructor(options = {}) {
    this.windowMs = options.windowMs || config.rateLimit.windowMs;
    this.maxRequests = options.maxRequests || config.rateLimit.maxRequests;
    
    // Cache to store request counts by user or IP
    this.cache = new NodeCache({
      stdTTL: this.windowMs / 1000, // Convert to seconds for NodeCache
      checkperiod: 60 // Check for expired keys every 60 seconds
    });
    
    logger.info(`Rate limiter initialized: ${this.maxRequests} requests per ${this.windowMs}ms`);
  }
  
  /**
   * Check if the request is within rate limits
   * @param {string} key - Identifier for the requestor (userId, IP, etc.)
   * @returns {Promise<boolean>} - True if within limits, false otherwise
   * @throws {Error} - If rate limit is exceeded with retry-after information
   */
  async check(key = 'default') {
    const current = this.cache.get(key) || 0;
    
    if (current >= this.maxRequests) {
      // Calculate time remaining in the current window
      const ttl = this.cache.getTtl(key);
      const now = Date.now();
      const retryAfter = Math.ceil((ttl - now) / 1000); // in seconds
      
      logger.warn(`Rate limit exceeded for ${key}: ${current}/${this.maxRequests}`);
      
      const error = new Error(`Rate limit exceeded. Try again in ${retryAfter} seconds.`);
      error.name = 'RateLimitExceededError';
      error.retryAfter = retryAfter;
      throw error;
    }
    
    // Increment the counter
    this.cache.set(key, current + 1);
    
    logger.debug(`Rate limit check passed for ${key}: ${current + 1}/${this.maxRequests}`);
    return true;
  }
  
  /**
   * Reset rate limit counter for a specific key
   * @param {string} key - Identifier to reset
   */
  reset(key = 'default') {
    this.cache.del(key);
    logger.debug(`Rate limit counter reset for ${key}`);
  }
  
  /**
   * Get current usage for a specific key
   * @param {string} key - Identifier to check
   * @returns {Object} - Current usage information
   */
  getUsage(key = 'default') {
    const current = this.cache.get(key) || 0;
    let remaining = this.maxRequests - current;
    if (remaining < 0) remaining = 0;
    
    let reset = null;
    if (current > 0) {
      const ttl = this.cache.getTtl(key);
      reset = Math.ceil((ttl - Date.now()) / 1000); // in seconds
    }
    
    return {
      limit: this.maxRequests,
      current,
      remaining,
      reset
    };
  }
}

// Create a singleton instance
const rateLimiter = new RateLimiter();

module.exports = { rateLimiter };