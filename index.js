const { MCPServer } = require('@anthropic-ai/mcp');
const config = require('./config');
const logger = require('./utils/logger');

// Import module exports
const authTools = require('./auth');
const emailTools = require('./email');
const calendarTools = require('./calendar');
const folderTools = require('./folder');
const rulesTools = require('./rules');

// Gather all tools from modules
const TOOLS = [
  ...authTools,
  ...emailTools,
  ...calendarTools,
  ...folderTools,
  ...rulesTools
];

// Set up MCP server with comprehensive error handling
const server = new MCPServer({
  name: config.server.name,
  version: config.server.version,
  port: config.server.port,
  tools: TOOLS,
  onRequest: async (req) => {
    logger.info(`Received request for tool: ${req.tool}`);
    
    // Rate limiting check (would be implemented in the rate-limiter utility)
    try {
      const { rateLimiter } = require('./utils/rate-limiter');
      await rateLimiter.check();
    } catch (error) {
      logger.warn(`Rate limit exceeded: ${error.message}`);
      return {
        error: {
          type: 'rate_limit_exceeded',
          message: 'Too many requests. Please try again later.',
          retry_after: error.retryAfter
        }
      };
    }
  },
  onError: (error, req) => {
    logger.error(`Error in MCP request for tool ${req?.tool}:`, error);
    
    // Provide more informative error response based on error type
    if (error.name === 'AuthenticationError') {
      return {
        error: {
          type: 'authentication_error',
          message: 'Authentication failed. Please re-authenticate using the authenticate tool.',
        }
      };
    }
    
    if (error.name === 'GraphAPIError') {
      return {
        error: {
          type: 'api_error',
          message: `Microsoft Graph API error: ${error.message}`,
          code: error.code
        }
      };
    }
    
    // Generic error fallback
    return {
      error: {
        type: 'internal_error',
        message: 'An unexpected error occurred. Please try again or check the server logs.',
      }
    };
  }
});

// Start the server
server.start().then(() => {
  logger.info(`${config.server.name} v${config.server.version} started on port ${config.server.port}`);
  
  if (config.testing.enabled) {
    logger.info('Server running in TEST MODE with mock data');
  }
  
  logger.info(`Authentication server should be running on port ${config.server.authPort}`);
}).catch(error => {
  logger.error('Failed to start MCP server:', error);
  process.exit(1);
});