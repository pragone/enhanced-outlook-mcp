const { 
  authenticateHandler,
  checkAuthStatusHandler,
  revokeAuthenticationHandler,
  listAuthenticatedUsersHandler
} = require('./tools');

// Authentication tool definitions
const authTools = [
  {
    name: 'authenticate',
    description: 'Authenticate with Microsoft Outlook using OAuth',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier for multi-user support (optional)'
        },
        scopes: {
          type: 'array',
          items: {
            type: 'string'
          },
          description: 'OAuth scopes to request (optional, defaults to configured scopes)'
        }
      }
    },
    handler: authenticateHandler
  },
  {
    name: 'check_auth_status',
    description: 'Check authentication status',
    parameters: {
      type: 'object',
      properties: {}
    },
    handler: checkAuthStatusHandler
  },
  {
    name: 'revoke_authentication',
    description: 'Revoke authentication and delete stored tokens',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        }
      }
    },
    handler: revokeAuthenticationHandler
  },
  {
    name: 'list_authenticated_users',
    description: 'List all authenticated users',
    parameters: {
      type: 'object',
      properties: {}
    },
    handler: listAuthenticatedUsersHandler
  }
];

module.exports = authTools;