require('dotenv').config();
const os = require('os');
const path = require('path');

const config = {
  // Server configuration
  server: {
    name: 'Enhanced Outlook MCP Server',
    version: '1.0.0',
    port: process.env.PORT || 3000,
    authPort: process.env.AUTH_PORT || 3333,
    logLevel: process.env.LOG_LEVEL || 'info',
    tokenStoragePath: process.env.TOKEN_STORAGE_PATH || path.join(os.homedir(), '.enhanced-outlook-mcp-tokens.json')
  },
  
  // Microsoft Graph API configuration
  microsoft: {
    clientId: process.env.MS_CLIENT_ID,
    authority: process.env.MS_AUTHORITY || 'https://login.microsoftonline.com/common',
    redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3333/auth/callback',
    scopes: process.env.MS_SCOPES ? process.env.MS_SCOPES.split(',') : [
      'openid',
      'profile',
      'offline_access',
      'User.Read',
      'Mail.Read',
      'Mail.ReadWrite',
      'Mail.Send',
      'MailboxSettings.Read',
      'Calendars.ReadWrite',
      'Contacts.Read'
    ],
    apiBaseUrl: process.env.MS_API_BASE_URL || 'https://graph.microsoft.com/v1.0',
    graphApiResponseLimit: parseInt(process.env.MS_GRAPH_API_RESPONSE_LIMIT || '50')
  },
  
  // Email configuration
  email: {
    defaultFields: process.env.EMAIL_DEFAULT_FIELDS ? process.env.EMAIL_DEFAULT_FIELDS.split(',') : [
      'id',
      'subject',
      'bodyPreview',
      'receivedDateTime',
      'from',
      'toRecipients',
      'ccRecipients',
      'importance',
      'hasAttachments',
      'isDraft'
    ],
    maxEmailsPerRequest: parseInt(process.env.MAX_EMAILS_PER_REQUEST || '20')
  },
  
  // Calendar configuration
  calendar: {
    defaultFields: process.env.CALENDAR_DEFAULT_FIELDS ? process.env.CALENDAR_DEFAULT_FIELDS.split(',') : [
      'id', 
      'subject', 
      'bodyPreview', 
      'start', 
      'end', 
      'location', 
      'attendees', 
      'organizer',
      'isAllDay',
      'isCancelled'
    ],
    maxEventsPerRequest: parseInt(process.env.MAX_EVENTS_PER_REQUEST || '20')
  },
  
  // Testing configuration
  testing: {
    enabled: process.env.TEST_MODE === 'true',
    mockDataDir: process.env.MOCK_DATA_DIR || './utils/mock-data'
  },
  
  // Rate limiting
  rateLimit: {
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS || '60000'),
    maxRequests: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS || '30')
  }
};

module.exports = config;