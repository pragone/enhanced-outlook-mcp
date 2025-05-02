require('dotenv').config();
const os = require('os');
const path = require('path');

const config = {
  // Server configuration
  server: {
    name: 'Enhanced Outlook MCP Server',
    version: '1.0.0',
    port: process.env.PORT || 3000,
    logLevel: process.env.LOG_LEVEL || 'info',
    tokenStoragePath: process.env.TOKEN_STORAGE_PATH || path.join(os.homedir(), '.enhanced-outlook-mcp-tokens.json')
  },
  
  // Microsoft Graph API configuration
  microsoft: {
    clientId: process.env.MS_CLIENT_ID,
    authority: process.env.MS_AUTHORITY || 'https://login.microsoftonline.com/common',
    redirectUri: process.env.MS_REDIRECT_URI || 'http://localhost:3000/auth/callback',
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
  },
  
  // Tool relationships and metadata
  toolMetadata: {
    // Auth tools
    'authenticate': {
      dependencies: [],
      category: 'auth',
    },
    'check_auth_status': {
      dependencies: [],
      category: 'auth',
    },
    'revoke_authentication': {
      dependencies: ['authenticate'],
      category: 'auth',
    },
    'list_authenticated_users': {
      dependencies: [],
      category: 'auth',
    },
    
    // Email tools
    'list_emails': {
      dependencies: ['authenticate'],
      category: 'email',
      related: ['read_email', 'search_emails', 'move_emails']
    },
    'search_emails': {
      dependencies: ['authenticate'],
      category: 'email',
      related: ['read_email', 'list_emails']
    },
    'read_email': {
      dependencies: ['authenticate', 'list_emails'],
      category: 'email',
      related: ['list_attachments', 'reply_email', 'forward_email']
    },
    'mark_email': {
      dependencies: ['authenticate', 'list_emails'],
      category: 'email',
      related: ['read_email']
    },
    'send_email': {
      dependencies: ['authenticate'],
      category: 'email',
      related: ['create_draft', 'add_attachment']
    },
    'create_draft': {
      dependencies: ['authenticate'],
      category: 'email',
      related: ['send_email', 'add_attachment']
    },
    'reply_email': {
      dependencies: ['authenticate', 'read_email'],
      category: 'email',
      related: ['list_emails', 'search_emails']
    },
    'forward_email': {
      dependencies: ['authenticate', 'read_email'],
      category: 'email',
      related: ['list_emails', 'search_emails']
    },
    
    // Attachment tools
    'list_attachments': {
      dependencies: ['authenticate', 'list_emails', 'read_email'],
      category: 'attachment',
      related: ['get_attachment']
    },
    'get_attachment': {
      dependencies: ['authenticate', 'list_emails', 'list_attachments'],
      category: 'attachment',
      related: ['read_email']
    },
    'add_attachment': {
      dependencies: ['authenticate', 'create_draft'],
      category: 'attachment',
      related: ['send_email', 'delete_attachment']
    },
    'delete_attachment': {
      dependencies: ['authenticate', 'list_attachments'],
      category: 'attachment',
      related: ['add_attachment']
    },
    
    // Folder tools
    'list_folders': {
      dependencies: ['authenticate'],
      category: 'folder',
      related: ['get_folder', 'list_emails']
    },
    'get_folder': {
      dependencies: ['authenticate', 'list_folders'],
      category: 'folder',
      related: ['list_emails']
    },
    'create_folder': {
      dependencies: ['authenticate', 'list_folders'],
      category: 'folder',
      related: ['update_folder', 'move_folder']
    },
    'update_folder': {
      dependencies: ['authenticate', 'list_folders', 'get_folder'],
      category: 'folder',
      related: ['delete_folder', 'move_folder']
    },
    'delete_folder': {
      dependencies: ['authenticate', 'list_folders'],
      category: 'folder',
      related: ['update_folder']
    },
    'move_folder': {
      dependencies: ['authenticate', 'list_folders'],
      category: 'folder',
      related: ['update_folder']
    },
    'move_emails': {
      dependencies: ['authenticate', 'list_emails', 'list_folders'],
      category: 'email',
      related: ['copy_emails']
    },
    'copy_emails': {
      dependencies: ['authenticate', 'list_emails', 'list_folders'],
      category: 'email',
      related: ['move_emails']
    },
    
    // Calendar tools
    'list_calendars': {
      dependencies: ['authenticate'],
      category: 'calendar',
      related: ['list_events', 'create_event']
    },
    'list_events': {
      dependencies: ['authenticate', 'list_calendars'],
      category: 'calendar',
      related: ['get_event', 'create_event']
    },
    'get_event': {
      dependencies: ['authenticate', 'list_events'],
      category: 'calendar',
      related: ['update_event', 'respond_to_event']
    },
    'create_event': {
      dependencies: ['authenticate', 'list_calendars'],
      category: 'calendar',
      related: ['find_meeting_times', 'update_event']
    },
    'update_event': {
      dependencies: ['authenticate', 'get_event'],
      category: 'calendar',
      related: ['cancel_event', 'delete_event']
    },
    'respond_to_event': {
      dependencies: ['authenticate', 'get_event'],
      category: 'calendar',
      related: ['list_events']
    },
    'delete_event': {
      dependencies: ['authenticate', 'get_event'],
      category: 'calendar',
      related: ['cancel_event']
    },
    'cancel_event': {
      dependencies: ['authenticate', 'get_event'],
      category: 'calendar',
      related: ['delete_event']
    },
    'find_meeting_times': {
      dependencies: ['authenticate'],
      category: 'calendar',
      related: ['create_event', 'list_calendars']
    },
    
    // Rules tools
    'list_rules': {
      dependencies: ['authenticate'],
      category: 'rule',
      related: ['get_rule', 'create_rule']
    },
    'get_rule': {
      dependencies: ['authenticate', 'list_rules'],
      category: 'rule',
      related: ['update_rule', 'delete_rule']
    },
    'create_rule': {
      dependencies: ['authenticate'],
      category: 'rule',
      related: ['list_rules', 'update_rule']
    },
    'update_rule': {
      dependencies: ['authenticate', 'get_rule'],
      category: 'rule',
      related: ['delete_rule']
    },
    'delete_rule': {
      dependencies: ['authenticate', 'get_rule'],
      category: 'rule',
      related: ['list_rules']
    }
  }
};

module.exports = config;