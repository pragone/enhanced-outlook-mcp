const { getAuthService } = require('../auth/auth-service');
const { EnhancedGraphApiClient } = require('./enhanced-graph-api');
const logger = require('./logger');
const auth = require('../auth/index');

/**
 * Adapter that provides legacy GraphApiClient compatible interface
 * but uses the new AuthService under the hood
 */
class GraphApiAdapter {
  /**
   * Create a new Graph API adapter
   * @param {string} userId - User ID for token retrieval
   */
  constructor(userId) {
    this.userId = userId;
    this.authService = null;
    this.enhancedClient = null;
    this.initialized = false;
  }

  /**
   * Initialize the adapter
   * @returns {Promise<void>}
   */
  async initialize() {
    if (this.initialized) {
      return;
    }

    try {
      // Get singleton auth service instance
      this.authService = await getAuthService();
      
      // Initialize the enhanced client
      this.enhancedClient = new EnhancedGraphApiClient(this.authService, this.userId);
      
      this.initialized = true;
      logger.info(`GraphApiAdapter initialized for user '${this.userId}'`);
    } catch (error) {
      logger.error(`Failed to initialize GraphApiAdapter: ${error.message}`);
      throw error;
    }
  }

  /**
   * Ensure the adapter is initialized
   * @private
   */
  async _ensureInitialized() {
    if (!this.initialized) {
      await this.initialize();
    }
  }

  /**
   * Make a request to the Microsoft Graph API using the enhanced client
   * This is the adapter's main method that translates to the enhanced client
   * @param {string} method - HTTP method
   * @param {string} endpoint - API endpoint (without base URL)
   * @param {Object} [data] - Request body for POST/PATCH/PUT
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async request(method, endpoint, data = null, params = null, options = {}) {
    await this._ensureInitialized();
    return this.enhancedClient.request(method, endpoint, data, params, options);
  }

  /**
   * Make a GET request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async get(endpoint, params = null, options = {}) {
    await this._ensureInitialized();
    return this.enhancedClient.get(endpoint, params, options);
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
    await this._ensureInitialized();
    return this.enhancedClient.post(endpoint, data, params, options);
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
    await this._ensureInitialized();
    return this.enhancedClient.patch(endpoint, data, params, options);
  }

  /**
   * Make a DELETE request to the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Object>} - API response
   */
  async delete(endpoint, params = null, options = {}) {
    await this._ensureInitialized();
    return this.enhancedClient.delete(endpoint, params, options);
  }

  /**
   * Handle paginated results from the Microsoft Graph API
   * @param {string} endpoint - API endpoint
   * @param {Object} [params] - Query parameters
   * @param {Object} [options] - Additional options
   * @returns {Promise<Array>} - Combined results from all pages
   */
  async getPaginated(endpoint, params = {}, options = {}) {
    await this._ensureInitialized();
    return this.enhancedClient.getPaginated(endpoint, params, options);
  }

  /**
   * Mark a message as read or unread
   * @param {string} userId - User ID
   * @param {string} messageId - Message ID
   * @param {boolean} isRead - Whether to mark the message as read
   * @returns {Promise<Object>} - API response
   */
  async markMessageAsRead(userId, messageId, isRead = true) {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}`).patch({
        isRead: isRead
      });
    });
  }
}

/**
 * Factory function to create Graph API clients based on configuration
 * This function decides whether to use the old GraphApiClient or the new adapter
 * for a smooth transition period
 * 
 * @param {string} userId - User ID for token retrieval
 * @param {Object} [options] - Options for client creation
 * @returns {Promise<GraphApiAdapter>} - Graph API client
 * @deprecated Use getGraphClient() or API-specific helpers (emailApi, calendarApi, etc.) instead
 * for proper token handling
 */
async function createGraphClient(userId, options = {}) {
  // Add deprecation warning
  logger.warn(`DEPRECATED: createGraphClient is deprecated and may be removed in future versions. ` +
              `Use getGraphClient() or API-specific helpers (emailApi, calendarApi, etc.) instead ` +
              `for proper token handling and authentication reuse.`);
  
  // Always use the new GraphApiAdapter regardless of environment variables
  logger.info(`Creating GraphApiAdapter for user '${userId}'`);
  const adapter = new GraphApiAdapter(userId);
  await adapter.initialize();
  return adapter;
}

/**
 * Get the appropriate Graph client based on feature and migration config
 * @param {string} userId The user ID
 * @param {string} feature The feature area (email, calendar, etc.)
 * @returns {Promise<Object>} The Graph client
 */
async function getGraphClient(userId, feature) {
  try {
    // Use the new auth system directly
    const auth = require('../auth/index');
    return await auth.getGraphClient(userId);
  } catch (error) {
    logger.error(`Fatal error getting Graph client: ${error.message}`);
    throw error;
  }
}

/**
 * Execute a Graph API request with the appropriate client based on migration config
 * 
 * @param {string} userId The user ID
 * @param {string} feature The feature area (email, calendar, etc.)
 * @param {Function} requestFn Function that takes a client and performs the request
 * @returns {Promise<any>} The API response
 */
async function executeGraphRequest(userId, feature, requestFn) {
  try {
    // Get the appropriate Graph client based on migration config
    const client = await getGraphClient(userId, feature);
    
    if (!client) {
      throw new Error(`Could not obtain a Graph client for ${feature}`);
    }
    
    // Ensure the client has an 'api' method
    if (typeof client.api !== 'function') {
      throw new Error('Invalid Graph client - missing api method');
    }
    
    // Execute the request
    return await requestFn(client);
  } catch (error) {
    logger.error(`Error executing Graph request for ${feature}: ${error.message}`);
    throw error;
  }
}

// Email related API calls
const emailApi = {
  listMessages: async (userId, options = {}) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      const { folderId, top, filter, orderBy, skip } = options;
      
      let endpoint = '/me/messages';
      if (folderId) {
        endpoint = `/me/mailFolders/${folderId}/messages`;
      }
      
      let request = client.api(endpoint);
      
      if (filter) request = request.filter(filter);
      if (orderBy) request = request.orderby(orderBy);
      if (top) request = request.top(top);
      if (skip) request = request.skip(skip);
      
      return await request.get();
    });
  },
  
  getMessage: async (userId, messageId, options = {}) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}`).get();
    });
  },
  
  sendMessage: async (userId, message, options = {}) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      const saveToSentItems = options.saveToSentItems !== false;
      return await client.api('/me/sendMail').post({ 
        message, 
        saveToSentItems 
      });
    });
  },
  
  createDraft: async (userId, message) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api('/me/messages').post(message);
    });
  },
  
  replyToMessage: async (userId, messageId, replyMessage, options = {}) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      const endpoint = options.replyAll 
        ? `/me/messages/${messageId}/replyAll` 
        : `/me/messages/${messageId}/reply`;
      
      return await client.api(endpoint).post({
        message: replyMessage,
        comment: options.comment
      });
    });
  },
  
  forwardMessage: async (userId, messageId, toRecipients, options = {}) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}/forward`).post({
        toRecipients,
        comment: options.comment
      });
    });
  },
  
  moveMessage: async (userId, messageId, destinationId) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}/move`).post({
        destinationId
      });
    });
  },
  
  copyMessage: async (userId, messageId, destinationId) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}/copy`).post({
        destinationId
      });
    });
  },
  
  markMessageAsRead: async (userId, messageId, isRead = true) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${messageId}`).patch({
        isRead: isRead
      });
    });
  },
  
  getAttachment: async (userId, emailId, attachmentId) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${emailId}/attachments/${attachmentId}`).get();
    });
  },
  
  listAttachments: async (userId, emailId) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${emailId}/attachments`).get();
    });
  },
  
  addAttachment: async (userId, emailId, attachment) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${emailId}/attachments`).post(attachment);
    });
  },
  
  deleteAttachment: async (userId, emailId, attachmentId) => {
    return executeGraphRequest(userId, 'email', async (client) => {
      return await client.api(`/me/messages/${emailId}/attachments/${attachmentId}`).delete();
    });
  }
};

// Calendar related API calls
const calendarApi = {
  listEvents: async (userId, options = {}) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      const { calendarId, top, filter, orderBy, skip } = options;
      
      let endpoint = '/me/events';
      if (calendarId) {
        endpoint = `/me/calendars/${calendarId}/events`;
      }
      
      let request = client.api(endpoint);
      
      if (filter) request = request.filter(filter);
      if (orderBy) request = request.orderby(orderBy);
      if (top) request = request.top(top);
      if (skip) request = request.skip(skip);
      
      return await request.get();
    });
  },
  
  getEvent: async (userId, eventId, calendarId) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      let endpoint;
      if (calendarId && calendarId !== 'primary') {
        endpoint = `/me/calendars/${calendarId}/events/${eventId}`;
      } else {
        endpoint = `/me/events/${eventId}`;
      }
      return await client.api(endpoint).get();
    });
  },
  
  createEvent: async (userId, event) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api('/me/events').post(event);
    });
  },
  
  updateEvent: async (userId, eventId, eventData) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api(`/me/events/${eventId}`).patch(eventData);
    });
  },
  
  deleteEvent: async (userId, eventId) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api(`/me/events/${eventId}`).delete();
    });
  },
  
  cancelEvent: async (userId, eventId, cancellationData = {}) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api(`/me/events/${eventId}/cancel`).post(cancellationData);
    });
  },
  
  respondToEvent: async (userId, eventId, response, responseData = {}) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      // response should be 'accept', 'tentativelyAccept', or 'decline'
      return await client.api(`/me/events/${eventId}/${response}`).post(responseData);
    });
  },
  
  findMeetingTimes: async (userId, findTimesRequest) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api('/me/findMeetingTimes').post(findTimesRequest);
    });
  },
  
  listCalendars: async (userId) => {
    return executeGraphRequest(userId, 'calendar', async (client) => {
      return await client.api('/me/calendars').get();
    });
  }
};

// Folder related API calls
const folderApi = {
  listFolders: async (userId, options = {}) => {
    return executeGraphRequest(userId, 'folder', async (client) => {
      const { parentFolderId, top } = options;
      
      let endpoint = '/me/mailFolders';
      if (parentFolderId) {
        endpoint = `/me/mailFolders/${parentFolderId}/childFolders`;
      }
      
      let request = client.api(endpoint);
      if (top) request = request.top(top);
      
      return await request.get();
    });
  },
  
  getFolder: async (userId, folderId) => {
    return executeGraphRequest(userId, 'folder', async (client) => {
      return await client.api(`/me/mailFolders/${folderId}`).get();
    });
  },
  
  createFolder: async (userId, folderData, parentFolderId) => {
    return executeGraphRequest(userId, 'folder', async (client) => {
      const endpoint = parentFolderId 
        ? `/me/mailFolders/${parentFolderId}/childFolders` 
        : '/me/mailFolders';
      
      return await client.api(endpoint).post(folderData);
    });
  },
  
  moveFolder: async (userId, sourceFolderId, destinationFolderId) => {
    // There's no direct way to move a folder in Graph API
    // This is a helper that copies the folder structure and contents to a new location
    // and then deletes the original if successful
    return executeGraphRequest(userId, 'folder', async (client) => {
      // Get the source folder details
      const sourceFolder = await client.api(`/me/mailFolders/${sourceFolderId}`).get();
      
      // Create a new folder with the same name in the destination parent
      const newFolder = await client.api(`/me/mailFolders/${destinationFolderId}/childFolders`).post({
        displayName: sourceFolder.displayName
      });
      
      // Move messages from source to new folder
      const messages = await client.api(`/me/mailFolders/${sourceFolderId}/messages`).get();
      
      if (messages && messages.value && messages.value.length > 0) {
        for (const message of messages.value) {
          await client.api(`/me/messages/${message.id}/move`).post({
            destinationId: newFolder.id
          });
        }
      }
      
      // Handle child folders recursively
      const childFolders = await client.api(`/me/mailFolders/${sourceFolderId}/childFolders`).get();
      
      if (childFolders && childFolders.value && childFolders.value.length > 0) {
        for (const childFolder of childFolders.value) {
          // Create a new child folder in the new parent
          const newChildFolder = await client.api(`/me/mailFolders/${newFolder.id}/childFolders`).post({
            displayName: childFolder.displayName
          });
          
          // Move child messages
          const childMessages = await client.api(`/me/mailFolders/${childFolder.id}/messages`).get();
          
          if (childMessages && childMessages.value && childMessages.value.length > 0) {
            for (const message of childMessages.value) {
              await client.api(`/me/messages/${message.id}/move`).post({
                destinationId: newChildFolder.id
              });
            }
          }
          
          // Could handle nested folders recursively here if needed
        }
      }
      
      // Delete the original folder now that everything is moved
      await client.api(`/me/mailFolders/${sourceFolderId}`).delete();
      
      return newFolder;
    });
  },
  
  deleteFolder: async (userId, folderId) => {
    return executeGraphRequest(userId, 'folder', async (client) => {
      return await client.api(`/me/mailFolders/${folderId}`).delete();
    });
  }
};

// Rules related API calls
const rulesApi = {
  listRules: async (userId) => {
    return executeGraphRequest(userId, 'rules', async (client) => {
      return await client.api('/me/mailFolders/inbox/messageRules').get();
    });
  },
  
  getRule: async (userId, ruleId) => {
    return executeGraphRequest(userId, 'rules', async (client) => {
      return await client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).get();
    });
  },
  
  createRule: async (userId, ruleData) => {
    return executeGraphRequest(userId, 'rules', async (client) => {
      return await client.api('/me/mailFolders/inbox/messageRules').post(ruleData);
    });
  },
  
  updateRule: async (userId, ruleId, ruleData) => {
    return executeGraphRequest(userId, 'rules', async (client) => {
      return await client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).patch(ruleData);
    });
  },
  
  deleteRule: async (userId, ruleId) => {
    return executeGraphRequest(userId, 'rules', async (client) => {
      return await client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).delete();
    });
  }
};

module.exports = {
  GraphApiAdapter,
  createGraphClient,
  getGraphClient,
  executeGraphRequest,
  email: emailApi,
  calendar: calendarApi,
  folder: folderApi,
  rules: rulesApi
}; 