const config = require('../config');
const logger = require('../utils/logger');
const { createGraphClient, calendar: calendarApi } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const { buildQueryParams } = require('../utils/odata-helpers');
const { normalizeParameters } = require('../utils/parameter-helpers');
const auth = require('../auth/index');

/**
 * Look up calendar ID by name
 * @param {Object} graphClient - Initialized Graph client
 * @param {string} calendarNameOrId - Calendar name or ID
 * @returns {Promise<string>} - Resolved calendar ID
 */
async function resolveCalendarId(graphClient, calendarNameOrId) {
  // Basic logging
  logger.info(`Resolving calendar ID from: ${calendarNameOrId}`);
  
  // If it's 'primary' or looks like a valid ID, use it directly
  if (calendarNameOrId === 'primary' || /^[A-Za-z0-9\-_]{10,}$/.test(calendarNameOrId)) {
    return calendarNameOrId;
  }
  
  // Otherwise, treat it as a name and try to find the matching calendar
  try {
    // Update to use Graph client API correctly - use api() method
    const response = await graphClient.api('/me/calendars').get();
    
    if (!response || !response.value) {
      throw new Error('Failed to retrieve calendars');
    }
    
    const calendar = response.value.find(cal => 
      cal.name.toLowerCase() === calendarNameOrId.toLowerCase()
    );
    
    if (!calendar) {
      throw new Error(`Calendar not found with name: ${calendarNameOrId}`);
    }
    
    logger.info(`Resolved calendar name "${calendarNameOrId}" to ID: ${calendar.id}`);
    return calendar.id;
  } catch (error) {
    throw new Error(`Failed to resolve calendar name: ${error.message}`);
  }
}

/**
 * List calendar events
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of events
 */
async function listEventsHandler(params = {}) {
  // Super detailed debugging for Claude Desktop
  logger.info(`LIST CALENDAR EVENTS HANDLER START ---------------------------`);
  logger.info(`Raw params: ${JSON.stringify(params)}`);
  
  // CRITICAL FIX FOR CLAUDE DESKTOP:
  // Check if this is a direct JSON-RPC request with the tool call format
  if (global.__last_message?.method === 'tools/call' && 
      global.__last_message?.params?.name === 'list_events' &&
      global.__last_message?.params?.arguments) {
    
    const directArgs = global.__last_message.params.arguments;
    logger.info(`Found direct JSON-RPC arguments: ${JSON.stringify(directArgs)}`);
    
    // Use the arguments directly from the JSON-RPC request
    params = directArgs;
    logger.info(`Using direct JSON-RPC arguments for list_events`);
  }
  
  // Check for raw message in params
  const rawMessage = params.__raw_message;
  if (rawMessage) {
    logger.info(`Found raw message in params: ${JSON.stringify(rawMessage)}`);
  }
  
  // Try all possible sources of parameters
  let requestParams = {};
  
  // Order of preference for parameter sources:
  // 1. Direct params
  // 2. params.arguments
  // 3. Raw message params.arguments
  // 4. params.contextData (Claude Desktop might use this)
  // 5. Global last message
  
  if (Object.keys(params).length > 1 || (params.calendarId || params.userId)) {
    // Use direct params
    requestParams = params;
    logger.info(`Using direct params`);
  } else if (params.arguments) {
    // Use params.arguments
    requestParams = typeof params.arguments === 'object' ? params.arguments : params;
    logger.info(`Using params.arguments`);
    
    // Check if arguments is a string containing JSON
    if (typeof params.arguments === 'string') {
      try {
        const parsedArgs = JSON.parse(params.arguments);
        if (parsedArgs && typeof parsedArgs === 'object') {
          requestParams = parsedArgs;
          logger.info(`Parsed string arguments into object: ${JSON.stringify(parsedArgs)}`);
        }
      } catch (e) {
        logger.info(`Arguments string is not valid JSON: ${params.arguments}`);
      }
    }
  } else if (rawMessage?.params?.arguments) {
    // Use raw message params
    requestParams = rawMessage.params.arguments;
    logger.info(`Using raw message params`);
  } else if (params.contextData) {
    // Try to use contextData (sometimes used by Claude Desktop)
    requestParams = params.contextData;
    logger.info(`Using params.contextData: ${JSON.stringify(params.contextData)}`);
  } else if (global.__last_message?.params?.arguments) {
    // Use global last message
    requestParams = global.__last_message.params.arguments;
    logger.info(`Using global last message params`);
  } else {
    // Last resort - use defaults
    requestParams = params;
    logger.info(`Using default params`);
  }
  
  logger.info(`Request params extracted: ${JSON.stringify(requestParams)}`);
  
  // Normalize parameters
  const normalizedParams = normalizeParameters(requestParams);
  let userId = normalizedParams.userId;
  
  if (!userId) {
    // Check authentication status with the unified auth FIRST
    try {
      const authStatusResult = await auth.checkAuthStatusHandler();
      const authStatusData = JSON.parse(authStatusResult.content[0].text);
      
      if (authStatusData.status === 'authenticated' && authStatusData.user?.email) {
        userId = authStatusData.user.email;
        logger.info(`Using authenticated user ID from session: ${userId}`);
      } else {
        // Fall back to old method
        const users = await listUsers();
        if (users.length === 0) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: 'error',
                message: 'No authenticated users found. Please authenticate first.'
              })
            }]
          };
        }
        userId = users.length === 1 ? users[0] : normalizedParams.userId;
      }
    } catch (authError) {
      logger.error(`Error checking auth status: ${authError.message}`);
      // Fall back to old method
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Please authenticate first.'
            })
          }]
        };
      }
      userId = users.length === 1 ? users[0] : normalizedParams.userId;
    }
    
    if (!userId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Multiple users found. Please specify userId parameter.'
          })
        }]
      };
    }
  }
  
  // Use normalized parameters
  const providedCalendarId = normalizedParams.calendarId || 'primary';
  let startDateTime = normalizedParams.startDateTime;
  let endDateTime = normalizedParams.endDateTime;
  
  if (!startDateTime) {
    const now = new Date();
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    startDateTime = firstDay.toISOString();
  }
  
  if (!endDateTime) {
    const now = new Date();
    const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    endDateTime = lastDay.toISOString();
  }
  
  const limit = Math.min(
    normalizedParams.limit || config.calendar.maxEventsPerRequest, 
    config.calendar.maxEventsPerRequest
  );
  
  try {
    logger.info(`Listing calendar events for user ${userId} from ${startDateTime} to ${endDateTime}`);
    
    // Get the Graph client
    const graphClient = await auth.getGraphClient(userId);
    
    // Resolve calendar ID if it's a name
    // This functionality allows users to specify a calendar by name instead of ID
    let resolvedCalendarId;
    try {
      resolvedCalendarId = await resolveCalendarId(graphClient, providedCalendarId);
      logger.info(`Using resolved calendar ID: ${resolvedCalendarId}`);
    } catch (error) {
      logger.warn(`Failed to resolve calendar ID by name: ${error.message}`);
      resolvedCalendarId = providedCalendarId;
    }
    
    // CRITICAL CHANGE: Use calendarApi instead of directly creating a graph client
    // First get all events using the API (this correctly uses auth.getGraphClient)
    const apiOptions = {
      calendarId: resolvedCalendarId === 'primary' ? undefined : resolvedCalendarId,
      top: limit,
      filter: normalizedParams.filter,
      orderBy: 'start/dateTime asc'
    };
    
    // First get the calendar events - this will correctly use the authentication token
    const eventsResponse = await calendarApi.listEvents(userId, apiOptions);
    
    // Filter events by date range since the calendarView endpoint is not used in the API
    const filteredEvents = eventsResponse.value.filter(event => {
      const eventStart = new Date(event.start.dateTime + 'Z');
      const eventEnd = new Date(event.end.dateTime + 'Z');
      const rangeStart = new Date(startDateTime);
      const rangeEnd = new Date(endDateTime);
      
      return (eventStart >= rangeStart && eventStart <= rangeEnd) || 
             (eventEnd >= rangeStart && eventEnd <= rangeEnd) ||
             (eventStart <= rangeStart && eventEnd >= rangeEnd);
    });
    
    // Apply limit after filtering
    const limitedEvents = filteredEvents.slice(0, limit);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          startDateTime,
          endDateTime,
          calendarId: providedCalendarId,
          count: limitedEvents.length,
          events: limitedEvents.map(formatEventResponse)
        })
      }]
    };
  } catch (error) {
    logger.error(`Error listing calendar events: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list calendar events: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Get a specific calendar event
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Event details
 */
async function getEventHandler(params = {}) {
  // Add similar param processing as listEventsHandler
  logger.info(`GET CALENDAR EVENT HANDLER START ---------------------------`);
  logger.info(`Raw params: ${JSON.stringify(params)}`);
  
  // Process params for Claude Desktop
  let requestParams = {};
  if (global.__last_message?.method === 'tools/call' && 
      global.__last_message?.params?.name === 'get_event' &&
      global.__last_message?.params?.arguments) {
    requestParams = global.__last_message.params.arguments;
    logger.info(`Using direct JSON-RPC arguments for get_event`);
  } else if (params.arguments) {
    requestParams = typeof params.arguments === 'object' ? params.arguments : params;
  } else {
    requestParams = params;
  }
  
  // Normalize parameters
  const normalizedParams = normalizeParameters(requestParams);
  let userId = normalizedParams.userId;
  
  if (!userId) {
    // Check authentication status with the unified auth
    try {
      const authStatusResult = await auth.checkAuthStatusHandler();
      const authStatusData = JSON.parse(authStatusResult.content[0].text);
      
      if (authStatusData.status === 'authenticated' && authStatusData.user?.email) {
        userId = authStatusData.user.email;
        logger.info(`Using authenticated user ID from session: ${userId}`);
      } else {
        // Fall back to old method
        const users = await listUsers();
        if (users.length === 0) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: 'error',
                message: 'No authenticated users found. Please authenticate first.'
              })
            }]
          };
        }
        userId = users.length === 1 ? users[0] : normalizedParams.userId;
      }
    } catch (authError) {
      logger.error(`Error checking auth status: ${authError.message}`);
      // Fall back to old method
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Please authenticate first.'
            })
          }]
        };
      }
      userId = users.length === 1 ? users[0] : normalizedParams.userId;
    }
    
    if (!userId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Multiple users found. Please specify userId parameter.'
          })
        }]
      };
    }
  }
  
  const eventId = normalizedParams.eventId;
  const calendarId = normalizedParams.calendarId || 'primary';
  
  if (!eventId) {
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: 'Event ID is required'
        })
      }]
    };
  }
  
  try {
    logger.info(`Getting calendar event ${eventId} for user ${userId}`);
    
    // Get the Graph client for calendar name resolution if needed
    const graphClient = await auth.getGraphClient(userId);
    
    // Resolve calendar ID if it's a name
    // This functionality allows users to specify a calendar by name instead of ID
    let resolvedCalendarId;
    try {
      resolvedCalendarId = await resolveCalendarId(graphClient, calendarId);
      logger.info(`Using resolved calendar ID: ${resolvedCalendarId}`);
    } catch (error) {
      logger.warn(`Failed to resolve calendar ID by name: ${error.message}`);
      resolvedCalendarId = calendarId;
    }
    
    // CRITICAL CHANGE: Use calendarApi instead of directly creating a graph client
    // This will correctly use the cached auth token
    const event = await calendarApi.getEvent(userId, eventId, resolvedCalendarId);
    
    if (!event) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: `Event not found with ID: ${eventId}`
          })
        }]
      };
    }
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          event: formatEventResponse(event)
        })
      }]
    };
  } catch (error) {
    logger.error(`Error getting calendar event: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to get calendar event: ${error.message}`
        })
      }]
    };
  }
}

/**
 * List available calendars
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of calendars
 */
async function listCalendarsHandler(params = {}) {
  // Add similar param processing as listEventsHandler
  logger.info(`LIST CALENDARS HANDLER START ---------------------------`);
  logger.info(`Raw params: ${JSON.stringify(params)}`);
  
  // Process params for Claude Desktop
  let requestParams = {};
  if (global.__last_message?.method === 'tools/call' && 
      global.__last_message?.params?.name === 'list_calendars' &&
      global.__last_message?.params?.arguments) {
    requestParams = global.__last_message.params.arguments;
    logger.info(`Using direct JSON-RPC arguments for list_calendars`);
  } else if (params.arguments) {
    requestParams = typeof params.arguments === 'object' ? params.arguments : params;
  } else {
    requestParams = params;
  }
  
  // Normalize parameters
  const normalizedParams = normalizeParameters(requestParams);
  let userId = normalizedParams.userId;
  
  if (!userId) {
    // Check authentication status with the unified auth
    try {
      const authStatusResult = await auth.checkAuthStatusHandler();
      const authStatusData = JSON.parse(authStatusResult.content[0].text);
      
      if (authStatusData.status === 'authenticated' && authStatusData.user?.email) {
        userId = authStatusData.user.email;
        logger.info(`Using authenticated user ID from session: ${userId}`);
      } else {
        // Fall back to old method
        const users = await listUsers();
        if (users.length === 0) {
          return {
            content: [{
              type: "text",
              text: JSON.stringify({
                status: 'error',
                message: 'No authenticated users found. Please authenticate first.'
              })
            }]
          };
        }
        userId = users.length === 1 ? users[0] : normalizedParams.userId;
      }
    } catch (authError) {
      logger.error(`Error checking auth status: ${authError.message}`);
      // Fall back to old method
      const users = await listUsers();
      if (users.length === 0) {
        return {
          content: [{
            type: "text",
            text: JSON.stringify({
              status: 'error',
              message: 'No authenticated users found. Please authenticate first.'
            })
          }]
        };
      }
      userId = users.length === 1 ? users[0] : normalizedParams.userId;
    }
    
    if (!userId) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Multiple users found. Please specify userId parameter.'
          })
        }]
      };
    }
  }
  
  try {
    logger.info(`Listing calendars for user ${userId}`);
    
    // CRITICAL CHANGE: Use auth.getGraphClient directly instead of createGraphClient
    const graphClient = await auth.getGraphClient(userId);
    
    // Get all calendars - we need to use the base client since there's no dedicated API method
    const response = await graphClient.api('/me/calendars').get();
    
    if (!response || !response.value) {
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            status: 'error',
            message: 'Failed to retrieve calendars'
          })
        }]
      };
    }
    
    // Format the calendars
    const calendars = response.value.map(calendar => ({
      id: calendar.id,
      name: calendar.name,
      color: calendar.color,
      isDefaultCalendar: calendar.isDefaultCalendar,
      canShare: calendar.canShare,
      canViewPrivateItems: calendar.canViewPrivateItems,
      canEdit: calendar.canEdit,
      owner: calendar.owner ? {
        name: calendar.owner.name,
        address: calendar.owner.address
      } : null
    }));
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          count: calendars.length,
          calendars
        })
      }]
    };
  } catch (error) {
    logger.error(`Error listing calendars: ${error.message}`);
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'error',
          message: `Failed to list calendars: ${error.message}`
        })
      }]
    };
  }
}

/**
 * Format event response
 * @param {Object} event - Raw event from Graph API
 * @returns {Object} - Formatted event
 */
function formatEventResponse(event) {
  // Extract organizer
  let organizer = null;
  if (event.organizer && event.organizer.emailAddress) {
    organizer = {
      name: event.organizer.emailAddress.name,
      email: event.organizer.emailAddress.address
    };
  }
  
  // Extract attendees
  let attendees = [];
  if (event.attendees && Array.isArray(event.attendees)) {
    attendees = event.attendees.map(attendee => ({
      type: attendee.type,
      status: attendee.status ? attendee.status.response : 'none',
      name: attendee.emailAddress.name,
      email: attendee.emailAddress.address
    }));
  }
  
  // Extract location
  let location = null;
  if (event.location) {
    location = {
      displayName: event.location.displayName,
      address: event.location.address,
      coordinates: event.location.coordinates
    };
  }
  
  // Create formatted response
  return {
    id: event.id,
    subject: event.subject,
    bodyPreview: event.bodyPreview,
    start: event.start,
    end: event.end,
    location,
    organizer,
    attendees,
    isAllDay: event.isAllDay,
    isCancelled: event.isCancelled,
    sensitivity: event.sensitivity,
    showAs: event.showAs,
    importance: event.importance,
    onlineMeetingUrl: event.onlineMeetingUrl,
    isOnlineMeeting: event.isOnlineMeeting || false,
    categories: event.categories || [],
    webLink: event.webLink
  };
}

module.exports = {
  listEventsHandler,
  getEventHandler,
  listCalendarsHandler
};