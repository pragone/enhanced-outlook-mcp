const config = require('../config');
const logger = require('../utils/logger');
const { createGraphClient } = require('../utils/graph-api-adapter');
const { listUsers } = require('../auth/token-manager');
const { buildQueryParams } = require('../utils/odata-helpers');
const { normalizeParameters } = require('../utils/parameter-helpers');

/**
 * List calendar events
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - List of events
 */
async function listEventsHandler(params = {}) {
  // Normalize parameters
  const normalizedParams = normalizeParameters(params);
  let userId = normalizedParams.userId;
  
  if (!userId) {
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
  const calendarId = normalizedParams.calendarId || 'primary';
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
    
    const graphClient = await createGraphClient(userId);
    
    // Determine the endpoint based on calendar ID
    let endpoint;
    if (calendarId === 'primary') {
      endpoint = '/me/calendar/calendarView';
    } else {
      endpoint = `/me/calendars/${calendarId}/calendarView`;
    }
    
    // Build query parameters
    const queryParams = buildQueryParams({
      select: normalizedParams.fields || config.calendar.defaultFields,
      top: limit,
      orderBy: 'start/dateTime asc',
      filter: normalizedParams.filter
    });
    
    // Add start and end time parameters
    queryParams.startDateTime = startDateTime;
    queryParams.endDateTime = endDateTime;
    
    // Get events
    const events = await graphClient.getPaginated(endpoint, queryParams, {
      maxPages: normalizedParams.maxPages || 1
    });
    
    return {
      content: [{
        type: "text",
        text: JSON.stringify({
          status: 'success',
          startDateTime,
          endDateTime,
          calendarId,
          count: events.length,
          events: events.map(formatEventResponse)
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
  // Normalize parameters
  const normalizedParams = normalizeParameters(params);
  let userId = normalizedParams.userId;
  
  if (!userId) {
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
    
    const graphClient = await createGraphClient(userId);
    
    // Get the event
    const event = await graphClient.get(`/me/events/${eventId}`);
    
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
  // Normalize parameters
  const normalizedParams = normalizeParameters(params);
  let userId = normalizedParams.userId;
  
  if (!userId) {
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
    
    const graphClient = await createGraphClient(userId);
    
    // Get all calendars
    const response = await graphClient.get('/me/calendars');
    
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