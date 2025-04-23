const config = require('../config');
const logger = require('../utils/logger');
const { GraphApiClient } = require('../utils/graph-api');
const { listUsers } = require('../auth/token-manager');

/**
 * Update an existing calendar event
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Update result
 */
async function updateEventHandler(params = {}) {
  let userId = params.userId;
  if (!userId) {
    const users = await listUsers();
    if (users.length === 0) {
      return formatMcpResponse({
        status: 'error',
        message: 'No authenticated users found. Please authenticate first.'
      });
    }
    userId = users.length === 1 ? users[0] : params.userId;
    if (!userId) {
      return formatMcpResponse({
        status: 'error',
        message: 'Multiple users found. Please specify userId parameter.'
      });
    }
  }
  const eventId = params.eventId;
  
  if (!eventId) {
    return formatMcpResponse({
      status: 'error',
      message: 'Event ID is required'
    });
  }
  
  try {
    logger.info(`Updating calendar event ${eventId} for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Prepare update data
    const updateData = {};
    
    if (params.subject) {
      updateData.subject = params.subject;
    }
    
    if (params.body) {
      updateData.body = {
        contentType: params.bodyType || 'HTML',
        content: params.body
      };
    }
    
    if (params.start) {
      updateData.start = {
        dateTime: params.start,
        timeZone: params.timeZone || 'UTC'
      };
    }
    
    if (params.end) {
      updateData.end = {
        dateTime: params.end,
        timeZone: params.timeZone || 'UTC'
      };
    }
    
    if (params.location) {
      updateData.location = formatLocation(params.location);
    }
    
    if (params.attendees) {
      updateData.attendees = formatAttendees(params.attendees);
    }
    
    if (params.isAllDay !== undefined) {
      updateData.isAllDay = params.isAllDay;
    }
    
    if (params.sensitivity) {
      updateData.sensitivity = params.sensitivity;
    }
    
    if (params.showAs) {
      updateData.showAs = params.showAs;
    }
    
    if (params.isOnlineMeeting !== undefined) {
      updateData.isOnlineMeeting = params.isOnlineMeeting;
      
      if (params.isOnlineMeeting && params.onlineMeetingProvider) {
        updateData.onlineMeetingProvider = params.onlineMeetingProvider;
      }
    }
    
    if (params.categories) {
      updateData.categories = params.categories;
    }
    
    // Update the event
    await graphClient.patch(`/me/events/${eventId}`, updateData);
    
    return formatMcpResponse({
      status: 'success',
      message: 'Event updated successfully',
      eventId
    });
  } catch (error) {
    logger.error(`Error updating calendar event: ${error.message}`);
    
    return formatMcpResponse({
      status: 'error',
      message: `Failed to update calendar event: ${error.message}`
    });
  }
}

/**
 * Format location data for API
 * @param {string|Object} location - Location information
 * @returns {Object} - Formatted location
 */
function formatLocation(location) {
  if (!location) {
    return null;
  }
  
  // If it's already in the correct format
  if (typeof location === 'object' && location.displayName) {
    return location;
  }
  
  // If it's a string, use as display name
  if (typeof location === 'string') {
    return {
      displayName: location
    };
  }
  
  // If it's an object with specific properties
  if (typeof location === 'object') {
    return {
      displayName: location.name || location.displayName || 'Unknown Location',
      address: location.address,
      coordinates: location.coordinates
    };
  }
  
  // Default case
  return {
    displayName: String(location)
  };
}

/**
 * Format attendees for API
 * @param {Array|string|Object} attendees - Attendees in various formats
 * @returns {Array} - Formatted attendees
 */
function formatAttendees(attendees) {
  if (!attendees) {
    return [];
  }
  
  // Handle string with comma or semicolon separators
  if (typeof attendees === 'string') {
    attendees = attendees.split(/[,;]/).map(a => a.trim()).filter(Boolean);
  }
  
  // Ensure it's an array
  if (!Array.isArray(attendees)) {
    attendees = [attendees];
  }
  
  // Format each attendee
  return attendees.map(attendee => {
    // If already in the correct format
    if (typeof attendee === 'object' && attendee.emailAddress) {
      return attendee;
    }
    
    // Handle string in format "Name <email@example.com>"
    if (typeof attendee === 'string') {
      const match = attendee.match(/^(.*?)\s*<([^>]+)>$/);
      if (match) {
        return {
          emailAddress: {
            name: match[1].trim(),
            address: match[2].trim()
          },
          type: 'required'
        };
      }
      
      // Just an email address
      return {
        emailAddress: {
          address: attendee.trim()
        },
        type: 'required'
      };
    }
    
    // Handle object with name and email properties
    if (typeof attendee === 'object' && attendee.email) {
      return {
        emailAddress: {
          name: attendee.name || '',
          address: attendee.email
        },
        type: attendee.type || 'required'
      };
    }
    
    // Default case
    return {
      emailAddress: {
        address: String(attendee)
      },
      type: 'required'
    };
  });
}

/**
 * Respond to an event
 * @param {Object} params - Tool parameters
 * @returns {Promise<Object>} - Response result
 */
async function respondToEventHandler(params = {}) {
  let userId = params.userId;
  if (!userId) {
    const users = await listUsers();
    if (users.length === 0) {
      return formatMcpResponse({
        status: 'error',
        message: 'No authenticated users found. Please authenticate first.'
      });
    }
    userId = users.length === 1 ? users[0] : params.userId;
    if (!userId) {
      return formatMcpResponse({
        status: 'error',
        message: 'Multiple users found. Please specify userId parameter.'
      });
    }
  }
  const eventId = params.eventId;
  const response = params.response;
  
  if (!eventId) {
    return formatMcpResponse({
      status: 'error',
      message: 'Event ID is required'
    });
  }
  
  if (!response || !['accept', 'tentativelyAccept', 'decline'].includes(response)) {
    return formatMcpResponse({
      status: 'error',
      message: 'Valid response is required (accept, tentativelyAccept, or decline)'
    });
  }
  
  try {
    logger.info(`Responding to calendar event ${eventId} with "${response}" for user ${userId}`);
    
    const graphClient = new GraphApiClient(userId);
    
    // Create response data
    const responseData = {
      comment: params.comment || ''
    };
    
    // Send the response
    await graphClient.post(`/me/events/${eventId}/${response}`, responseData);
    
    return formatMcpResponse({
      status: 'success',
      message: `Event ${getResponseText(response)} successfully`,
      eventId
    });
  } catch (error) {
    logger.error(`Error responding to calendar event: ${error.message}`);
    
    return formatMcpResponse({
      status: 'error',
      message: `Failed to respond to calendar event: ${error.message}`
    });
  }
}

/**
 * Get human-readable response text
 * @param {string} response - Response type
 * @returns {string} - Human-readable text
 */
function getResponseText(response) {
  switch (response) {
    case 'accept':
      return 'accepted';
    case 'tentativelyAccept':
      return 'tentatively accepted';
    case 'decline':
      return 'declined';
    default:
      return 'responded to';
  }
}

/**
 * Format response for MCP
 * @param {Object} data - Response data
 * @returns {Object} - MCP formatted response
 */
function formatMcpResponse(data) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(data)
      }
    ]
  };
}

module.exports = {
  updateEventHandler,
  respondToEventHandler
};