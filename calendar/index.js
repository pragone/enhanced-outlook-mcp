const { listEventsHandler, getEventHandler, listCalendarsHandler } = require('./list-events');
const { createEventHandler } = require('./create-event');
const { updateEventHandler, respondToEventHandler } = require('./update-event');
const { deleteEventHandler, cancelEventHandler, findMeetingTimesHandler } = require('./delete-event');

// Calendar tool definitions
const calendarTools = [
  // List and Get
  {
    name: 'list_events',
    description: 'List calendar events within a date range',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional, defaults to primary calendar)'
        },
        startDateTime: {
          type: 'string',
          description: 'Start date and time in ISO format (optional, defaults to first day of current month)'
        },
        endDateTime: {
          type: 'string',
          description: 'End date and time in ISO format (optional, defaults to last day of current month)'
        },
        limit: {
          type: 'number',
          description: 'Maximum number of events to return (optional)'
        },
        fields: {
          type: ['array', 'string'],
          description: 'Fields to include in the response (optional)'
        },
        filter: {
          type: 'object',
          description: 'OData filter criteria (optional)'
        },
        maxPages: {
          type: 'number',
          description: 'Maximum number of pages to fetch (optional)'
        }
      }
    },
    handler: listEventsHandler
  },
  {
    name: 'get_event',
    description: 'Get details of a specific calendar event',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        eventId: {
          type: 'string',
          description: 'Event ID to get'
        }
      },
      required: ['eventId']
    },
    handler: getEventHandler
  },
  {
    name: 'list_calendars',
    description: 'List available calendars',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        }
      }
    },
    handler: listCalendarsHandler
  },
  
  // Create
  {
    name: 'create_event',
    description: 'Create a new calendar event',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        calendarId: {
          type: 'string',
          description: 'Calendar ID (optional, defaults to primary calendar)'
        },
        subject: {
          type: 'string',
          description: 'Event subject'
        },
        body: {
          type: 'string',
          description: 'Event body content (optional)'
        },
        bodyType: {
          type: 'string',
          enum: ['Text', 'HTML'],
          description: 'Body content type (optional, defaults to HTML)'
        },
        start: {
          type: 'string',
          description: 'Start time in ISO format'
        },
        end: {
          type: 'string',
          description: 'End time in ISO format (optional, defaults to 30 minutes after start)'
        },
        timeZone: {
          type: 'string',
          description: 'Time zone (optional, defaults to UTC)'
        },
        location: {
          type: ['string', 'object'],
          description: 'Event location (optional)'
        },
        attendees: {
          type: ['string', 'array'],
          description: 'Event attendees (optional)'
        },
        isAllDay: {
          type: 'boolean',
          description: 'Whether this is an all-day event (optional)'
        },
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Whether this is an online meeting (optional)'
        },
        onlineMeetingProvider: {
          type: 'string',
          description: 'Online meeting provider (optional, defaults to teamsForBusiness)'
        },
        categories: {
          type: 'array',
          items: {
            type: 'string'
          },
          description: 'Event categories (optional)'
        },
        sensitivity: {
          type: 'string',
          enum: ['normal', 'personal', 'private', 'confidential'],
          description: 'Event sensitivity (optional, defaults to normal)'
        },
        showAs: {
          type: 'string',
          enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'],
          description: 'Show as status (optional, defaults to busy)'
        }
      },
      required: ['subject', 'start']
    },
    handler: createEventHandler
  },
  
  // Update
  {
    name: 'update_event',
    description: 'Update an existing calendar event',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        eventId: {
          type: 'string',
          description: 'Event ID to update'
        },
        subject: {
          type: 'string',
          description: 'New event subject (optional)'
        },
        body: {
          type: 'string',
          description: 'New event body content (optional)'
        },
        bodyType: {
          type: 'string',
          enum: ['Text', 'HTML'],
          description: 'Body content type (optional)'
        },
        start: {
          type: 'string',
          description: 'New start time in ISO format (optional)'
        },
        end: {
          type: 'string',
          description: 'New end time in ISO format (optional)'
        },
        timeZone: {
          type: 'string',
          description: 'New time zone (optional)'
        },
        location: {
          type: ['string', 'object'],
          description: 'New event location (optional)'
        },
        attendees: {
          type: ['string', 'array'],
          description: 'New event attendees (optional)'
        },
        isAllDay: {
          type: 'boolean',
          description: 'Whether this is an all-day event (optional)'
        },
        isOnlineMeeting: {
          type: 'boolean',
          description: 'Whether this is an online meeting (optional)'
        },
        onlineMeetingProvider: {
          type: 'string',
          description: 'Online meeting provider (optional)'
        },
        categories: {
          type: 'array',
          items: {
            type: 'string'
          },
          description: 'Event categories (optional)'
        },
        sensitivity: {
          type: 'string',
          enum: ['normal', 'personal', 'private', 'confidential'],
          description: 'Event sensitivity (optional)'
        },
        showAs: {
          type: 'string',
          enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'],
          description: 'Show as status (optional)'
        }
      },
      required: ['eventId']
    },
    handler: updateEventHandler
  },
  {
    name: 'respond_to_event',
    description: 'Respond to a calendar event invitation',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        eventId: {
          type: 'string',
          description: 'Event ID to respond to'
        },
        response: {
          type: 'string',
          enum: ['accept', 'tentativelyAccept', 'decline'],
          description: 'Response type'
        },
        comment: {
          type: 'string',
          description: 'Optional comment with the response (optional)'
        }
      },
      required: ['eventId', 'response']
    },
    handler: respondToEventHandler
  },
  
  // Delete and Cancel
  {
    name: 'delete_event',
    description: 'Delete a calendar event',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        eventId: {
          type: 'string',
          description: 'Event ID to delete'
        }
      },
      required: ['eventId']
    },
    handler: deleteEventHandler
  },
  {
    name: 'cancel_event',
    description: 'Cancel a calendar event and notify attendees',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        eventId: {
          type: 'string',
          description: 'Event ID to cancel'
        },
        comment: {
          type: 'string',
          description: 'Cancellation message (optional)'
        }
      },
      required: ['eventId']
    },
    handler: cancelEventHandler
  },
  
  // Scheduling
  {
    name: 'find_meeting_times',
    description: 'Find available meeting times for attendees',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        attendees: {
          type: ['string', 'array'],
          description: 'Meeting attendees'
        },
        startDateTime: {
          type: 'string',
          description: 'Start of time range in ISO format (optional, defaults to now)'
        },
        endDateTime: {
          type: 'string',
          description: 'End of time range in ISO format (optional, defaults to 7 days from now)'
        },
        duration: {
          type: 'string',
          description: 'Meeting duration in ISO8601 format (optional, defaults to PT30M - 30 minutes)'
        },
        timeZone: {
          type: 'string',
          description: 'Time zone (optional, defaults to UTC)'
        },
        locations: {
          type: 'array',
          items: {
            type: ['string', 'object']
          },
          description: 'Potential meeting locations (optional)'
        },
        isLocationRequired: {
          type: 'boolean',
          description: 'Whether a location is required (optional)'
        },
        suggestLocation: {
          type: 'boolean',
          description: 'Whether to suggest locations (optional)'
        },
        minimumAttendeePercentage: {
          type: 'number',
          description: 'Minimum percentage of attendees required (optional, defaults to 100)'
        }
      },
      required: ['attendees']
    },
    handler: findMeetingTimesHandler
  }
];

module.exports = calendarTools;