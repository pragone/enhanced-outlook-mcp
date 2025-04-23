const { listEventsHandler, getEventHandler, listCalendarsHandler } = require('./list-events');
const { createEventHandler } = require('./create-event');
const { updateEventHandler, respondToEventHandler } = require('./update-event');
const { deleteEventHandler, cancelEventHandler, findMeetingTimesHandler } = require('./delete-event');

// Export all handlers directly
module.exports = {
  // List and Get
  listEventsHandler,
  getEventHandler,
  listCalendarsHandler,
  
  // Create
  createEventHandler,
  
  // Update
  updateEventHandler,
  respondToEventHandler,
  
  // Delete
  deleteEventHandler,
  cancelEventHandler,
  findMeetingTimesHandler
};