const { listEmailsHandler } = require('./list');
const { searchEmailsHandler } = require('./search');
const { readEmailHandler, markEmailHandler } = require('./read');
const { 
  sendEmailHandler, 
  createDraftHandler, 
  replyEmailHandler, 
  forwardEmailHandler 
} = require('./send');
const { 
  getAttachmentHandler, 
  listAttachmentsHandler, 
  addAttachmentHandler, 
  deleteAttachmentHandler 
} = require('./attachments');

// Export all handlers directly
module.exports = {
  // List and Search
  listEmailsHandler,
  searchEmailsHandler,
  
  // Read
  readEmailHandler,
  markEmailHandler,
  
  // Send and Reply
  sendEmailHandler,
  createDraftHandler,
  replyEmailHandler,
  forwardEmailHandler,
  
  // Attachments
  getAttachmentHandler,
  listAttachmentsHandler,
  addAttachmentHandler,
  deleteAttachmentHandler
};