const { listRulesHandler, getRuleHandler } = require('./list');
const { createRuleHandler, updateRuleHandler, deleteRuleHandler } = require('./create');

// Export all handlers directly
module.exports = {
  // List and Get
  listRulesHandler,
  getRuleHandler,
  
  // Create, Update, Delete
  createRuleHandler,
  updateRuleHandler,
  deleteRuleHandler
};