// rules/delete.js
const { GraphApiClient } = require('../utils/graph-api');

async function deleteRuleHandler(params = {}) {
  const { userId = 'default', ruleId } = params;
  if (!ruleId) return { status: 'error', message: 'Rule ID required' };

  try {
    const graphClient = new GraphApiClient(userId);
    await graphClient.delete(`/me/mailFolders/inbox/messageRules/${ruleId}`);
    return { status: 'success', message: 'Rule deleted', ruleId };
  } catch (error) {
    return { status: 'error', message: `Failed: ${error.message}` };
  }
}

module.exports = { deleteRuleHandler };
