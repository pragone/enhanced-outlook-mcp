// rules/delete.js
const { createGraphClient } = require('../utils/graph-api-adapter');

async function deleteRuleHandler(params = {}) {
  const { userId = 'default', ruleId } = params;
  if (!ruleId) return formatMcpResponse({ status: 'error', message: 'Rule ID required' });

  try {
    const graphClient = await createGraphClient(userId);
    await graphClient.delete(`/me/mailFolders/inbox/messageRules/${ruleId}`);
    return formatMcpResponse({ status: 'success', message: 'Rule deleted', ruleId });
  } catch (error) {
    return formatMcpResponse({ status: 'error', message: `Failed: ${error.message}` });
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

module.exports = { deleteRuleHandler };
