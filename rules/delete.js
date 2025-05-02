// rules/delete.js
const { rules: rulesApi } = require('../utils/graph-api-adapter');
const logger = require('../utils/logger');
const auth = require('../auth/index');

async function deleteRuleHandler(params = {}) {
  const { userId = 'default', ruleId } = params;
  if (!ruleId) return formatMcpResponse({ status: 'error', message: 'Rule ID required' });

  try {
    logger.info(`Deleting mail rule ${ruleId} for user ${userId}`);
    
    // Add deleteRule method to rulesApi if it doesn't exist
    if (!rulesApi.deleteRule) {
      // Fallback to direct API call using executeGraphRequest
      const { executeGraphRequest } = require('../utils/graph-api-adapter');
      await executeGraphRequest(userId, 'rules', async (client) => {
        return await client.api(`/me/mailFolders/inbox/messageRules/${ruleId}`).delete();
      });
    } else {
      // Use the API helper if it exists
      await rulesApi.deleteRule(userId, ruleId);
    }
    
    return formatMcpResponse({ status: 'success', message: 'Rule deleted', ruleId });
  } catch (error) {
    logger.error(`Error deleting mail rule: ${error.message}`);
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
