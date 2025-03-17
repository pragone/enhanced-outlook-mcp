const { listRulesHandler, getRuleHandler } = require('./list');
const { createRuleHandler, updateRuleHandler, deleteRuleHandler } = require('./create');

// Rule tool definitions
const ruleTools = [
  // List and Get
  {
    name: 'list_rules',
    description: 'List mail rules',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        }
      }
    },
    handler: listRulesHandler
  },
  {
    name: 'get_rule',
    description: 'Get details of a specific mail rule',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        ruleId: {
          type: 'string',
          description: 'Rule ID to get'
        }
      },
      required: ['ruleId']
    },
    handler: getRuleHandler
  },
  
  // Create, Update, Delete
  {
    name: 'create_rule',
    description: 'Create a new mail rule',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        displayName: {
          type: 'string',
          description: 'Display name for the rule'
        },
        sequence: {
          type: 'number',
          description: 'Sequence number for rule execution order (optional, defaults to 0)'
        },
        isEnabled: {
          type: 'boolean',
          description: 'Whether the rule is enabled (optional, defaults to true)'
        },
        conditions: {
          type: 'object',
          description: 'Rule conditions'
        },
        actions: {
          type: 'object',
          description: 'Rule actions'
        }
      },
      required: ['displayName', 'conditions', 'actions']
    },
    handler: createRuleHandler
  },
  {
    name: 'update_rule',
    description: 'Update an existing mail rule',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        ruleId: {
          type: 'string',
          description: 'Rule ID to update'
        },
        displayName: {
          type: 'string',
          description: 'New display name for the rule (optional)'
        },
        sequence: {
          type: 'number',
          description: 'New sequence number for rule execution order (optional)'
        },
        isEnabled: {
          type: 'boolean',
          description: 'Whether the rule is enabled (optional)'
        },
        conditions: {
          type: 'object',
          description: 'New rule conditions (optional)'
        },
        actions: {
          type: 'object',
          description: 'New rule actions (optional)'
        }
      },
      required: ['ruleId']
    },
    handler: updateRuleHandler
  },
  {
    name: 'delete_rule',
    description: 'Delete a mail rule',
    parameters: {
      type: 'object',
      properties: {
        userId: {
          type: 'string',
          description: 'User identifier (optional, defaults to "default")'
        },
        ruleId: {
          type: 'string',
          description: 'Rule ID to delete'
        }
      },
      required: ['ruleId']
    },
    handler: deleteRuleHandler
  }
];

module.exports = ruleTools;