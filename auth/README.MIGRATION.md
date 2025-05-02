# Authentication System Migration

This document outlines the process for migrating from the old authentication system to the new authentication system in the Enhanced Outlook MCP Server.

## Migration Status

Phase 1 (Core Authentication Module) and Phase 2 (Graph Client Integration) have been completed. Phase 3 (Migration) is now being implemented, which allows side-by-side operation of both authentication systems.

## Migration Components

The following components have been implemented to facilitate the migration:

1. **Migration Configuration (`auth/migration.js`)**
   - Feature flags to control which system is used for each functional area
   - Tracking of migrated users
   - Helper functions to control the migration process

2. **Unified Authentication (`auth/unified-index.js`)**
   - Single entry point for authentication tools that routes to either old or new auth system
   - Handles authentication, token refresh, and user management
   - Graceful fallback to old system if new system fails

3. **Graph API Adapter (`utils/graph-api-adapter.js`)**
   - Provides a consistent interface for Microsoft Graph API calls
   - Routes requests through either old or new Graph client based on feature flags
   - Maintains backward compatibility with existing code

4. **Migration Test Suite (`test-auth-migration.js`)**
   - Validates both auth systems side-by-side
   - Tests token refresh and API functionality
   - Configuration tools for feature flags

## Migration Strategy

The migration process follows these steps:

1. **Feature Flag Control**
   - Each functional area (email, calendar, folder, rules) can be individually migrated
   - Global flags control overall auth system and Graph client usage
   - Feature flags are stored persistently between server restarts

2. **Gradual Rollout**
   - New users are automatically directed to the new auth system
   - Existing users can continue using the old system until all features are migrated
   - Individual features can be migrated one by one to minimize risk

3. **Validation and Testing**
   - The test script validates both systems working side-by-side
   - Each feature can be tested with both auth systems before full migration

## How to Use

### Managing Feature Flags

To control which authentication system is used for each feature, use the migration module:

```javascript
const migration = require('./auth/migration');

// Enable new auth system for a specific feature
migration.toggleFeature('email', true);

// Enable new Graph client globally
migration.enableNewGraphClient(true);

// Enable new auth system globally
migration.enableNewAuthSystem(true);
```

### Testing the Migration

Run the migration test script to validate both authentication systems:

```
node test-auth-migration.js
```

The test script provides options to:
- Test both old and new authentication systems
- Test the unified authentication layer
- Test Graph API calls through both systems
- Configure feature flags
- Validate token refresh functionality

### Monitoring Migration Progress

To get the current migration status:

```javascript
const migration = require('./auth/migration');
const status = migration.getMigrationStatus();
console.log(status);
```

The status includes:
- Start date of the migration
- Count of migrated users
- Status of feature flags
- Whether migration is complete

## Completing the Migration

Once all features have been successfully migrated and validated, the final step is to enable the new auth system globally:

```javascript
const migration = require('./auth/migration');
migration.enableNewAuthSystem(true);
migration.enableNewGraphClient(true);
```

After the migration is complete and stable, the old auth system components can be removed in Phase 4 (Cleanup).

## Troubleshooting

If issues occur during migration, you can:

1. Toggle feature flags back to use the old system for specific features
2. Check the logs for specific authentication errors
3. Run the test script to validate both systems
4. Revert to the old auth system completely if necessary

## Notes for Developers

When adding new features or modifying existing ones:

1. Always use the unified auth module:
   ```javascript
   const auth = require('./auth/unified-index');
   ```

2. Use the Graph API adapter for Microsoft Graph API calls:
   ```javascript
   const graphAdapter = require('./utils/graph-api-adapter');
   ```

3. When making Graph API calls, specify the feature area:
   ```javascript
   const client = await graphAdapter.getGraphClient(userId, 'email');
   // Or use the helper methods
   const messages = await graphAdapter.email.listMessages(userId, options);
   ``` 