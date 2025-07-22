# Shared Utilities

This directory contains shared utilities and common code used across different Microsoft service integrations in the monolith.

## Structure

- `utils/` - Common utility functions
  - `common.js` - Core utilities for logging, configuration, Office.js helpers, and Microsoft Graph API

## Usage

Import shared utilities in your integration projects:

```javascript
const { Logger, Config, OfficeHelper, GraphApiHelper } = require('../../shared/utils/common');

// Use consistent logging across integrations
Logger.info('Starting Word integration', 'WordAddin');

// Get environment-specific configuration
const serverUrl = Config.getServerUrl(3000);

// Handle Office.js operations with error handling
await OfficeHelper.runWithErrorHandling(async () => {
    // Your Office.js code here
});
```

## Available Utilities

### Logger
- Consistent logging with timestamps and context
- Methods: `info()`, `warn()`, `error()`, `debug()`

### Config
- Environment detection and configuration
- Server URL generation for dev/prod environments

### OfficeHelper
- Office.js initialization helpers
- Error handling wrappers for Office operations
- Error message formatting

### GraphApiHelper
- Microsoft Graph API request helpers
- Authentication header generation
- Standardized error handling
