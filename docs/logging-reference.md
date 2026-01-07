# MCP Microsoft Office - Logging Reference Strategy

## Required Logging Patterns

Every controller function must implement these 4 patterns to ensure comprehensive logging coverage:

### Pattern Implementation Checklist

- [ ] 1. Development Debug Logs - Conditional on NODE_ENV
- [ ] 2. User Activity Logs - For successful operations
- [ ] 3. Infrastructure Error Logging - For server operations
- [ ] 4. User Error Tracking - For user-visible errors

## Reference Implementation

```javascript
async function exampleFunction(req, res) {
  try {
    // Pattern 1: Development Debug Logs
    // Only emitted in development environment
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Action description', {
        sessionId: req.session?.id,
        userAgent: req.get('User-Agent'),
        timestamp: new Date().toISOString()
        // Include relevant context data
      }, 'category'); // Always specify category (e.g., 'auth', 'mail', etc.)
    }
    
    // ... business logic ...
    
    // Pattern 2: User Activity Logs
    // Always log successful user actions with context
    const userId = req?.user?.userId;
    if (userId) {
      MonitoringService.info('Action completed successfully', {
        // Include relevant action details
        timestamp: new Date().toISOString()
      }, 'category', null, userId); // Include category and userId
    } else if (req.session?.id) {
      // Fallback to session ID if user ID not available
      MonitoringService.info('Action completed with session', {
        sessionId: req.session.id,
        timestamp: new Date().toISOString()
      }, 'category');
    }
    
    return res.json({ success: true });
    
  } catch (error) {
    // Pattern 3: Infrastructure Error Logging
    // Create structured error and log for operations team
    const mcpError = ErrorService.createError(
      'category', // Error category (e.g., 'auth', 'api', 'graph')
      'Descriptive error message', // Human-readable message
      'error', // Severity: 'error', 'warn', etc.
      { 
        endpoint: '/api/path',
        error: error.message,
        // Additional context for debugging
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    // Log errors with user context for visibility in user logs
    const userId = req?.user?.userId;
    if (userId) {
      MonitoringService.error('Action failed', {
        error: error.message,
        timestamp: new Date().toISOString()
      }, 'category', null, userId); // Include category and userId
    } else if (req.session?.id) {
      // Fallback to session ID if user ID not available
      MonitoringService.error('Action failed', {
        sessionId: req.session.id,
        error: error.message,
        timestamp: new Date().toISOString()
      }, 'category');
    }
    
    return res.status(500).json({
      error: 'error_code',
      error_description: 'User-friendly error message'
    });
  }
}
```