/**
 * @fileoverview New Event-Based MonitoringService with circular buffer implementation.
 * Replaces direct logging with event emission and uses circular buffer to prevent memory growth.
 * Maintains backward compatibility with existing monitoring service API.
 */

const winston = require('winston');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

// Import ErrorService for error creation
const ErrorService = require('./error-service.cjs');

// Lazy load storage service to avoid circular dependency
let storageService = null;
function getStorageService() {
  if (!storageService) {
    try {
      storageService = require('./storage-service.cjs');
    } catch (error) {
      console.warn('[MONITORING] Could not load storage service:', error.message);
    }
  }
  return storageService;
}

/**
 * Determines if a log message is infrastructure-related and should not be shown to users
 * @param {string} message - The log message
 * @param {string} category - The log category
 * @returns {boolean} True if this is an infrastructure log
 */
function isInfrastructureLog(message, category) {
  // In development mode, show all logs for debugging purposes
  if (process.env.NODE_ENV === 'development') {
    return false;
  }
  
  // Infrastructure log patterns to exclude from user logs in production
  const infrastructurePatterns = [
    /^Request entered (routes|controller)$/,
    /^Request exiting (routes|controller) - \d+$/,
    /^Processing (GET|POST|PUT|DELETE|PATCH) /,
    /^Fetching status/,
    /^Connection status updated/,
    /^Started (periodic status refresh|log auto-refresh)/,
    /^Authentication status/,
    /^User (not )?authenticated/,
    /^Session-based authentication/,
    /^Microsoft 365 session authentication/,
    /^API call authenticated/,
    /^JWT token validated/,
    /^Database connection acquired/,
    /^Started log auto-refresh/,
    /^Fetching logs/
  ];
  
  // Check if message matches any infrastructure pattern
  return infrastructurePatterns.some(pattern => pattern.test(message));
}

// Import event service for event-based architecture (lazy loaded to avoid circular dependency)
let eventService = null;

// Event types based on analysis
const eventTypes = {
  ERROR: 'log:error',
  INFO: 'log:info', 
  WARN: 'log:warn',
  DEBUG: 'log:debug',
  METRIC: 'log:metric',
  SYSTEM_MEMORY_WARNING: 'system:memory:warning',
  SYSTEM_EMERGENCY: 'system:emergency'
};

// Circular buffer implementation for memory-safe log storage
class CircularBuffer {
  constructor(size = 100) {
    this.size = size;
    this.buffer = [];
    this.currentIndex = 0;
  }
  
  add(item) {
    if (this.buffer.length < this.size) {
      this.buffer.push(item);
    } else {
      this.buffer[this.currentIndex] = item;
    }
    this.currentIndex = (this.currentIndex + 1) % this.size;
    return item;
  }
  
  getAll() {
    if (this.buffer.length < this.size) {
      // Buffer not full yet, return in insertion order
      return [...this.buffer];
    } else {
      // Buffer is full, need to return in correct chronological order
      // Items from currentIndex to end are oldest, items from 0 to currentIndex-1 are newest
      const newerItems = this.buffer.slice(0, this.currentIndex);
      const olderItems = this.buffer.slice(this.currentIndex);
      return [...olderItems, ...newerItems];
    }
  }
  
  clear() {
    this.buffer = [];
    this.currentIndex = 0;
  }
}

// Initialize circular buffer with size based on memory analysis
const logBuffer = new CircularBuffer(100);

// Copy Winston configuration and constants from original service
let logger = null;
const dateSuffix = new Date().toISOString().slice(0,10).replace(/-/g, '');
let LOG_FILE_PATH = process.env.MCP_LOG_PATH || path.join(__dirname, `../../logs/mcp${dateSuffix}.log`);

// Read version from package.json
let appVersion = 'unknown';
try {
    const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, '../../package.json'), 'utf8'));
    appVersion = pkg.version || 'unknown';
} catch (e) {}

// Memory monitoring constants from original
const MEMORY_CHECK_INTERVAL = 30000; // 30 seconds
const MEMORY_WARNING_THRESHOLD = 0.85; // 85% of max memory
// RSS-based memory limit - use environment variable or default to 1.5GB for Azure B1 tier (1.75GB)
const RSS_MEMORY_LIMIT_MB = parseInt(process.env.NODE_MEMORY_LIMIT_MB || '1500', 10);
const RSS_WARNING_THRESHOLD_MB = RSS_MEMORY_LIMIT_MB * 0.85; // 85% of RSS limit
const RSS_EMERGENCY_THRESHOLD_MB = RSS_MEMORY_LIMIT_MB * 0.95; // 95% of RSS limit

// Error throttling from original
const errorThrottles = new Map();
const ERROR_THRESHOLD = 10;
const ERROR_WINDOW_MS = 1000;

// Emergency memory protection from original
let emergencyLoggingDisabled = false;
let lastMemoryCheck = Date.now();
const MEMORY_CHECK_INTERVAL_MS = 5000;

// Event subscriptions for cleanup
let subscriptions = [];

/**
 * Copy Winston logger initialization from original service
 */
function initLogger(logFilePath, logLevel = 'info') {
    if (!logFilePath && !process.env.MCP_LOG_PATH) {
        const dateSuffix = new Date().toISOString().slice(0,10).replace(/-/g, '');
        LOG_FILE_PATH = path.join(__dirname, `../../logs/mcp${dateSuffix}.log`);
    } else {
        LOG_FILE_PATH = logFilePath || process.env.MCP_LOG_PATH;
    }
    
    const logsDir = path.dirname(LOG_FILE_PATH);
    if (!fs.existsSync(logsDir)) {
        fs.mkdirSync(logsDir, { recursive: true });
    }
    
    const consoleFormat = winston.format.printf(({ level, message, timestamp, context, category, ...rest }) => {
        const prefix = category ? `[MCP ${category.toUpperCase()}]` : '[MCP]';
        
        // Handle context properly - it could be in different places
        let contextObj = context;
        if (!contextObj && rest.context) {
            contextObj = rest.context;
        }
        
        // Stringify context if it exists and is an object
        let contextStr = '';
        if (contextObj && typeof contextObj === 'object' && Object.keys(contextObj).length > 0) {
            try {
                contextStr = ` ${JSON.stringify(contextObj)}`;
            } catch (err) {
                contextStr = ` {context serialization failed}`;
            }
        }
        
        // Ensure message is a string
        const messageStr = typeof message === 'object' ? JSON.stringify(message) : String(message);
        
        return `${prefix} ${messageStr}${contextStr}`;
    });
    
    const fileFormat = winston.format.combine(
        winston.format.timestamp(),
        winston.format.json()
    );
    
    // Configure transports - exclude console transport in MCP silent mode
    const transports = [
        new winston.transports.File({ 
            filename: LOG_FILE_PATH, 
            maxsize: 2097152,
            maxFiles: 5,
            tailable: true,
            format: fileFormat,
            handleExceptions: true,
            handleRejections: true
        })
    ];
    
    // Only add console transport if not in MCP silent mode
    if (process.env.MCP_SILENT_MODE !== 'true') {
        transports.push(new winston.transports.Console({ 
            format: winston.format.combine(
                winston.format.colorize(),
                consoleFormat
            ),
            stderrLevels: ['error', 'warn'],
            consoleWarnLevels: [],
            handleExceptions: true,
            handleRejections: true
        }));
    }

    logger = winston.createLogger({
        level: logLevel,
        defaultMeta: {
            pid: process.pid,
            hostname: os.hostname(),
            version: appVersion
        },
        transports: transports,
        exitOnError: false
    });
}

/**
 * Copy error throttling logic from original service
 */
function shouldLogError(category) {
  const now = Date.now();
  const key = `error:${category || 'unknown'}`;
  
  if (!errorThrottles.has(key)) {
    errorThrottles.set(key, { count: 1, timestamp: now, suppressed: 0 });
    return true;
  }
  
  const record = errorThrottles.get(key);
  
  if (now - record.timestamp > ERROR_WINDOW_MS) {
    if (record.suppressed > 0) {
      console.error(`[MONITORING] Suppressed ${record.suppressed} similar errors in category '${category}' in the last ${ERROR_WINDOW_MS}ms`);
    }
    
    record.count = 1;
    record.timestamp = now;
    record.suppressed = 0;
    return true;
  }
  
  if (record.count >= ERROR_THRESHOLD) {
    record.suppressed++;
    return false;
  }
  
  record.count++;
  return true;
}

/**
 * Copy memory monitoring from original service
 * Updated to use RSS-based thresholds for more accurate monitoring
 */
function startMemoryMonitoring() {
  let memoryCheckInterval = null;

  const checkMemory = () => {
    try {
      const memoryUsage = process.memoryUsage();
      const rssMB = memoryUsage.rss / 1024 / 1024;
      const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
      const heapTotalMB = memoryUsage.heapTotal / 1024 / 1024;

      // Use RSS-based threshold for actual memory usage (not heap ratio which can be misleading)
      if (rssMB > RSS_WARNING_THRESHOLD_MB) {
        console.warn(`[MEMORY WARNING] High memory usage: RSS ${Math.round(rssMB)}MB / ${RSS_MEMORY_LIMIT_MB}MB limit (heap: ${Math.round(heapUsedMB)}MB / ${Math.round(heapTotalMB)}MB)`);

        // Get EventService with safe error handling
        const service = getEventService();
        if (service) {
          try {
            service.emit(eventTypes.SYSTEM_MEMORY_WARNING, {
              rssMB,
              rssLimitMB: RSS_MEMORY_LIMIT_MB,
              heapUsedMB,
              heapTotalMB,
              timestamp: new Date().toISOString()
            }).catch(err => {
              console.error('[MEMORY WARNING] Memory warning event promise rejected:', err.message);
            });
          } catch (emitError) {
            console.error('[MEMORY WARNING] Failed to emit memory warning event:', emitError.message);
          }
        }

        if (global.gc) {
          console.log('[MEMORY] Forcing garbage collection');
          global.gc();
        }
      }
    } catch (err) {
      // Silently ignore memory monitoring errors
    }
  };

  memoryCheckInterval = setInterval(checkMemory, MEMORY_CHECK_INTERVAL);

  process.on('exit', () => {
    if (memoryCheckInterval) {
      clearInterval(memoryCheckInterval);
    }
  });

  checkMemory();
}

/**
 * Copy emergency memory check from original service
 * Updated to use RSS-based thresholds for more accurate monitoring
 */
function checkMemoryForEmergency() {
  const now = Date.now();
  if (now - lastMemoryCheck < MEMORY_CHECK_INTERVAL_MS) {
    return emergencyLoggingDisabled;
  }

  lastMemoryCheck = now;

  try {
    const memoryUsage = process.memoryUsage();
    const rssMB = memoryUsage.rss / 1024 / 1024;
    const heapUsedMB = memoryUsage.heapUsed / 1024 / 1024;
    const heapTotalMB = memoryUsage.heapTotal / 1024 / 1024;

    // Use RSS-based threshold (actual memory usage) instead of heap ratio
    if (rssMB > RSS_EMERGENCY_THRESHOLD_MB) {
      if (!emergencyLoggingDisabled) {
        console.error(`[EMERGENCY] Disabling all logging due to critical memory usage: RSS ${Math.round(rssMB)}MB / ${RSS_MEMORY_LIMIT_MB}MB limit`);
        emergencyLoggingDisabled = true;

        // Emit emergency event if event service is available
        if (eventService) {
          try {
            eventService.emit(eventTypes.SYSTEM_EMERGENCY, {
              type: 'memory_critical',
              rssMB,
              rssLimitMB: RSS_MEMORY_LIMIT_MB,
              heapUsedMB,
              heapTotalMB,
              timestamp: new Date().toISOString()
            });
          } catch (error) {
            // Silently ignore event emission errors
          }
        }
      }
      return true;
    } else if (emergencyLoggingDisabled && rssMB < RSS_WARNING_THRESHOLD_MB) {
      console.log(`[EMERGENCY] Re-enabling logging as memory usage has decreased: RSS ${Math.round(rssMB)}MB`);
      emergencyLoggingDisabled = false;
    }
  } catch (e) {
    // If we can't check memory, assume it's safe to log
  }

  return emergencyLoggingDisabled;
}

/**
 * Determines if a log should be filtered out to reduce noise
 */
function shouldFilterLog(level, message, category, context = {}) {
    // Always allow error logs
    if (level === 'error') {
        return false;
    }
    
    // Filter static file requests
    if (category === 'api' && message && message.includes('GET') && 
        (message.includes('.js') || message.includes('.css') || message.includes('.ico') || 
         message.includes('.png') || message.includes('.jpg') || message.includes('.gif'))) {
        return true;
    }
    
    // Filter excessive event system metrics
    if (category === 'metrics' && message && (
        message.includes('event_subscribe_success') ||
        message.includes('event_listeners_count') ||
        message.includes('event_emit_success') ||
        message.includes('event_emit_no_listeners') ||
        message.includes('event_emit_error')
    )) {
        return true;
    }
    
    // Filter duplicate module registration messages
    if (message && message.includes('Module') && message.includes('registered')) {
        return true;
    }
    
    // Filter health check logs in production
    if (process.env.NODE_ENV === 'production' && 
        message && (message.includes('health') || message.includes('ping'))) {
        return true;
    }
    
    // Filter verbose debug logs in production
    if (process.env.NODE_ENV === 'production' && level === 'debug') {
        return true;
    }
    
    // Filter memory usage metrics unless they're warnings
    if (category === 'metrics' && message && message.includes('memory_usage') && level === 'info') {
        return true;
    }
    
    // Filter performance metrics for very fast operations (< 10ms)
    if (category === 'metrics' && context && context.metricValue && 
        typeof context.metricValue === 'number' && context.metricValue < 10) {
        return true;
    }
    
    return false;
}

/**
 * Handle log events from other components (event subscription)
 */
function handleLogEvent(logData) {
    // For metric events, handle differently to prevent recursion
    if (logData.type === 'metric' && logData.name) {
        // Add directly to buffer without calling trackMetric
        logBuffer.add(logData);
        
        // Log to Winston if available, but don't emit any new events
        if (logger) {
            try {
                logger.debug(`[METRIC] ${logData.name}: ${logData.value}`, { 
                    metricName: logData.name, 
                    metricValue: logData.value,
                    ...(logData.context || {})
                });
            } catch (err) {
                console.error(`[MONITORING] Failed to log metric to Winston: ${err.message}`);
            }
        }
        
        // Store user-specific metrics if userId is provided
        if (logData.userId) {
            try {
                const storage = getStorageService();
                if (storage) {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(logData.userId, 'info', `Metric: ${logData.name}`, 'metrics', {
                        metricName: logData.name,
                        metricValue: logData.value,
                        ...(logData.context || {})
                    }, null, logData.deviceId)
                        .catch(err => {
                            console.error(`[MCP METRIC] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP METRIC] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        return;
    }
    
    // Add to circular buffer for non-metric events
    logBuffer.add(logData);
    
    // Log to Winston if available
    if (logger) {
        try {
            logger.log(logData.level || 'info', logData.message, {
                context: logData.context,
                category: logData.category,
                timestamp: logData.timestamp,
                id: logData.id,
                pid: logData.pid,
                hostname: logData.hostname,
                version: logData.version,
                traceId: logData.traceId
            });
        } catch (err) {
            console.error(`[MONITORING] Failed to log to Winston: ${err.message}`);
        }
    }
}

/**
 * Get EventService instance with safe error handling for circular dependencies
 * @returns {Object} EventService instance or null
 */
function getEventService() {
  if (eventService) {
    return eventService;
  }
  
  try {
    // Create a proxy that will defer method calls until the real EventService is available
    const eventServiceProxy = new Proxy({}, {
      get: function(target, prop) {
        return function(...args) {
          // Try to load the real EventService if not already loaded
          if (!eventService) {
            try {
              eventService = require('./event-service.cjs');
            } catch (error) {
              console.warn(`[MONITORING] Failed to load EventService for ${prop}:`, error.message);
              return Promise.resolve(); // Return resolved promise for async methods
            }
          }
          
          // If the method exists on the real EventService, call it
          if (eventService && typeof eventService[prop] === 'function') {
            return eventService[prop](...args);
          }
          
          console.warn(`[MONITORING] EventService.${prop} is not a function`);
          return Promise.resolve(); // Return resolved promise for async methods
        };
      }
    });
    
    // Try to load the real EventService
    eventService = require('./event-service.cjs');
    return eventService;
  } catch (error) {
    console.warn('[MONITORING] EventService not available:', error.message);
    return null;
  }
}

/**
 * Initialize event service subscriptions
 */
async function initialize() {
  subscriptions = [];
  
  // Get EventService with safe error handling
  const service = getEventService();
  if (!service) {
    console.warn('[MONITORING] Event service not available for subscription');
    return;
  }
  
  // Subscribe to log events EXCEPT metrics to prevent recursion
  try {
    subscriptions.push(
      await service.subscribe(eventTypes.ERROR, handleLogEvent),
      await service.subscribe(eventTypes.INFO, handleLogEvent),
      await service.subscribe(eventTypes.WARN, handleLogEvent),
      await service.subscribe(eventTypes.DEBUG, handleLogEvent)
      // Removed subscription to METRIC events to prevent recursion
    );
    console.log('[MONITORING] Successfully subscribed to log events (excluding metrics)');
  } catch (error) {
    console.warn('[MONITORING] Failed to subscribe to event service:', error.message);
  }
}

/**
 * Create log data object compatible with original format
 */
function createLogData(level, message, context = {}, category = '', traceId = null, userId = null, deviceId = null) {
  const logData = {
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    level,
    category,
    message,
    context,
    pid: process.pid,
    hostname: os.hostname(),
    version: appVersion
  };
  
  if (traceId) {
    logData.traceId = traceId;
  }
  
  // Add multi-user context for isolation and monitoring
  if (userId) {
    logData.userId = userId;
  }
  
  if (deviceId) {
    logData.deviceId = deviceId;
  }
  
  return logData;
}

/**
 * Logs an error event - maintains same signature as original
 */
async function logError(error, userId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        console.debug('[MONITORING] Processing logError request', {
            errorType: error?.constructor?.name,
            hasCategory: !!(error && error.category),
            sessionId,
            timestamp: new Date().toISOString(),
            userId
        });
    }
    
    // Skip if emergency mode is active
    if (emergencyLoggingDisabled) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Skipping logError due to emergency mode', {
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }

    // Check memory before processing expensive operation
    checkMemoryForEmergency();
    
    try {
        // If not an MCP error, create one
        const mcpError = error && error.category ? error : ErrorService.createError(
            'unknown',
            error?.message || error?.toString() || 'Unknown error',
            'error',
            error
        );

        // Extract necessary fields
        const category = mcpError.category || 'unknown';
        const message = mcpError.message || 'Unknown error';
        
        // Check if we should throttle this error
        if (!shouldLogError(category)) {
            return;
        }
        
        // Format context
        const context = {
            ...mcpError.context,
            error: mcpError.stack || mcpError.toString(),
            category: mcpError.category,
            severity: mcpError.severity
        };
        
        // Generate traceId if not provided
        const traceId = mcpError.traceId || uuidv4();
        const errorUserId = mcpError.userId || userId;
        const deviceId = mcpError.deviceId;

        // Create log data with error details
        const logData = createLogData('error', message, context, category, traceId, errorUserId, deviceId);
        
        // Log to file
        logger.error(message, {
            category,
            context,
            traceId,
            userId: errorUserId,
            deviceId
        });
        
        // Emit error event
        if (eventService) {
            try {
                await eventService.emit(eventTypes.ERROR, logData);
            } catch (emitError) {
                console.error(`[MCP ERROR] Failed to emit error event: ${emitError.message}`);
            }
        }

        // Persist user-specific logs if userId is provided
        if (errorUserId) {
            try {
                const storage = getStorageService();
                if (storage) {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(errorUserId, 'error', message, category, context, traceId, deviceId)
                        .catch(err => {
                            console.error(`[MCP ERROR] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP ERROR] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        // Pattern 2: User Activity Logs (successful error logging)
        if (userId) {
            // Note: We don't call MonitoringService.info here to avoid recursion
            console.log(`[MONITORING] Error logged successfully for user ${userId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.log(`[MONITORING] Error logged successfully for session ${sessionId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        }
        
    } catch (e) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the error logging service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to log error in MonitoringService`, {
            originalError: error?.message || error?.toString(),
            logError: e.message,
            stack: e.stack,
            operation: 'logError',
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Error logging failed for user ${userId}`, {
                error: e.message,
                operation: 'logError',
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Error logging failed for session ${sessionId}`, {
                error: e.message,
                operation: 'logError',
                timestamp: new Date().toISOString()
            });
        }
        
        // Last resort fallback to console
        console.error(`[MCP ERROR] Failed to log error: ${e.message}`);
        console.error(`Original error: ${error?.message || error?.toString()}`);
    }
}

/**
 * Logs an error message - maintains same signature as original
 */
async function error(message, context = {}, category = '', traceId = null, userId = null, deviceId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        console.debug('[MONITORING] Processing error log request', {
            messageLength: message?.length,
            category,
            hasUserId: !!userId,
            hasDeviceId: !!deviceId,
            sessionId,
            timestamp: new Date().toISOString()
        });
    }
    
    if (checkMemoryForEmergency()) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Skipping error log due to emergency mode', {
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    if (!shouldLogError(category)) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Error log throttled', {
                category,
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    try {
        // Apply same filtering as original for calendar/graph errors
        if ((category === 'calendar' || category === 'graph') && 
            (message.includes('Graph API request failed') || 
             message.includes('Unable to read error response'))) {
            if (process.env.NODE_ENV === 'development') {
                console.warn(`[FILTERED] ${category} error: ${message}`);
            }
            return;
        }
        
        const logData = createLogData('error', message, context, category, traceId, userId, deviceId);
        
        // Add to circular buffer
        logBuffer.add(logData);
        
        // Don't emit event for our own logs - only handle events from other services
        
        if (logger) {
            logger.error(logData.message, {
                context: logData.context,
                category: logData.category,
                timestamp: logData.timestamp,
                id: logData.id,
                pid: logData.pid,
                hostname: logData.hostname,
                version: logData.version,
                traceId: logData.traceId,
                userId: logData.userId,
                deviceId: logData.deviceId
            });
            
            // Persist user-specific logs if userId is provided
            if (userId) {
                try {
                    const storage = getStorageService();
                    if (storage) {
                        // Don't await to avoid blocking - let it run in background
                        storage.addUserLog(userId, 'error', message, category, context, traceId, deviceId)
                            .catch(err => {
                                console.error(`[MCP ERROR] Failed to persist user log: ${err.message}`);
                            });
                    }
                } catch (storageError) {
                    console.error(`[MCP ERROR] Error accessing storage for user log: ${storageError.message}`);
                }
            }
        }
        
        // Pattern 2: User Activity Logs (successful error message logging)
        if (userId) {
            // Note: We don't call MonitoringService.info here to avoid recursion
            console.log(`[MONITORING] Error message logged successfully for user ${userId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.log(`[MONITORING] Error message logged successfully for session ${sessionId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        }
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the error logging service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to log error message in MonitoringService`, {
            originalMessage: message,
            logError: err.message,
            stack: err.stack,
            operation: 'error',
            category,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Error message logging failed for user ${userId}`, {
                error: err.message,
                operation: 'error',
                category,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Error message logging failed for session ${sessionId}`, {
                error: err.message,
                operation: 'error',
                category,
                timestamp: new Date().toISOString()
            });
        }
        
        console.error(`[MONITORING] Failed to log error: ${err.message}`);
    }
}

/**
 * Logs an info message - maintains same signature as original
 */
async function info(message, context = {}, category = 'general', traceId = null, userId = null, deviceId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        console.debug('[MONITORING] Processing info log request', {
            messageLength: message?.length,
            category,
            hasUserId: !!userId,
            hasDeviceId: !!deviceId,
            sessionId,
            timestamp: new Date().toISOString()
        });
    }
    
    if (checkMemoryForEmergency()) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Skipping info log due to emergency mode', {
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    if (!logger) initLogger();
    
    if (shouldFilterLog('info', message, category, context)) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Info log filtered', {
                category,
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    try {
        const logData = createLogData('info', message, context, category, traceId, userId, deviceId);
        
        // Add to circular buffer
        logBuffer.add(logData);
        
        // Don't emit event for our own logs - only handle events from other services
        
        logger.info(logData.message, {
            context: logData.context,
            category: logData.category,
            timestamp: logData.timestamp,
            id: logData.id,
            pid: logData.pid,
            hostname: logData.hostname,
            version: logData.version,
            traceId: logData.traceId,
            userId: logData.userId,
            deviceId: logData.deviceId
        });
        
        // Persist user-specific logs if userId is provided (but skip infrastructure logs)
        if (userId && !isInfrastructureLog(message, category)) {
            try {
                const storage = getStorageService();
                if (storage) {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(userId, 'info', message, category, context, traceId, deviceId)
                        .catch(err => {
                            console.error(`[MCP INFO] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP INFO] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        // Pattern 2: User Activity Logs (successful info message logging)
        if (userId) {
            // Note: We don't call MonitoringService.info here to avoid recursion
            console.log(`[MONITORING] Info message logged successfully for user ${userId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.log(`[MONITORING] Info message logged successfully for session ${sessionId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        }
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the monitoring service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to log info message in MonitoringService`, {
            originalMessage: message,
            logError: err.message,
            stack: err.stack,
            operation: 'info',
            category,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Info message logging failed for user ${userId}`, {
                error: err.message,
                operation: 'info',
                category,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Info message logging failed for session ${sessionId}`, {
                error: err.message,
                operation: 'info',
                category,
                timestamp: new Date().toISOString()
            });
        }
        
        console.error(`[MONITORING] Failed to log info message: ${err.message}`);
    }
}

/**
 * Logs a warning message - maintains same signature as original
 */
async function warn(message, context = {}, category = 'general', traceId = null, userId = null, deviceId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        console.debug('[MONITORING] Processing warn log request', {
            messageLength: message?.length,
            category,
            hasUserId: !!userId,
            hasDeviceId: !!deviceId,
            sessionId,
            timestamp: new Date().toISOString()
        });
    }
    
    if (checkMemoryForEmergency()) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Skipping warn log due to emergency mode', {
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    if (!logger) initLogger();
    
    if (shouldFilterLog('warn', message, category, context)) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Warn log filtered', {
                category,
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    try {
        const logData = createLogData('warn', message, context, category, traceId, userId, deviceId);
        
        // Add to circular buffer
        logBuffer.add(logData);
        
        logger.warn(logData.message, {
            context: logData.context,
            category: logData.category,
            timestamp: logData.timestamp,
            id: logData.id,
            pid: logData.pid,
            hostname: logData.hostname,
            version: logData.version,
            traceId: logData.traceId,
            userId: logData.userId,
            deviceId: logData.deviceId
        });
        
        // Persist user-specific logs if userId is provided (but skip infrastructure logs)
        if (userId && !isInfrastructureLog(message, category)) {
            try {
                const storage = getStorageService();
                if (storage) {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(userId, 'warn', message, category, context, traceId, deviceId)
                        .catch(err => {
                            console.error(`[MCP WARN] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP WARN] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        // Pattern 2: User Activity Logs (successful warn message logging)
        if (userId) {
            // Note: We don't call MonitoringService.info here to avoid recursion
            console.log(`[MONITORING] Warning message logged successfully for user ${userId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.log(`[MONITORING] Warning message logged successfully for session ${sessionId}`, {
                category,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        }
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the monitoring service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to log warning message in MonitoringService`, {
            originalMessage: message,
            logError: err.message,
            stack: err.stack,
            operation: 'warn',
            category,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Warning message logging failed for user ${userId}`, {
                error: err.message,
                operation: 'warn',
                category,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Warning message logging failed for session ${sessionId}`, {
                error: err.message,
                operation: 'warn',
                category,
                timestamp: new Date().toISOString()
            });
        }
        
        console.error(`[MONITORING] Failed to log warning message: ${err.message}`);
    }
}

/**
 * Logs a debug message - maintains same signature as original
 */
async function debug(message, context = {}, category = 'general', traceId = null, userId = null, deviceId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Skip if emergency logging is disabled
    if (emergencyLoggingDisabled) {
        // Note: We can't log debug messages about emergency mode since debug logging is disabled
        return;
    }
    
    // Skip debug logs in production silent mode
    const isProduction = process.env.NODE_ENV === 'production';
    const isSilentMode = process.env.MCP_SILENT_MODE === 'true';
    if (isProduction && isSilentMode) {
        return;
    }
    
    // Skip infrastructure logs in production
    if (isInfrastructureLog(message, category)) {
        return;
    }
    
    // Skip if this log should be filtered
    if (shouldFilterLog('debug', message, category, context)) {
        return;
    }
    
    try {
        const logData = createLogData('debug', message, context, category, traceId, userId, deviceId);
        
        // Add to circular buffer
        logBuffer.add(logData);
        
        // Don't emit event for our own logs - only handle events from other services
        
        logger.debug(logData.message, {
            context: logData.context,
            category: logData.category,
            timestamp: logData.timestamp,
            id: logData.id,
            pid: logData.pid,
            hostname: logData.hostname,
            version: logData.version,
            traceId: logData.traceId,
            userId: logData.userId,
            deviceId: logData.deviceId
        });
        
        // Persist user-specific logs if userId is provided (but skip infrastructure logs)
        if (userId && !isInfrastructureLog(message, category)) {
            try {
                const storage = getStorageService();
                if (storage) {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(userId, 'debug', message, category, context, traceId, deviceId)
                        .catch(err => {
                            console.error(`[MCP DEBUG] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP DEBUG] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        // Pattern 2: User Activity Logs (successful debug message logging)
        // Only in development mode to avoid noise
        if (process.env.NODE_ENV === 'development') {
            if (userId) {
                // Note: We don't call MonitoringService.info here to avoid recursion
                console.log(`[MONITORING] Debug message logged successfully for user ${userId}`, {
                    category,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                });
            } else if (sessionId) {
                console.log(`[MONITORING] Debug message logged successfully for session ${sessionId}`, {
                    category,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                });
            }
        }
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the monitoring service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to log debug message in MonitoringService`, {
            originalMessage: message,
            logError: err.message,
            stack: err.stack,
            operation: 'debug',
            category,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Debug message logging failed for user ${userId}`, {
                error: err.message,
                operation: 'debug',
                category,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Debug message logging failed for session ${sessionId}`, {
                error: err.message,
                operation: 'debug',
                category,
                timestamp: new Date().toISOString()
            });
        }
        
        console.error(`[MONITORING] Failed to log debug message: ${err.message}`);
    }
}

/**
 * Track a metric
 * @param {string} name - Metric name
 * @param {any} value - Metric value
 * @param {Object} context - Additional context
 * @param {string|null} userId - User ID if available
 * @param {string|null} deviceId - Device ID if available
 * @param {boolean} fromEvent - Internal flag to prevent recursion, true if this call originated from an event
 * @param {string|null} sessionId - Session ID if available
 */
async function trackMetric(name, value, context = {}, userId = null, deviceId = null, fromEvent = false, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        console.debug('[MONITORING] Processing trackMetric request', {
            metricName: name,
            metricValue: value,
            hasUserId: !!userId,
            hasDeviceId: !!deviceId,
            sessionId,
            timestamp: new Date().toISOString()
        });
    }
    
    // Skip metrics about storage operations to prevent recursive loops
    if (name.startsWith('storage_') || context.category === 'storage') {
        // Just log to console instead of persisting to prevent recursion
        if (logger) {
            logger.debug(`[METRIC] ${name}: ${value}`, { 
                metricName: name, 
                metricValue: value,
                ...context
            });
        }
        return;
    }

    // Emergency memory protection
    if (emergencyLoggingDisabled) {
        if (process.env.NODE_ENV === 'development') {
            console.debug('[MONITORING] Skipping trackMetric due to emergency mode', {
                sessionId,
                userId,
                timestamp: new Date().toISOString()
            });
        }
        return;
    }
    
    // Check memory periodically during metric tracking
    if (Date.now() - lastMemoryCheck > MEMORY_CHECK_INTERVAL_MS) {
        checkMemoryForEmergency();
    }
    
    try {
        // Filter out excessive metrics
        if (shouldFilterLog('info', `Metric: ${name}`, 'metrics', { name, value, ...context })) {
            if (process.env.NODE_ENV === 'development') {
                console.debug('[MONITORING] Metric filtered', {
                    metricName: name,
                    sessionId,
                    userId,
                    timestamp: new Date().toISOString()
                });
            }
            return;
        }
        
        if (!logger) initLogger();
        
        // Add to circular buffer
        const logData = {
            type: 'metric',
            name,
            value,
            context,
            timestamp: new Date().toISOString(),
            userId,
            deviceId
        };
        
        logBuffer.add(logData);
        
        // Log to Winston
        logger.debug(`[METRIC] ${name}: ${value}`, { 
            metricName: name, 
            metricValue: value,
            ...context
        });
        
        // Create metric log data for detailed logging
        const metricLogData = {
            level: 'info',
            message: `Metric: ${name}`,
            context: {
                metricName: name,
                metricValue: value,
                ...context
            },
            category: 'metrics',
            timestamp: new Date().toISOString(),
            userId,
            deviceId
        };
        
        // Add to circular buffer
        logBuffer.add(metricLogData);
        
        // Log to Winston
        logger.info(`Metric: ${name}`, {
            context: metricLogData.context,
            category: metricLogData.category,
            timestamp: metricLogData.timestamp,
            userId: metricLogData.userId,
            deviceId: metricLogData.deviceId
        });
        
        // Store user-specific metrics in database if storage service is available
        // (consolidated duplicate code)
        if (userId && name !== 'storage_add_user_log_success') {
            try {
                const storage = getStorageService();
                if (storage && typeof storage.addUserLog === 'function') {
                    // Don't await to avoid blocking - let it run in background
                    storage.addUserLog(userId, 'info', `Metric: ${name}`, 'metrics', {
                        metricName: name,
                        metricValue: value,
                        ...context
                    }, null, deviceId)
                        .catch(err => {
                            console.error(`[MCP METRIC] Failed to persist user log: ${err.message}`);
                        });
                }
            } catch (storageError) {
                console.error(`[MCP METRIC] Error accessing storage for user log: ${storageError.message}`);
            }
        }
        
        // Pattern 2: User Activity Logs (successful metric tracking)
        if (userId) {
            // Note: We don't call MonitoringService.info here to avoid recursion
            console.log(`[MONITORING] Metric tracked successfully for user ${userId}`, {
                metricName: name,
                metricValue: value,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.log(`[MONITORING] Metric tracked successfully for session ${sessionId}`, {
                metricName: name,
                metricValue: value,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            });
        }
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        // Use console.error to avoid recursion since this IS the monitoring service
        console.error(`[MCP INFRASTRUCTURE ERROR] Failed to track metric in MonitoringService`, {
            metricName: name,
            metricValue: value,
            logError: err.message,
            stack: err.stack,
            operation: 'trackMetric',
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            console.error(`[MCP USER ERROR] Metric tracking failed for user ${userId}`, {
                error: err.message,
                operation: 'trackMetric',
                metricName: name,
                timestamp: new Date().toISOString()
            });
        } else if (sessionId) {
            console.error(`[MCP SESSION ERROR] Metric tracking failed for session ${sessionId}`, {
                error: err.message,
                operation: 'trackMetric',
                metricName: name,
                timestamp: new Date().toISOString()
            });
        }
        
        console.error(`[MONITORING] Failed to track metric: ${err.message}`);
    }
    
    // Don't emit events for metrics to prevent recursion
}

/**
 * Subscribe to log events - maintains same signature as original
 */
function subscribeToLogs(callback) {
    // For backward compatibility, subscribe to all log events
    const unsubscribeFunctions = [];
    
    // Get EventService with safe error handling
    const service = getEventService();
    if (!service) {
        console.warn('[MONITORING] Event service not available for log subscription');
        return () => {}; // Return no-op unsubscribe function
    }
    
    // Helper to subscribe to a specific event type
    const subscribeToEvent = async (eventType) => {
        try {
            const unsubscribeId = await service.subscribe(eventType, callback);
            return () => {
                if (service && typeof service.unsubscribe === 'function') {
                    return service.unsubscribe(unsubscribeId).catch(err => {
                        console.warn(`[MONITORING] Error unsubscribing from ${eventType}:`, err.message);
                    });
                }
            };
        } catch (error) {
            console.warn(`[MONITORING] Failed to subscribe to ${eventType}:`, error.message);
            return () => {}; // Return no-op unsubscribe function
        }
    };
    
    // Subscribe to all log event types
    Promise.all([
        subscribeToEvent(eventTypes.ERROR),
        subscribeToEvent(eventTypes.INFO),
        subscribeToEvent(eventTypes.WARN),
        subscribeToEvent(eventTypes.DEBUG)
    ]).then(unsubscribes => {
        unsubscribeFunctions.push(...unsubscribes);
    }).catch(error => {
        console.warn('[MONITORING] Failed to subscribe to log events:', error.message);
    });
    
    // Return unsubscribe function that cleans up all subscriptions
    return () => {
        unsubscribeFunctions.forEach(unsub => {
            if (typeof unsub === 'function') {
                try {
                    unsub();
                } catch (error) {
                    console.warn('[MONITORING] Error during unsubscribe:', error.message);
                }
            }
        });
    };
}

/**
 * Subscribe to metric events - maintains same signature as original
 */
function subscribeToMetrics(callback) {
    let unsubscribeFunction = null;
    
    // Get EventService with safe error handling
    const service = getEventService();
    if (!service) {
        console.warn('[MONITORING] Event service not available for metrics subscription');
        return () => {}; // Return no-op unsubscribe function
    }
    
    service.subscribe(eventTypes.METRIC, callback).then(id => {
        unsubscribeFunction = () => {
            if (service && typeof service.unsubscribe === 'function') {
                return service.unsubscribe(id).catch(err => {
                    console.warn('[MONITORING] Error unsubscribing from metrics:', err.message);
                });
            }
        };
    }).catch(error => {
        console.warn('[MONITORING] Failed to subscribe to metric events:', error.message);
    });
    
    return () => {
        if (typeof unsubscribeFunction === 'function') {
            try {
                unsubscribeFunction();
            } catch (error) {
                console.warn('[MONITORING] Error during metrics unsubscribe:', error.message);
            }
        }
    };
}

/**
 * Get latest logs from circular buffer instead of files
 */
async function getLatestLogs(limit = 100) {
    const logs = logBuffer.getAll();
    
    // Sort by timestamp (newest first) and limit
    return logs
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0, limit);
}

/**
 * Get the circular buffer for direct access (new method)
 */
function getLogBuffer() {
    return logBuffer;
}

/**
 * For test: allow resetting logger with new path - maintains same signature as original
 */
function _resetLoggerForTest(logFilePath, logLevel = 'info') {
    if (logger) {
        for (const t of logger.transports) logger.remove(t);
    }
    initLogger(logFilePath, logLevel);
}

// Initialize logger and memory monitoring at startup
initLogger();
startMemoryMonitoring();

// Defer event service initialization to avoid circular dependency
setTimeout(() => {
  initialize().catch(error => {
    console.warn('[MONITORING] Event service initialization failed:', error.message);
  });
}, 100); // Small delay to ensure all modules are loaded

module.exports = {
    logError,
    error,
    info,
    warn,
    debug,
    trackMetric,
    LOG_FILE_PATH,
    _resetLoggerForTest,
    initLogger,
    subscribeToLogs,
    subscribeToMetrics,
    getLatestLogs,
    getLogBuffer // New method for direct buffer access
};