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

// Import event service for event-based architecture
const eventService = require('./event-service.cjs');

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
    
    const consoleFormat = winston.format.printf(({ level, message, timestamp, context, category }) => {
        const prefix = category ? `[MCP ${category.toUpperCase()}]` : '[MCP]';
        return `${prefix} ${message}`;
    });
    
    const fileFormat = winston.format.combine(
        winston.format.timestamp(),
        winston.format.json()
    );
    
    logger = winston.createLogger({
        level: logLevel,
        defaultMeta: {
            pid: process.pid,
            hostname: os.hostname(),
            version: appVersion
        },
        transports: [
            new winston.transports.File({ 
                filename: LOG_FILE_PATH, 
                maxsize: 2097152,
                maxFiles: 5,
                tailable: true,
                format: fileFormat,
                handleExceptions: true,
                handleRejections: true
            }),
            new winston.transports.Console({ 
                format: winston.format.combine(
                    winston.format.colorize(),
                    consoleFormat
                ),
                stderrLevels: ['error', 'warn'],
                consoleWarnLevels: [],
                handleExceptions: true,
                handleRejections: true
            })
        ],
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
 */
function startMemoryMonitoring() {
  let memoryCheckInterval = null;
  
  const checkMemory = () => {
    try {
      const memoryUsage = process.memoryUsage();
      const heapUsed = memoryUsage.heapUsed;
      const heapTotal = memoryUsage.heapTotal;
      const usageRatio = heapUsed / heapTotal;
      
      if (usageRatio > MEMORY_WARNING_THRESHOLD) {
        console.warn(`[MEMORY WARNING] High memory usage: ${Math.round(usageRatio * 100)}% (${Math.round(heapUsed / 1024 / 1024)}MB / ${Math.round(heapTotal / 1024 / 1024)}MB)`);
        
        // Emit memory warning event
        eventService.emit(eventTypes.SYSTEM_MEMORY_WARNING, {
          usageRatio,
          heapUsed,
          heapTotal,
          timestamp: new Date().toISOString()
        });
        
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
 */
function checkMemoryForEmergency() {
  const now = Date.now();
  if (now - lastMemoryCheck < MEMORY_CHECK_INTERVAL_MS) {
    return false;
  }
  
  lastMemoryCheck = now;
  
  try {
    const memoryUsage = process.memoryUsage();
    const heapUsed = memoryUsage.heapUsed;
    const heapTotal = memoryUsage.heapTotal;
    const usageRatio = heapUsed / heapTotal;
    
    if (usageRatio > 0.95) {
      if (!emergencyLoggingDisabled) {
        console.error(`[EMERGENCY] Disabling all logging due to critical memory usage: ${Math.round(usageRatio * 100)}%`);
        emergencyLoggingDisabled = true;
        
        // Emit emergency event
        eventService.emit(eventTypes.SYSTEM_EMERGENCY, {
          type: 'memory_critical',
          usageRatio,
          timestamp: new Date().toISOString()
        });
      }
      return true;
    } else if (emergencyLoggingDisabled && usageRatio < 0.80) {
      console.log(`[EMERGENCY] Re-enabling logging as memory usage has decreased: ${Math.round(usageRatio * 100)}%`);
      emergencyLoggingDisabled = false;
    }
  } catch (e) {
    // If we can't check memory, assume it's safe to log
  }
  
  return emergencyLoggingDisabled;
}

/**
 * Handle log events from other components (event subscription)
 */
function handleLogEvent(logData) {
  // Add to circular buffer
  logBuffer.add(logData);
  
  // Log to Winston if available
  if (logger) {
    try {
      logger.log(logData.level || 'info', logData);
    } catch (err) {
      console.error(`[MONITORING] Failed to log to Winston: ${err.message}`);
    }
  }
}

/**
 * Initialize event service subscriptions
 */
async function initialize() {
  subscriptions = [];
  
  // Subscribe to all log events
  subscriptions.push(
    await eventService.subscribe(eventTypes.ERROR, handleLogEvent),
    await eventService.subscribe(eventTypes.INFO, handleLogEvent),
    await eventService.subscribe(eventTypes.WARN, handleLogEvent),
    await eventService.subscribe(eventTypes.DEBUG, handleLogEvent),
    await eventService.subscribe(eventTypes.METRIC, handleLogEvent)
  );
}

/**
 * Create log data object compatible with original format
 */
function createLogData(level, message, context = {}, category = '', traceId = null) {
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
  
  return logData;
}

/**
 * Logs an error event - maintains same signature as original
 */
function logError(error) {
    if (!logger) initLogger();
    
    if (!shouldLogError(error.category)) {
        return;
    }
    
    const logData = {
        id: error.id,
        category: error.category,
        message: error.message,
        severity: error.severity,
        context: error.context || {},
        timestamp: error.timestamp || new Date().toISOString(),
        level: 'error'
    };
    
    if (error.traceId) {
        logData.traceId = error.traceId;
    }
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    try {
        logger.error(logData);
    } catch (err) {
        console.error(`[MONITORING] Failed to log error: ${err.message}`);
    }
}

/**
 * Logs an error message - maintains same signature as original
 */
function error(message, context = {}, category = '', traceId = null) {
    if (checkMemoryForEmergency()) {
        return;
    }
    
    if (!shouldLogError(category)) {
        return;
    }
    
    // Apply same filtering as original for calendar/graph errors
    if ((category === 'calendar' || category === 'graph') && 
        (message.includes('Graph API request failed') || 
         message.includes('Unable to read error response'))) {
        if (process.env.NODE_ENV === 'development') {
            console.warn(`[FILTERED] ${category} error: ${message}`);
        }
        return;
    }
    
    const logData = createLogData('error', message, context, category, traceId);
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    if (logger) {
        try {
            logger.error(logData);
        } catch (err) {
            console.error(`[MONITORING] Failed to log error: ${err.message}`);
        }
    }
}

/**
 * Logs an info message - maintains same signature as original
 */
function info(message, context = {}, category = '', traceId = null) {
    if (checkMemoryForEmergency()) {
        return;
    }
    
    if (!logger) initLogger();
    
    const logData = createLogData('info', message, context, category, traceId);
    
    // Apply same category filtering as original
    if (category === 'api' || category === 'calendar' || category === 'graph') {
        if (process.env.NODE_ENV !== 'development') {
            return;
        }
    }
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    try {
        logger.info(logData);
    } catch (err) {
        console.error(`[MONITORING] Failed to log info message: ${err.message}`);
    }
}

/**
 * Logs a warning message - maintains same signature as original
 */
function warn(message, context = {}, category = '', traceId = null) {
    if (!logger) initLogger();
    
    const logData = createLogData('warn', message, context, category, traceId);
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    try {
        logger.warn(logData);
    } catch (err) {
        console.error(`[MONITORING] Failed to log warning message: ${err.message}`);
    }
}

/**
 * Logs a debug message - maintains same signature as original
 */
function debug(message, context = {}, category = '', traceId = null) {
    if (!logger) initLogger();
    
    const logData = createLogData('debug', message, context, category, traceId);
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    try {
        logger.debug(logData);
    } catch (err) {
        console.error(`[MONITORING] Failed to log debug message: ${err.message}`);
    }
}

/**
 * Tracks a performance metric - maintains same signature as original
 */
function trackMetric(name, value, context = {}) {
    if (!logger) initLogger();
    
    const logData = {
        type: 'metric',
        metric: name,
        value,
        context,
        timestamp: new Date().toISOString(),
        pid: process.pid,
        hostname: os.hostname(),
        version: appVersion
    };
    
    // Add to circular buffer
    logBuffer.add(logData);
    
    // Don't emit event for our own logs - only handle events from other services
    
    logger.info(logData);
}

/**
 * Subscribe to log events - maintains same signature as original
 */
function subscribeToLogs(callback) {
    // For backward compatibility, subscribe to all log events
    const unsubscribeFunctions = [];
    
    const subscribeToEvent = async (eventType) => {
        const id = await eventService.subscribe(eventType, callback);
        return () => eventService.unsubscribe(id);
    };
    
    Promise.all([
        subscribeToEvent(eventTypes.ERROR),
        subscribeToEvent(eventTypes.INFO),
        subscribeToEvent(eventTypes.WARN),
        subscribeToEvent(eventTypes.DEBUG)
    ]).then(unsubscribes => {
        unsubscribeFunctions.push(...unsubscribes);
    });
    
    // Return unsubscribe function that cleans up all subscriptions
    return () => {
        unsubscribeFunctions.forEach(unsub => unsub());
    };
}

/**
 * Subscribe to metric events - maintains same signature as original
 */
function subscribeToMetrics(callback) {
    let unsubscribeFunction = null;
    
    eventService.subscribe(eventTypes.METRIC, callback).then(id => {
        unsubscribeFunction = () => eventService.unsubscribe(id);
    });
    
    return () => {
        if (unsubscribeFunction) unsubscribeFunction();
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

// Initialize logger and event subscriptions at startup
initLogger();
initialize();
startMemoryMonitoring();

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