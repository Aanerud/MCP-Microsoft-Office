/**
 * @fileoverview StorageService handles persistent storage using multiple database types for MCP Server.
 * Provides async CRUD for settings/history and encryption for sensitive data. Production-ready with connection pooling.
 */

const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const ErrorService = require('./error-service.cjs');
const MonitoringService = require('./monitoring-service.cjs');
const { databaseFactory } = require('./database-factory.cjs');
const { MigrationManager } = require('./database-migrations.cjs');
const { BackupManager } = require('./database-backup.cjs');
const { getConfig } = require('../config/index.cjs');

// Legacy SQLite path for backward compatibility
const DB_PATH = path.join(__dirname, '../../data/mcp.sqlite');

// SEC-1: Encryption key must be set via environment variable in production
const ENCRYPTION_KEY = process.env.DEVICE_REGISTRY_ENCRYPTION_KEY || process.env.MCP_ENCRYPTION_KEY;
if (!ENCRYPTION_KEY) {
    if (process.env.NODE_ENV === 'production') {
        throw new Error('CRITICAL: DEVICE_REGISTRY_ENCRYPTION_KEY or MCP_ENCRYPTION_KEY environment variable must be set in production');
    }
    // Development-only fallback with warning
    console.warn('[SECURITY WARNING] Using default encryption key - set DEVICE_REGISTRY_ENCRYPTION_KEY for production');
}
const EFFECTIVE_ENCRYPTION_KEY = ENCRYPTION_KEY || 'dev_default_key_32bytes_long__!!';

// Service instances
let migrationManager = null;
let backupManager = null;
let initialized = false;

// Log service initialization
MonitoringService.info('Storage service initialized', {
    serviceName: 'storage-service',
    timestamp: new Date().toISOString()
}, 'storage');

if (Buffer.from(EFFECTIVE_ENCRYPTION_KEY).length !== 32) {
    const mcpError = ErrorService.createError(
        ErrorService.CATEGORIES.SYSTEM,
        'DEVICE_REGISTRY_ENCRYPTION_KEY must be exactly 32 bytes for AES-256-CBC',
        ErrorService.SEVERITIES.CRITICAL,
        {
            keyLength: Buffer.from(EFFECTIVE_ENCRYPTION_KEY).length,
            requiredLength: 32,
            timestamp: new Date().toISOString()
        }
    );
    MonitoringService.logError(mcpError);
    throw mcpError;
}

/**
 * Get database connection from factory
 */
async function getConnection() {
    if (!initialized) {
        throw new Error('Storage service not initialized. Call init() first.');
    }
    return await databaseFactory.getConnection();
}

function encrypt(text) {
    const startTime = Date.now();
    
    try {
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Encryption operation started', {
                textLength: text.length,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const iv = crypto.randomBytes(16);
        const cipher = crypto.createCipheriv('aes-256-cbc', Buffer.from(EFFECTIVE_ENCRYPTION_KEY), iv);
        let encrypted = cipher.update(text, 'utf8', 'base64');
        encrypted += cipher.final('base64');
        const result = iv.toString('base64') + ':' + encrypted;
        
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('storage_encryption_success', executionTime, {
            textLength: text.length,
            timestamp: new Date().toISOString()
        });
        
        return result;
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        const mcpError = ErrorService.createError(
            ErrorService.CATEGORIES.SYSTEM,
            `Encryption failed: ${error.message}`,
            ErrorService.SEVERITIES.ERROR,
            {
                textLength: text.length,
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
        
        MonitoringService.logError(mcpError);
        MonitoringService.trackMetric('storage_encryption_failure', executionTime, {
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

function decrypt(data) {
    const startTime = Date.now();
    
    try {
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Decryption operation started', {
                dataLength: data.length,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const [ivStr, encrypted] = data.split(':');
        if (!ivStr || !encrypted) {
            throw new Error('Invalid encrypted data format');
        }
        
        const iv = Buffer.from(ivStr, 'base64');
        const decipher = crypto.createDecipheriv('aes-256-cbc', Buffer.from(EFFECTIVE_ENCRYPTION_KEY), iv);
        let decrypted = decipher.update(encrypted, 'base64', 'utf8');
        decrypted += decipher.final('utf8');
        
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('storage_decryption_success', executionTime, {
            dataLength: data.length,
            timestamp: new Date().toISOString()
        });
        
        return decrypted;
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        const mcpError = ErrorService.createError(
            ErrorService.CATEGORIES.SYSTEM,
            `Decryption failed: ${error.message}`,
            ErrorService.SEVERITIES.ERROR,
            {
                dataLength: data.length,
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
        
        MonitoringService.logError(mcpError);
        MonitoringService.trackMetric('storage_decryption_failure', executionTime, {
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function init(userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Storage service initialization started', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        // Get config asynchronously
        const config = await getConfig();
        
        // Initialize database factory first
        await databaseFactory.init(config, userId, sessionId);
        
        migrationManager = new MigrationManager(databaseFactory, userId, sessionId);
        backupManager = new BackupManager(config, userId, sessionId);
        
        await migrationManager.migrate(userId, sessionId);
        
        initialized = true;
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Storage service initialized successfully', {
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Storage service initialized with session', {
                sessionId,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_init_success', executionTime, {
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Storage initialization failed: ${error.message}`,
            'error',
            {
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Storage initialization failed', {
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Storage initialization failed', {
                sessionId,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_init_failure', executionTime, {
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function setSetting(key, value, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Setting storage operation started', {
                key,
                valueType: typeof value,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        await connection.query('INSERT OR REPLACE INTO settings (key, value, user_id) VALUES (?, ?, ?)', [key, JSON.stringify(value), userId]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Setting stored successfully', {
                key,
                valueType: typeof value,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Setting stored with session', {
                sessionId,
                key,
                valueType: typeof value,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_set_setting_success', executionTime, {
            key,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to set setting: ${error.message}`,
            'error',
            {
                key,
                valueType: typeof value,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Setting storage failed', {
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Setting storage failed', {
                sessionId,
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_set_setting_failure', executionTime, {
            key,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function getSetting(key, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting setting operation started', {
                key,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        const row = await connection.query('SELECT value FROM settings WHERE key = ? AND user_id = ?', [key, userId]);
        connection.release();
        
        const result = row && row.length > 0 ? JSON.parse(row[0].value) : null;
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Setting retrieved successfully', {
                key,
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Setting retrieved with session', {
                sessionId,
                key,
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_get_setting_success', executionTime, {
            key,
            found: !!result,
            timestamp: new Date().toISOString()
        });
        
        return result;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to get setting: ${error.message}`,
            'error',
            {
                key,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Setting retrieval failed', {
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Setting retrieval failed', {
                sessionId,
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_get_setting_failure', executionTime, {
            key,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function setSecure(key, value, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Secure setting storage operation started', {
                key,
                valueLength: value.length,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const enc = encrypt(value);
        const connection = await getConnection();
        await connection.query('INSERT OR REPLACE INTO secure_settings (key, encrypted_value, user_id) VALUES (?, ?, ?)', [key, enc, userId]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Secure setting stored successfully', {
                key,
                valueLength: value.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Secure setting stored with session', {
                sessionId,
                key,
                valueLength: value.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_set_secure_success', executionTime, {
            key,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to set secure setting: ${error.message}`,
            'error',
            {
                key,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Secure setting storage failed', {
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Secure setting storage failed', {
                sessionId,
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_set_secure_failure', executionTime, {
            key,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function getSecure(key, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting secure setting operation started', {
                key,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        const row = await connection.query('SELECT encrypted_value FROM secure_settings WHERE key = ? AND user_id = ?', [key, userId]);
        connection.release();
        
        const result = row && row.length > 0 ? decrypt(row[0].encrypted_value) : null;
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Secure setting retrieved successfully', {
                key,
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Secure setting retrieved with session', {
                sessionId,
                key,
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_get_secure_success', executionTime, {
            key,
            found: !!result,
            timestamp: new Date().toISOString()
        });
        
        return result;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to get secure setting: ${error.message}`,
            'error',
            {
                key,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Secure setting retrieval failed', {
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Secure setting retrieval failed', {
                sessionId,
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        MonitoringService.trackMetric('storage_get_secure_failure', executionTime, {
            key,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

/**
 * Delete an encrypted setting from the database
 * @param {string} key - The setting key
 * @param {string} userId - User identifier
 * @param {string} [sessionId] - Session identifier for logging
 */
async function deleteSecure(key, userId, sessionId) {
    const startTime = Date.now();

    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Delete secure setting operation started', {
                key,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }

        const connection = await getConnection();
        await connection.query('DELETE FROM secure_settings WHERE key = ? AND user_id = ?', [key, userId]);
        connection.release();

        const executionTime = Date.now() - startTime;

        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Secure setting deleted successfully', {
                key,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Secure setting deleted with session', {
                sessionId,
                key,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }

        MonitoringService.trackMetric('storage_delete_secure_success', executionTime, {
            key,
            timestamp: new Date().toISOString()
        });

    } catch (error) {
        const executionTime = Date.now() - startTime;

        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to delete secure setting: ${error.message}`,
            'error',
            {
                key,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);

        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Secure setting deletion failed', {
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Secure setting deletion failed', {
                sessionId,
                key,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        MonitoringService.trackMetric('storage_delete_secure_failure', executionTime, {
            key,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });

        throw mcpError;
    }
}

/**
 * Store a user-specific log
 * @param {string} userId - User ID
 * @param {string} level - Log level (error, warn, info, debug)
 * @param {string} message - Log message
 * @param {string} [category] - Log category
 * @param {Object} [context] - Additional context for the log
 * @param {string} [traceId] - Trace ID for request correlation
 * @param {string} [deviceId] - Device ID that generated the log
 * @returns {Promise<Object>} The stored log entry
 */
async function addUserLog(userId, level, message, category = null, context = null, traceId = null, deviceId = null, sessionId = null) {
  const startTime = Date.now();
  
  try {
    // Skip all logging in addUserLog to prevent recursive loops
    
    if (!userId) {
      throw new Error('userId is required for user log storage');
    }

    const contextStr = context ? JSON.stringify(context) : null;
    const db = await getConnection();
    
    // Use parameterized query to prevent SQL injection
    const query = `INSERT INTO user_logs (user_id, level, message, category, context, trace_id, device_id) 
                  VALUES (?, ?, ?, ?, ?, ?, ?)`;

    const params = [userId, level, message, category, contextStr, traceId, deviceId];
    
    // Execute query
    const result = await db.query(query, params);
    db.release();
    
    // Get the inserted row ID (handled differently per database)
    let logId;
    if (db.type === 'postgresql') {
      // PostgreSQL returns the inserted row
      logId = result.rows[0].id;
    } else if (db.type === 'sqlite') {
      // SQLite - use lastID
      logId = result.lastID;
    } else {
      // MySQL - use insertId
      logId = result.insertId;
    }

    const executionTime = Date.now() - startTime;
    
    // Skip all logging in addUserLog to prevent recursive loops
    // The MonitoringService likely calls back to addUserLog, creating infinite recursion

    // Return the created log entry
    return {
      id: logId,
      userId,
      level,
      message,
      category,
      context,
      timestamp: new Date(),
      traceId,
      deviceId
    };
  } catch (error) {
    // Skip all logging in addUserLog to prevent recursive loops
    // Just create a simple error and throw it
    const mcpError = ErrorService.createError(
      'storage',
      `Failed to add user log: ${error.message}`,
      'error',
      { 
        userId, 
        level, 
        message, 
        category: category || 'unknown',
        sessionId,
        deviceId,
        error: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    throw mcpError;
  }
}

/**
 * Retrieve logs for a specific user
 * @param {string} userId - User ID
 * @param {Object} options - Query options
 * @param {number} [options.limit=100] - Maximum number of logs to retrieve
 * @param {number} [options.offset=0] - Offset for pagination
 * @param {string} [options.level] - Filter by log level
 * @param {string} [options.category] - Filter by category
 * @param {string} [options.search] - Search term in message
 * @param {Date} [options.startDate] - Start date for filtering
 * @param {Date} [options.endDate] - End date for filtering
 * @returns {Promise<Array>} Array of log entries
 */
async function getUserLogs(userId, options = {}, sessionId = null) {
  const startTime = Date.now();

  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('User logs retrieval operation started', {
        userId,
        options: JSON.stringify(options),
        sessionId,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    if (!userId) {
      throw new Error('userId is required for retrieving user logs');
    }

    const { 
      limit = 100, 
      offset = 0, 
      level = null, 
      category = null, 
      search = null,
      startDate = null,
      endDate = null
    } = options;
    
    const db = await getConnection();
    
    // Build the WHERE clause and params dynamically based on filters
    let whereConditions = ['user_id = ?'];
    let params = [userId];
    
    if (level) {
      whereConditions.push('level = ?');
      params.push(level);
    }
    
    if (category) {
      whereConditions.push('category = ?');
      params.push(category);
    }
    
    if (search) {
      whereConditions.push('message LIKE ?');
      params.push(`%${search}%`);
    }
    
    if (startDate) {
      whereConditions.push('timestamp >= ?');
      params.push(startDate.toISOString());
    }
    
    if (endDate) {
      whereConditions.push('timestamp <= ?');
      params.push(endDate.toISOString());
    }
    
    // Create the final WHERE clause
    const whereClause = whereConditions.join(' AND ');
    
    // Create the query with pagination - format timestamp as proper ISO string
    const query = `SELECT id, user_id, message, category, level, 
                   REPLACE(timestamp, ' ', 'T') || 'Z' as timestamp
                   FROM user_logs WHERE ${whereClause} 
                  ORDER BY timestamp DESC LIMIT ? OFFSET ?`;
    params.push(limit, offset);
    
    const rows = await db.query(query, params);
    db.release();
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (userId) {
      MonitoringService.info('User logs retrieved successfully', {
        resultCount: rows.length,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.info('User logs retrieved with session', {
        sessionId,
        resultCount: rows.length,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_get_user_logs_success', executionTime, {
      userId: userId,
      resultCount: rows.length,
      filters: JSON.stringify(options),
      timestamp: new Date().toISOString()
    }, 'storage');
    
    // Process the results
    return rows.map(row => {
      // Parse context if it exists and is a string
      let parsedContext = row.context;
      if (row.context && typeof row.context === 'string') {
        try {
          parsedContext = JSON.parse(row.context);
        } catch (e) {
          // If parsing fails, keep as string
        }
      }
      
      return {
        id: row.id,
        userId: row.user_id,
        level: row.level,
        message: row.message,
        category: row.category,
        context: parsedContext,
        timestamp: new Date(row.timestamp),
        traceId: row.trace_id,
        deviceId: row.device_id
      };
    });
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'storage',
      `Failed to get user logs: ${error.message}`,
      'error',
      { 
        userId, 
        options, 
        sessionId,
        error: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (userId) {
      MonitoringService.error('User logs retrieval failed', {
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.error('User logs retrieval failed', {
        sessionId,
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_get_user_logs_failure', executionTime, {
      userId: userId,
      filters: JSON.stringify(options),
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'storage');
    
    throw mcpError;
  }
}

/**
 * Count total user logs matching filter criteria
 * @param {string} userId - User ID
 * @param {Object} options - Filter options
 * @returns {Promise<number>} Total count of logs
 */
async function countUserLogs(userId, options = {}, sessionId = null) {
  const startTime = Date.now();

  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('User logs count operation started', {
        userId,
        options: JSON.stringify(options),
        sessionId,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    if (!userId) {
      throw new Error('userId is required for counting user logs');
    }

    const { 
      level = null, 
      category = null, 
      search = null,
      startDate = null,
      endDate = null
    } = options;
    
    const db = await getConnection();
    
    // Build the WHERE clause and params dynamically
    let whereConditions = ['user_id = ?'];
    let params = [userId];
    
    if (level) {
      whereConditions.push('level = ?');
      params.push(level);
    }
    
    if (category) {
      whereConditions.push('category = ?');
      params.push(category);
    }
    
    if (search) {
      whereConditions.push('message LIKE ?');
      params.push(`%${search}%`);
    }
    
    if (startDate) {
      whereConditions.push('timestamp >= ?');
      params.push(startDate.toISOString());
    }
    
    if (endDate) {
      whereConditions.push('timestamp <= ?');
      params.push(endDate.toISOString());
    }
    
    // Create the final WHERE clause
    const whereClause = whereConditions.join(' AND ');
    
    // Build the count query
    const query = `SELECT COUNT(*) as count FROM user_logs WHERE ${whereClause}`;
    
    const result = await db.query(query, params);
    db.release();
    
    const executionTime = Date.now() - startTime;
    const count = parseInt(result[0].count, 10);
    
    // Pattern 2: User Activity Logs
    if (userId) {
      MonitoringService.info('User logs count retrieved successfully', {
        count,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.info('User logs count retrieved with session', {
        sessionId,
        count,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_count_user_logs_success', executionTime, {
      userId: userId,
      filters: JSON.stringify(options),
      timestamp: new Date().toISOString()
    }, 'storage');
    
    return count;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'storage',
      `Failed to count user logs: ${error.message}`,
      'error',
      { 
        userId, 
        options, 
        sessionId,
        error: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (userId) {
      MonitoringService.error('User logs count failed', {
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.error('User logs count failed', {
        sessionId,
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_count_user_logs_failure', executionTime, {
      userId: userId,
      filters: JSON.stringify(options),
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'storage');
    
    throw mcpError;
  }
}

/**
 * Clear user logs by criteria
 * @param {string} userId - User ID
 * @param {Object} options - Criteria for deletion
 * @param {Date} [options.olderThan] - Delete logs older than this date
 * @param {string} [options.level] - Delete logs of this level
 * @param {string} [options.category] - Delete logs in this category
 * @returns {Promise<number>} Number of logs deleted
 */
async function clearUserLogs(userId, options = {}, sessionId = null) {
  const startTime = Date.now();

  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('User logs clear operation started', {
        userId,
        options: JSON.stringify(options),
        sessionId,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    if (!userId) {
      throw new Error('userId is required for clearing user logs');
    }

    const { 
      olderThan = null, 
      level = null, 
      category = null 
    } = options;
    
    const db = await getConnection();
    
    // Build the WHERE clause and params dynamically
    let whereConditions = ['user_id = ?'];
    let params = [userId];
    
    if (olderThan) {
      whereConditions.push('timestamp < ?');
      params.push(olderThan.toISOString());
    }
    
    if (level) {
      whereConditions.push('level = ?');
      params.push(level);
    }
    
    if (category) {
      whereConditions.push('category = ?');
      params.push(category);
    }
    
    // Create the final WHERE clause
    const whereClause = whereConditions.join(' AND ');
    
    // Build delete query
    const query = `DELETE FROM user_logs WHERE ${whereClause}`;
    
    const result = await db.query(query, params);
    db.release();
    
    // Get number of rows deleted
    let deletedRows = 0;
    if (result.changes !== undefined) {
      // SQLite
      deletedRows = result.changes;
    } else if (result.affectedRows !== undefined) {
      // MySQL
      deletedRows = result.affectedRows;
    } else if (result.rowCount !== undefined) {
      // PostgreSQL
      deletedRows = result.rowCount;
    }
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (userId) {
      MonitoringService.info('User logs cleared successfully', {
        deletedCount: deletedRows,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.info('User logs cleared with session', {
        sessionId,
        deletedCount: deletedRows,
        filters: JSON.stringify(options),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_clear_user_logs_success', executionTime, {
      userId: userId,
      deletedCount: deletedRows,
      filters: JSON.stringify(options),
      timestamp: new Date().toISOString()
    }, 'storage');
    
    return deletedRows;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'storage',
      `Failed to clear user logs: ${error.message}`,
      'error',
      { 
        userId, 
        options, 
        sessionId,
        error: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (userId) {
      MonitoringService.error('User logs clear failed', {
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage', null, userId);
    } else if (sessionId) {
      MonitoringService.error('User logs clear failed', {
        sessionId,
        filters: JSON.stringify(options),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'storage');
    }
    
    MonitoringService.trackMetric('storage_clear_user_logs_failure', executionTime, {
      userId: userId,
      filters: JSON.stringify(options),
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'storage');
    
    throw mcpError;
  }
}

async function addHistory(event, payload, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Adding history entry', {
                event,
                payloadSize: JSON.stringify(payload).length,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        await connection.query('INSERT INTO history (event, payload) VALUES (?, ?)', [event, JSON.stringify(payload)]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('History entry added successfully', {
                event,
                payloadSize: JSON.stringify(payload).length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('History entry added with session', {
                sessionId,
                event,
                payloadSize: JSON.stringify(payload).length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_add_history_success', executionTime, {
            event,
            timestamp: new Date().toISOString()
        });
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to add history entry: ${error.message}`,
            'error',
            {
                event,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('History entry addition failed', {
                event,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('History entry addition failed', {
                sessionId,
                event,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_add_history_failure', executionTime, {
            event,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function getHistory(limit = 50, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting history entries', {
                limit,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        const rows = await connection.query('SELECT event, payload, ts FROM history ORDER BY ts DESC LIMIT ?', [limit]);
        connection.release();
        
        const result = rows && rows.length > 0 ? rows.map(row => ({ event: row.event, payload: JSON.parse(row.payload), ts: row.ts })) : [];
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('History entries retrieved successfully', {
                limit,
                resultCount: result.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('History entries retrieved with session', {
                sessionId,
                limit,
                resultCount: result.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_get_history_success', executionTime, {
            limit,
            resultCount: result.length,
            timestamp: new Date().toISOString()
        });
        
        return result;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to get history: ${error.message}`,
            'error',
            {
                limit,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('History retrieval failed', {
                limit,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('History retrieval failed', {
                sessionId,
                limit,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        MonitoringService.trackMetric('storage_get_history_failure', executionTime, {
            limit,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        
        throw mcpError;
    }
}

async function registerDevice(deviceId, deviceSecret, deviceCode, userCode, verificationUri, expiresAt, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Device registration operation started', {
                deviceId: deviceId?.substring(0, 8) + '...',
                userCode,
                verificationUri,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        if (!deviceId || !deviceSecret) {
            const mcpError = ErrorService.createError(
                'storage',
                'Device ID and secret are required',
                'error',
                { 
                    deviceId: !!deviceId, 
                    deviceSecret: !!deviceSecret,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }
        
        const connection = await getConnection();
        await connection.query(`INSERT OR REPLACE INTO devices 
            (device_id, device_secret, device_code, user_code, verification_uri, expires_at, created_at, is_authorized) 
            VALUES (?, ?, ?, ?, ?, ?, ?, FALSE)`, 
            [deviceId, deviceSecret, deviceCode, userCode, verificationUri, expiresAt, Date.now()]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Device registered successfully', {
                deviceId: deviceId?.substring(0, 8) + '...',
                userCode,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Device registered with session', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                userCode,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_register_device_success', executionTime, {
            deviceId: deviceId,
            timestamp: new Date().toISOString()
        }, 'storage');
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to register device: ${error.message}`,
            'error',
            {
                deviceId: deviceId,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Device registration failed', {
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Device registration failed', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_register_device_error', executionTime, {
            deviceId: deviceId,
            error: error.message,
            timestamp: new Date().toISOString()
        }, 'storage');
        
        throw mcpError;
    }
}

async function getDevice(deviceId, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Device retrieval operation started', {
                deviceId: deviceId?.substring(0, 8) + '...',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        if (!deviceId) {
            const mcpError = ErrorService.createError(
                'storage',
                'Device ID is required',
                'error',
                { 
                    deviceId: deviceId,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }
        
        const connection = await getConnection();
        const row = await connection.query('SELECT * FROM devices WHERE device_id = ?', [deviceId]);
        connection.release();
        
        const result = row && row.length > 0 ? {
            deviceId: row[0].device_id,
            deviceSecret: row[0].device_secret,
            deviceCode: row[0].device_code,
            userCode: row[0].user_code,
            verificationUri: row[0].verification_uri,
            expiresAt: row[0].expires_at,
            isAuthorized: row[0].is_authorized,
            userId: row[0].user_id,
            createdAt: row[0].created_at,
            lastSeen: row[0].last_seen,
            metadata: row[0].metadata ? JSON.parse(row[0].metadata) : null
        } : null;
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Device retrieved successfully', {
                deviceId: deviceId?.substring(0, 8) + '...',
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Device retrieved with session', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                found: !!result,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        return result;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to get device: ${error.message}`,
            'error',
            {
                deviceId: deviceId,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Device retrieval failed', {
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Device retrieval failed', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        throw mcpError;
    }
}

async function authorizeDevice(deviceId, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Device authorization operation started', {
                deviceId: deviceId?.substring(0, 8) + '...',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        if (!deviceId || !userId) {
            const mcpError = ErrorService.createError(
                'storage',
                'Device ID and user ID are required',
                'error',
                { 
                    deviceId: !!deviceId, 
                    userId: !!userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }
        
        const connection = await getConnection();
        await connection.query('UPDATE devices SET is_authorized = TRUE, user_id = ?, last_seen = ? WHERE device_id = ?', 
            [userId, Date.now(), deviceId]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        MonitoringService.info('Device authorized successfully', {
            deviceId: deviceId?.substring(0, 8) + '...',
            userId,
            executionTimeMs: executionTime,
            timestamp: new Date().toISOString()
        }, 'storage', null, userId);
        
        MonitoringService.trackMetric('storage_authorize_device_success', executionTime, {
            deviceId: deviceId,
            userId: userId,
            timestamp: new Date().toISOString()
        }, 'storage');
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to authorize device: ${error.message}`,
            'error',
            {
                deviceId: deviceId,
                userId: userId,
                stack: error.stack,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        MonitoringService.error('Device authorization failed', {
            deviceId: deviceId?.substring(0, 8) + '...',
            userId,
            error: error.message,
            executionTimeMs: executionTime,
            timestamp: new Date().toISOString()
        }, 'storage', null, userId);
        
        MonitoringService.trackMetric('storage_authorize_device_error', executionTime, {
            deviceId: deviceId,
            userId: userId,
            error: error.message,
            timestamp: new Date().toISOString()
        }, 'storage');
        
        throw mcpError;
    }
}

async function updateDeviceMetadata(deviceId, metadata, userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Device metadata update operation started', {
                deviceId: deviceId?.substring(0, 8) + '...',
                hasMetadata: !!metadata,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        if (!deviceId) {
            const mcpError = ErrorService.createError(
                'storage',
                'Device ID is required',
                'error',
                { 
                    deviceId: deviceId,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }
        
        const connection = await getConnection();
        await connection.query('UPDATE devices SET metadata = ?, last_seen = ? WHERE device_id = ?', 
            [JSON.stringify(metadata), Date.now(), deviceId]);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Device metadata updated successfully', {
                deviceId: deviceId?.substring(0, 8) + '...',
                hasMetadata: !!metadata,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Device metadata updated with session', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                hasMetadata: !!metadata,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_update_device_metadata_success', executionTime, {
            deviceId: deviceId,
            timestamp: new Date().toISOString()
        }, 'storage');
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Failed to update device metadata: ${error.message}`,
            'error',
            {
                deviceId: deviceId,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Device metadata update failed', {
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Device metadata update failed', {
                sessionId,
                deviceId: deviceId?.substring(0, 8) + '...',
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        MonitoringService.trackMetric('storage_update_device_metadata_error', executionTime, {
            deviceId: deviceId,
            error: error.message,
            timestamp: new Date().toISOString()
        }, 'storage');
        
        throw mcpError;
    }
}

/**
 * Get backup manager instance
 */
function getBackupManager() {
    if (!backupManager) {
        throw new Error('Storage service not initialized. Call init() first.');
    }
    return backupManager;
}

/**
 * Get migration manager instance
 */
function getMigrationManager() {
    if (!migrationManager) {
        throw new Error('Storage service not initialized. Call init() first.');
    }
    return migrationManager;
}

/**
 * Health check for storage service
 */
async function healthCheck(userId, sessionId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Storage health check operation started', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        const connection = await getConnection();
        const result = await connection.query('SELECT 1 as test', []);
        connection.release();
        
        const executionTime = Date.now() - startTime;
        const healthResult = {
            status: 'healthy',
            database: databaseFactory.getConfig().type,
            timestamp: new Date().toISOString()
        };
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Storage health check completed successfully', {
                status: 'healthy',
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Storage health check completed with session', {
                sessionId,
                status: 'healthy',
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        return healthResult;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'storage',
            `Storage health check failed: ${error.message}`,
            'error',
            {
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Storage health check failed', {
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Storage health check failed', {
                sessionId,
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'storage');
        }
        
        return {
            status: 'unhealthy',
            error: error.message,
            timestamp: new Date().toISOString()
        };
    }
}

module.exports = {
    init,
    getConnection,
    setSetting,
    getSetting,
    setSecureSetting: setSecure,
    getSecureSetting: getSecure,
    deleteSecureSetting: deleteSecure,
    addHistory,
    getHistory,
    registerDevice,
    getDevice,
    authorizeDevice,
    updateDeviceMetadata,
    getBackupManager,
    getMigrationManager,
    healthCheck,
    // User log management methods
    addUserLog,
    getUserLogs,
    countUserLogs,
    clearUserLogs
};
