/**
 * @fileoverview Database Factory for MCP Server - Production Database Support
 * Provides abstraction layer for SQLite, PostgreSQL, and MySQL databases with connection pooling.
 */

const path = require('path');
const fs = require('fs');
const ErrorService = require('./error-service.cjs');
const MonitoringService = require('./monitoring-service.cjs');

// Database drivers
let sqlite3, pg, mysql2;

// Connection pools
let pgPool = null;
let mysqlPool = null;

// SQLite singleton connection (SQLite works best with a single connection)
let sqliteDb = null;
let sqliteDbReady = false;

// SQLite retry configuration for SQLITE_BUSY errors
const SQLITE_RETRY_COUNT = 5;
const SQLITE_RETRY_BASE_DELAY_MS = 50;

/**
 * Database connection interface
 */
class DatabaseConnection {
  constructor(type, connection) {
    this.type = type;
    this.connection = connection;
  }

  async query(sql, params = [], userId, sessionId) {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Executing database query', {
          dbType: this.type,
          sqlPreview: sql.substring(0, 100) + (sql.length > 100 ? '...' : ''),
          paramCount: params.length,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      let result;
      switch (this.type) {
        case 'sqlite':
          result = await this.querySQLite(sql, params, userId, sessionId);
          break;
        case 'postgresql':
          result = await this.queryPostgreSQL(sql, params, userId, sessionId);
          break;
        case 'mysql':
          result = await this.queryMySQL(sql, params, userId, sessionId);
          break;
        default:
          throw new Error(`Unsupported database type: ${this.type}`);
      }
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Database query executed successfully', {
          dbType: this.type,
          duration: Date.now() - startTime,
          resultCount: Array.isArray(result) ? result.length : (result?.changes || 0),
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Database query executed with session', {
          sessionId,
          dbType: this.type,
          duration: Date.now() - startTime,
          resultCount: Array.isArray(result) ? result.length : (result?.changes || 0),
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      return result;
      
    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Database query failed: ${error.message}`,
        'error',
        {
          dbType: this.type,
          sqlPreview: sql.substring(0, 100) + (sql.length > 100 ? '...' : ''),
          paramCount: params.length,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Database query execution failed', {
          error: error.message,
          dbType: this.type,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database query execution failed', {
          sessionId,
          error: error.message,
          dbType: this.type,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  async querySQLite(sql, params, userId, sessionId) {
    // Retry logic for SQLITE_BUSY errors
    for (let attempt = 0; attempt < SQLITE_RETRY_COUNT; attempt++) {
      try {
        return await new Promise((resolve, reject) => {
          if (sql.toLowerCase().startsWith('select') || sql.toLowerCase().startsWith('pragma')) {
            this.connection.all(sql, params, (err, rows) => {
              if (err) reject(err);
              else resolve(rows);
            });
          } else {
            this.connection.run(sql, params, function(err) {
              if (err) reject(err);
              else resolve({ changes: this.changes, lastID: this.lastID });
            });
          }
        });
      } catch (err) {
        // Check if this is a SQLITE_BUSY error and we have retries left
        if (err.code === 'SQLITE_BUSY' && attempt < SQLITE_RETRY_COUNT - 1) {
          // Exponential backoff: 50ms, 100ms, 200ms, 400ms, 800ms
          const delay = SQLITE_RETRY_BASE_DELAY_MS * Math.pow(2, attempt);
          await new Promise(resolve => setTimeout(resolve, delay));
          continue;
        }
        throw err;
      }
    }
  }

  async queryPostgreSQL(sql, params, userId, sessionId) {
    // Convert SQLite-style ? placeholders to PostgreSQL $1, $2, etc.
    let pgSql = sql;
    let pgParams = params;
    
    if (params.length > 0) {
      let paramIndex = 1;
      pgSql = sql.replace(/\?/g, () => `$${paramIndex++}`);
    }

    const result = await this.connection.query(pgSql, pgParams);
    return result.rows;
  }

  async queryMySQL(sql, params, userId, sessionId) {
    const [rows] = await this.connection.execute(sql, params);
    return rows;
  }

  async close() {
    switch (this.type) {
      case 'sqlite':
        // SQLite uses singleton connection - don't close it
        // The connection will be closed when the factory is closed
        return;
      case 'postgresql':
        // Connection is returned to pool automatically
        this.connection.release();
        break;
      case 'mysql':
        // Connection is returned to pool automatically
        this.connection.release();
        break;
    }
  }

  async release() {
    return this.close();
  }
}

/**
 * Database Factory
 */
class DatabaseFactory {
  constructor() {
    this.config = null;
    this.initialized = false;
  }

  /**
   * Initialize the database factory with configuration
   */
  async init(config, userId, sessionId) {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Initializing database factory', {
          dbType: config.DB_TYPE,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      this.config = config;

      switch (config.DB_TYPE) {
        case 'sqlite':
          await this.initSQLite(userId, sessionId);
          break;
        case 'postgresql':
          await this.initPostgreSQL(userId, sessionId);
          break;
        case 'mysql':
          await this.initMySQL(userId, sessionId);
          break;
        default:
          throw new Error(`Unsupported database type: ${config.DB_TYPE}`);
      }

      this.initialized = true;
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Database factory initialized successfully', {
          dbType: config.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Database factory initialized with session', {
          sessionId,
          dbType: config.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to initialize database factory: ${error.message}`,
        'error',
        {
          dbType: config.DB_TYPE,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Database factory initialization failed', {
          error: error.message,
          dbType: config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database factory initialization failed', {
          sessionId,
          error: error.message,
          dbType: config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  async initSQLite(userId, sessionId) {
    const startTime = Date.now();

    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Initializing SQLite database', {
          dbPath: this.config.DB_PATH,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }

      sqlite3 = require('sqlite3').verbose();

      // Ensure data directory exists
      const dataDir = path.dirname(this.config.DB_PATH);
      if (!fs.existsSync(dataDir)) {
        fs.mkdirSync(dataDir, { recursive: true });
      }

      // Create singleton connection for SQLite
      sqliteDb = new sqlite3.Database(this.config.DB_PATH);

      // Enable WAL mode for better concurrent access
      await new Promise((resolve, reject) => {
        sqliteDb.run('PRAGMA journal_mode=WAL', (err) => {
          if (err) {
            console.warn('[DATABASE] WAL mode not enabled:', err.message);
            resolve(); // Continue even if WAL fails
          } else {
            console.log('[DATABASE] SQLite WAL mode enabled');
            resolve();
          }
        });
      });

      // Set busy timeout to wait for locks instead of failing immediately (5 seconds)
      await new Promise((resolve, reject) => {
        sqliteDb.run('PRAGMA busy_timeout=5000', (err) => {
          if (err) {
            console.warn('[DATABASE] busy_timeout not set:', err.message);
            resolve();
          } else {
            resolve();
          }
        });
      });

      // Enable foreign keys
      await new Promise((resolve, reject) => {
        sqliteDb.run('PRAGMA foreign_keys=ON', (err) => {
          if (err) resolve(); // Continue even if this fails
          else resolve();
        });
      });

      sqliteDbReady = true;

      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('SQLite database initialized successfully', {
          dbPath: this.config.DB_PATH,
          walMode: true,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('SQLite database initialized with session', {
          sessionId,
          dbPath: this.config.DB_PATH,
          walMode: true,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }

    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to initialize SQLite: ${error.message}`,
        'error',
        {
          dbPath: this.config.DB_PATH,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);

      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('SQLite database initialization failed', {
          error: error.message,
          dbPath: this.config.DB_PATH,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('SQLite database initialization failed', {
          sessionId,
          error: error.message,
          dbPath: this.config.DB_PATH,
          timestamp: new Date().toISOString()
        }, 'database');
      }

      throw mcpError;
    }
  }

  async initPostgreSQL(userId, sessionId) {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Initializing PostgreSQL database', {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          port: this.config.DB_PORT,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      const { Pool } = require('pg');
      
      pgPool = new Pool({
        host: this.config.DB_HOST,
        port: this.config.DB_PORT,
        database: this.config.DB_NAME,
        user: this.config.DB_USER,
        password: this.config.DB_PASSWORD,
        max: this.config.DB_POOL_MAX || 20,
        idleTimeoutMillis: this.config.DB_IDLE_TIMEOUT || 30000,
        connectionTimeoutMillis: this.config.DB_CONNECTION_TIMEOUT || 2000,
      });

      // Test connection
      const client = await pgPool.connect();
      await client.query('SELECT NOW()');
      client.release();

      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('PostgreSQL connection pool initialized successfully', {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          maxConnections: this.config.DB_POOL_MAX || 20,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('PostgreSQL connection pool initialized with session', {
          sessionId,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          maxConnections: this.config.DB_POOL_MAX || 20,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }

    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to initialize PostgreSQL: ${error.message}`,
        'error',
        {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          port: this.config.DB_PORT,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('PostgreSQL database initialization failed', {
          error: error.message,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('PostgreSQL database initialization failed', {
          sessionId,
          error: error.message,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  async initMySQL(userId, sessionId) {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Initializing MySQL database', {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          port: this.config.DB_PORT,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      mysql2 = require('mysql2/promise');
      
      mysqlPool = mysql2.createPool({
        host: this.config.DB_HOST,
        port: this.config.DB_PORT,
        database: this.config.DB_NAME,
        user: this.config.DB_USER,
        password: this.config.DB_PASSWORD,
        connectionLimit: this.config.DB_POOL_MAX || 20,
        idleTimeout: this.config.DB_IDLE_TIMEOUT || 30000,
        acquireTimeout: this.config.DB_CONNECTION_TIMEOUT || 2000,
      });

      // Test connection
      const connection = await mysqlPool.getConnection();
      await connection.query('SELECT 1');
      connection.release();

      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('MySQL connection pool initialized successfully', {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          connectionLimit: this.config.DB_POOL_MAX || 20,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('MySQL connection pool initialized with session', {
          sessionId,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          connectionLimit: this.config.DB_POOL_MAX || 20,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }

    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to initialize MySQL: ${error.message}`,
        'error',
        {
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          port: this.config.DB_PORT,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('MySQL database initialization failed', {
          error: error.message,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('MySQL database initialization failed', {
          sessionId,
          error: error.message,
          host: this.config.DB_HOST,
          database: this.config.DB_NAME,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  /**
   * Get a database connection
   */
  async getConnection(userId, sessionId) {
    if (!this.initialized) {
      throw new Error('Database factory not initialized');
    }

    const startTime = Date.now();

    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Acquiring database connection', {
          dbType: this.config.DB_TYPE,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      let connection;

      switch (this.config.DB_TYPE) {
        case 'sqlite':
          // Use singleton connection for SQLite (better for concurrent access with WAL mode)
          if (!sqliteDb || !sqliteDbReady) {
            throw new Error('SQLite singleton connection not initialized');
          }
          connection = sqliteDb;
          break;
        case 'postgresql':
          connection = await pgPool.connect();
          break;
        case 'mysql':
          connection = await mysqlPool.getConnection();
          break;
        default:
          throw new Error(`Unsupported database type: ${this.config.DB_TYPE}`);
      }

      const executionTime = Date.now() - startTime;
      MonitoringService.trackMetric('database_connection_acquired', executionTime, {
        dbType: this.config.DB_TYPE,
        userId,
        sessionId,
        timestamp: new Date().toISOString()
      });

      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Database connection acquired successfully', {
          dbType: this.config.DB_TYPE,
          duration: executionTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Database connection acquired with session', {
          sessionId,
          dbType: this.config.DB_TYPE,
          duration: executionTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }

      return new DatabaseConnection(this.config.DB_TYPE, connection);

    } catch (error) {
      const executionTime = Date.now() - startTime;
      
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to get database connection: ${error.message}`,
        'error',
        { 
          dbType: this.config.DB_TYPE,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      
      MonitoringService.logError(mcpError);
      MonitoringService.trackMetric('database_connection_failed', executionTime, {
        dbType: this.config.DB_TYPE,
        errorType: error.code || 'unknown',
        userId,
        sessionId,
        timestamp: new Date().toISOString()
      });
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Database connection acquisition failed', {
          error: error.message,
          dbType: this.config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database connection acquisition failed', {
          sessionId,
          error: error.message,
          dbType: this.config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  /**
   * Close all database connections and pools
   */
  async close(userId, sessionId) {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Closing database factory', {
          dbType: this.config?.DB_TYPE,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      // Close SQLite singleton connection
      if (sqliteDb) {
        await new Promise((resolve) => {
          sqliteDb.close((err) => {
            if (err) {
              console.warn('[DATABASE] Error closing SQLite:', err.message);
            }
            resolve();
          });
        });
        sqliteDb = null;
        sqliteDbReady = false;
      }

      if (pgPool) {
        await pgPool.end();
        pgPool = null;
      }

      if (mysqlPool) {
        await mysqlPool.end();
        mysqlPool = null;
      }

      this.initialized = false;
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Database factory closed successfully', {
          dbType: this.config?.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Database factory closed with session', {
          sessionId,
          dbType: this.config?.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }

    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Failed to close database factory: ${error.message}`,
        'error',
        {
          dbType: this.config?.DB_TYPE,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Database factory close failed', {
          error: error.message,
          dbType: this.config?.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database factory close failed', {
          sessionId,
          error: error.message,
          dbType: this.config?.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      throw mcpError;
    }
  }

  /**
   * Get database health status
   */
  async getHealthStatus(userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Checking database health status', {
        dbType: this.config?.DB_TYPE,
        initialized: this.initialized,
        userId,
        sessionId,
        timestamp: new Date().toISOString()
      }, 'database');
    }
    
    if (!this.initialized) {
      // Pattern 4: User Error Tracking for not initialized
      if (userId) {
        MonitoringService.error('Database health check failed - not initialized', {
          status: 'not_initialized',
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database health check failed - not initialized', {
          sessionId,
          status: 'not_initialized',
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      return { status: 'not_initialized', healthy: false };
    }

    try {
      const connection = await this.getConnection(userId, sessionId);
      
      // Simple health check query
      switch (this.config.DB_TYPE) {
        case 'sqlite':
          await connection.query('SELECT 1');
          break;
        case 'postgresql':
          await connection.query('SELECT 1');
          break;
        case 'mysql':
          await connection.query('SELECT 1');
          break;
      }
      
      await connection.close();
      
      const healthStatus = {
        status: 'healthy',
        healthy: true,
        dbType: this.config.DB_TYPE,
        duration: Date.now() - startTime,
        timestamp: new Date().toISOString()
      };
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Database health check completed successfully', {
          status: 'healthy',
          dbType: this.config.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Database health check completed with session', {
          sessionId,
          status: 'healthy',
          dbType: this.config.DB_TYPE,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      return healthStatus;

    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'database',
        `Database health check failed: ${error.message}`,
        'error',
        {
          dbType: this.config.DB_TYPE,
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Database health check failed', {
          error: error.message,
          dbType: this.config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Database health check failed', {
          sessionId,
          error: error.message,
          dbType: this.config.DB_TYPE,
          timestamp: new Date().toISOString()
        }, 'database');
      }
      
      return {
        status: 'unhealthy',
        healthy: false,
        error: error.message,
        dbType: this.config.DB_TYPE,
        timestamp: new Date().toISOString()
      };
    }
  }
}

// Singleton instance
const databaseFactory = new DatabaseFactory();

module.exports = {
  DatabaseFactory,
  DatabaseConnection,
  databaseFactory
};
