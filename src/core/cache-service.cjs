/**
 * @fileoverview In-memory cache service with TTL and async API for MCP Desktop.
 * Provides get, set, invalidate, and stats methods. Follows project async and modular rules.
 */

const ErrorService = require('./error-service.cjs');
const MonitoringService = require('./monitoring-service.cjs');

const DEFAULT_TTL = 3600; // seconds

// Log service initialization
MonitoringService.info('Cache service initialized', {
    serviceName: 'cache-service',
    defaultTtl: DEFAULT_TTL,
    timestamp: new Date().toISOString()
}, 'cache');

class CacheService {
    constructor() {
        this._cache = new Map();
        this._timeouts = new Map();
        this._stats = { hits: 0, misses: 0, sets: 0, deletes: 0 };
    }

    /**
     * Async get from cache. Returns null if missing or expired.
     * @param {string} key
     * @param {string} [userId] - Optional user ID for user context logging
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<any|null>}
     */
    async get(key, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Cache get operation started', {
                key,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            const entry = this._cache.get(key);
            
            if (!entry) {
                this._stats.misses++;
                const executionTime = Date.now() - startTime;
                
                MonitoringService.trackMetric('cache_miss', executionTime, {
                    key,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                });
                
                return null;
            }
            
            if (Date.now() > entry.expiry) {
                await this.invalidate(key, userId, sessionId);
                this._stats.misses++;
                const executionTime = Date.now() - startTime;
                
                MonitoringService.trackMetric('cache_expired', executionTime, {
                    key,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                });
                
                return null;
            }
            
            this._stats.hits++;
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('Cache get completed successfully', {
                    key,
                    hit: true,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Cache get completed with session', {
                    sessionId,
                    key,
                    hit: true,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_hit', executionTime, {
                key,
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            return entry.value;
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `Cache get operation failed: ${error.message}`,
                'error',
                {
                    key,
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Cache get failed', {
                    key,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Cache get failed', {
                    sessionId,
                    key,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_get_failure', executionTime, {
                key,
                errorType: error.code || 'unknown',
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Async set in cache with TTL (seconds).
     * @param {string} key
     * @param {any} value
     * @param {number} [ttl=DEFAULT_TTL]
     * @param {string} [userId] - Optional user ID for user context logging
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async set(key, value, ttl = DEFAULT_TTL, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Cache set operation started', {
                key,
                ttl,
                valueType: typeof value,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            await this.invalidate(key, userId, sessionId);
            const expiry = Date.now() + ttl * 1000;
            this._cache.set(key, { value, expiry });
            this._stats.sets++;
            
            // Schedule expiration cleanup
            const timeout = setTimeout(() => {
                this.invalidate(key, userId, sessionId);
            }, ttl * 1000);
            this._timeouts.set(key, timeout);
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('Cache set completed successfully', {
                    key,
                    ttl,
                    duration: executionTime,
                    cacheSize: this._cache.size,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Cache set completed with session', {
                    sessionId,
                    key,
                    ttl,
                    duration: executionTime,
                    cacheSize: this._cache.size,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_set_success', executionTime, {
                key,
                ttl,
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            // Track cache size metrics
            MonitoringService.trackMetric('cache_size', this._cache.size, {
                operation: 'set',
                timestamp: new Date().toISOString()
            });
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `Cache set operation failed: ${error.message}`,
                'error',
                {
                    key,
                    ttl,
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Cache set failed', {
                    key,
                    ttl,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Cache set failed', {
                    sessionId,
                    key,
                    ttl,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_set_failure', executionTime, {
                key,
                errorType: error.code || 'unknown',
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Remove a key from cache.
     * @param {string} key
     * @param {string} [userId] - Optional user ID for user context logging
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async invalidate(key, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Cache invalidate operation started', {
                key,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            let wasDeleted = false;
            
            if (this._cache.has(key)) {
                this._cache.delete(key);
                this._stats.deletes++;
                wasDeleted = true;
            }
            
            if (this._timeouts.has(key)) {
                clearTimeout(this._timeouts.get(key));
                this._timeouts.delete(key);
            }
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('Cache invalidate completed successfully', {
                    key,
                    wasDeleted,
                    duration: executionTime,
                    cacheSize: this._cache.size,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Cache invalidate completed with session', {
                    sessionId,
                    key,
                    wasDeleted,
                    duration: executionTime,
                    cacheSize: this._cache.size,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_invalidate_success', executionTime, {
                key,
                wasDeleted,
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            // Track cache size metrics
            MonitoringService.trackMetric('cache_size', this._cache.size, {
                operation: 'invalidate',
                timestamp: new Date().toISOString()
            });
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `Cache invalidate operation failed: ${error.message}`,
                'error',
                {
                    key,
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Cache invalidate failed', {
                    key,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Cache invalidate failed', {
                    sessionId,
                    key,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_invalidate_failure', executionTime, {
                key,
                errorType: error.code || 'unknown',
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Get cache stats.
     * @param {string} [userId] - Optional user ID for user context logging
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<object>}
     */
    async stats(userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Cache stats operation started', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            const stats = { 
                ...this._stats, 
                size: this._cache.size,
                hitRatio: this._stats.hits + this._stats.misses > 0 ? 
                    (this._stats.hits / (this._stats.hits + this._stats.misses)).toFixed(3) : 0
            };
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('Cache stats retrieved successfully', {
                    stats,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Cache stats retrieved with session', {
                    sessionId,
                    stats,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_stats_success', executionTime, {
                cacheSize: stats.size,
                hitRatio: stats.hitRatio,
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            // Log cache performance periodically
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Cache statistics retrieved', {
                    stats,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            return stats;
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `Cache stats operation failed: ${error.message}`,
                'error',
                {
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Cache stats failed', {
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Cache stats failed', {
                    sessionId,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_stats_failure', executionTime, {
                errorType: error.code || 'unknown',
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Clear all cache entries.
     * @param {string} [userId] - Optional user ID for user context logging
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async clear(userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Cache clear operation started', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            const initialSize = this._cache.size;
            const initialTimeouts = this._timeouts.size;

            MonitoringService.info('Cache clear operation started', {
                initialSize,
                timestamp: new Date().toISOString()
            }, 'cache');

            // PERF-1: Bulk clear instead of N+1 invalidate calls
            // Clear all timeouts first to prevent orphaned callbacks
            for (const timeout of this._timeouts.values()) {
                clearTimeout(timeout);
            }
            this._timeouts.clear();
            this._cache.clear();
            this._stats.deletes += initialSize;

            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('Cache clear completed successfully', {
                    clearedEntries: initialSize,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Cache clear completed with session', {
                    sessionId,
                    clearedEntries: initialSize,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_clear_success', executionTime, {
                clearedEntries: initialSize,
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            MonitoringService.info('Cache cleared successfully', {
                clearedEntries: initialSize,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'cache');
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `Cache clear operation failed: ${error.message}`,
                'error',
                {
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Cache clear failed', {
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Cache clear failed', {
                    sessionId,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.trackMetric('cache_clear_failure', executionTime, {
                errorType: error.code || 'unknown',
                duration: executionTime,
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Generate user-scoped cache key for multi-user isolation
     * @param {string} userId - User ID
     * @param {string} key - Base cache key
     * @returns {string} User-scoped cache key
     */
    _getUserScopedKey(userId, key) {
        return userId ? `user:${userId}:${key}` : key;
    }

    /**
     * Get from cache with user scope for multi-user isolation
     * @param {string} userId - User ID for scoping
     * @param {string} key - Cache key
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<any|null>}
     */
    async getUserScoped(userId, key, sessionId) {
        const scopedKey = this._getUserScopedKey(userId, key);
        return this.get(scopedKey, userId, sessionId);
    }

    /**
     * Set in cache with user scope for multi-user isolation
     * @param {string} userId - User ID for scoping
     * @param {string} key - Cache key
     * @param {any} value - Value to cache
     * @param {number} [ttl=DEFAULT_TTL] - Time to live in seconds
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async setUserScoped(userId, key, value, ttl = DEFAULT_TTL, sessionId) {
        const scopedKey = this._getUserScopedKey(userId, key);
        return this.set(scopedKey, value, ttl, userId, sessionId);
    }

    /**
     * Invalidate cache entry with user scope for multi-user isolation
     * @param {string} userId - User ID for scoping
     * @param {string} key - Cache key
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async invalidateUserScoped(userId, key, sessionId) {
        const scopedKey = this._getUserScopedKey(userId, key);
        return this.invalidate(scopedKey, userId, sessionId);
    }

    /**
     * Clear all cache entries for a specific user
     * @param {string} userId - User ID to clear cache for
     * @param {string} [sessionId] - Optional session ID for fallback logging
     * @returns {Promise<void>}
     */
    async clearUserScoped(userId, sessionId) {
        if (!userId) return;
        
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('User-scoped cache clear operation started', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'cache');
        }
        
        try {
            const userPrefix = `user:${userId}:`;
            const keysToDelete = [];
            
            for (const key of this._cache.keys()) {
                if (key.startsWith(userPrefix)) {
                    keysToDelete.push(key);
                }
            }
            
            for (const key of keysToDelete) {
                await this.invalidate(key, userId, sessionId);
            }
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs (successful operations)
            if (userId) {
                MonitoringService.info('User-scoped cache clear completed successfully', {
                    userId,
                    keysCleared: keysToDelete.length,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.info('User-scoped cache clear completed with session', {
                    sessionId,
                    keysCleared: keysToDelete.length,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            MonitoringService.info('User-scoped cache cleared', {
                userId,
                keysCleared: keysToDelete.length,
                duration: executionTime,
                timestamp: new Date().toISOString()
            }, 'cache');
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'cache',
                `User-scoped cache clear operation failed: ${error.message}`,
                'error',
                {
                    userId,
                    error: error.message,
                    stack: error.stack,
                    duration: executionTime,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('User-scoped cache clear failed', {
                    userId,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache', null, userId);
            } else if (sessionId) {
                MonitoringService.error('User-scoped cache clear failed', {
                    sessionId,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'cache');
            }
            
            throw mcpError;
        }
    }
}

module.exports = new CacheService();
