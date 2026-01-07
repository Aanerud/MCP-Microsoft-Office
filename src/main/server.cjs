/**
 * @fileoverview Local Express server for MCP API.
 * Sets up RESTful endpoints, middleware, error handling, and logging.
 * Follows MCP modularity, async, and testable design.
 */

const express = require('express');
const cors = require('cors');
const morgan = require('morgan');
const bodyParser = require('body-parser');
const rateLimit = require('express-rate-limit');
const errorService = require('../core/error-service.cjs');
const monitoringService = require('../core/monitoring-service.cjs');

const app = express();

/**
 * Sets up middleware for the Express app.
 * @param {express.Application} expressApp - The Express app to set up middleware for
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 */
function setupMiddleware(expressApp, userId, sessionId) {
    const startTime = new Date().toISOString();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            monitoringService.debug('Setting up Express middleware', {
                sessionId: sessionId,
                timestamp: startTime,
                expressAppType: typeof expressApp
            }, 'server');
        }
    // Configure CORS for production and development
    const corsOptions = {
        origin: function (origin, callback) {
            // Allow requests with no origin (like mobile apps or curl requests)
            if (!origin) return callback(null, true);
            
            // Get allowed origins from environment or use defaults
            const allowedOrigins = process.env.CORS_ALLOWED_ORIGINS 
                ? process.env.CORS_ALLOWED_ORIGINS.split(',').map(origin => origin.trim())
                : [
                    'http://localhost:3000',
                    'https://localhost:3000',
                    'http://127.0.0.1:3000',
                    'https://127.0.0.1:3000'
                ];
            
            // In development, allow all localhost origins
            if (process.env.NODE_ENV !== 'production') {
                const localhostPattern = /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/;
                if (localhostPattern.test(origin)) {
                    return callback(null, true);
                }
            }
            
            // Check if origin is in allowed list
            if (allowedOrigins.includes(origin)) {
                callback(null, true);
            } else {
                monitoringService?.warn('CORS blocked request from unauthorized origin', { 
                    origin,
                    allowedOrigins,
                    userAgent: 'N/A' // Will be available in req context
                }, 'security');
                callback(new Error('Not allowed by CORS'));
            }
        },
        credentials: true, // Allow cookies and authorization headers
        methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS', 'PATCH'],
        allowedHeaders: [
            'Origin',
            'X-Requested-With',
            'Content-Type',
            'Accept',
            'Authorization',
            'Cache-Control',
            'X-API-Key'
        ],
        exposedHeaders: ['X-Total-Count', 'X-Request-ID'],
        maxAge: 86400 // Cache preflight requests for 24 hours
    };
    
    // Middleware
    expressApp.use(cors(corsOptions));
    expressApp.use(bodyParser.json({ limit: '2mb' }));
    // Configure morgan to skip logging for all API endpoints to reduce log volume
    expressApp.use(morgan('dev', {
        skip: (req, res) => req.originalUrl.startsWith('/api/')
    }));

    // Request logging - skip API endpoints to reduce log volume
    expressApp.use((req, res, next) => {
        // Skip logging API requests to reduce log volume and prevent memory issues
        if (!req.originalUrl.startsWith('/api/')) {
            monitoringService.info(`Request: ${req.method} ${req.url}`, { ip: req.ip });
        }
        
        // IMPORTANT: Ensure API routes always return JSON
        if (req.url.startsWith('/api/v1/')) {
            // Make sure Content-Type is application/json for API responses
            const originalJson = res.json;
            res.json = function(body) {
                // Set Content-Type explicitly to ensure proper parsing by the client
                res.setHeader('Content-Type', 'application/json');
                return originalJson.call(this, body);
            };
        }
        
        next();
    });

    // Rate limiting middleware for authentication endpoints
    monitoringService?.info('Setting up rate limiting for authentication endpoints...', {
        authWindowMs: 15 * 60 * 1000,
        authMaxRequests: 100
    }, 'security');
    
    const authLimiter = rateLimit({
        windowMs: 15 * 60 * 1000, // 15 minutes
        max: 100, // Limit each IP to 100 requests per windowMs
        standardHeaders: true, // Return rate limit info in the `RateLimit-*` headers
        legacyHeaders: false, // Disable the `X-RateLimit-*` headers
        handler: (request, response, next) => {
            monitoringService?.warn('Rate limit exceeded for authentication endpoint', {
                ip: request.ip,
                userAgent: request.get('User-Agent'),
                endpoint: request.path
            }, 'security');
            
            const error = new Error('Too many authentication requests, please try again later.');
            error.statusCode = 429;
            next(error);
        }
    });

    // General API rate limiting (more lenient)
    monitoringService?.info('Setting up rate limiting for API endpoints...', {
        apiWindowMs: 15 * 60 * 1000,
        apiMaxRequests: 1000
    }, 'security');
    
    const apiLimiter = rateLimit({
        windowMs: 15 * 60 * 1000, // 15 minutes
        max: 1000, // Limit each IP to 1000 requests per windowMs for general API
        standardHeaders: true,
        legacyHeaders: false,
        handler: (request, response, next) => {
            monitoringService?.warn('Rate limit exceeded for API endpoint', {
                ip: request.ip,
                userAgent: request.get('User-Agent'),
                endpoint: request.path
            }, 'security');
            
            const error = new Error('Too many API requests, please try again later.');
            error.statusCode = 429;
            next(error);
        }
    });

        // Apply rate limiters
        expressApp.use('/api/v1/auth', authLimiter);
        expressApp.use('/api', apiLimiter);

        // Input sanitization middleware
        monitoringService?.info('Setting up input sanitization...', {
            sanitizeBody: true,
            sanitizeQuery: true
        }, 'security');
        
        expressApp.use((req, res, next) => {
            // Sanitize request body
            if (req.body && typeof req.body === 'object') {
                req.body = sanitizeObject(req.body, req?.user?.userId, req.session?.id);
            }
            
            // Sanitize query parameters
            if (req.query && typeof req.query === 'object') {
                req.query = sanitizeObject(req.query, req?.user?.userId, req.session?.id);
            }
            
            next();
        });
        
        // Pattern 2: User Activity Logs
        const endTime = new Date().toISOString();
        if (userId) {
            monitoringService.info('Express middleware setup completed successfully', {
                duration: new Date(endTime) - new Date(startTime),
                timestamp: endTime
            }, 'server', null, userId);
        } else if (sessionId) {
            monitoringService.info('Express middleware setup completed with session', {
                sessionId: sessionId,
                duration: new Date(endTime) - new Date(startTime),
                timestamp: endTime
            }, 'server');
        }
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = errorService.createError(
            'server',
            'Failed to setup Express middleware',
            'error',
            {
                error: error.message,
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
        monitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            monitoringService.error('Express middleware setup failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'server', null, userId);
        } else if (sessionId) {
            monitoringService.error('Express middleware setup failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'server');
        }
        
        throw error;
    }
}

// Helper function to sanitize objects recursively
function sanitizeObject(obj, userId, sessionId) {
    const startTime = new Date().toISOString();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            monitoringService.debug('Sanitizing object', {
                sessionId: sessionId,
                timestamp: startTime,
                objectType: typeof obj,
                isArray: Array.isArray(obj)
            }, 'security');
        }
        
        if (obj === null || obj === undefined) {
            return obj;
        }
        
        if (typeof obj === 'string') {
            return sanitizeString(obj, userId, sessionId);
        }
        
        if (Array.isArray(obj)) {
            return obj.map(item => sanitizeObject(item, userId, sessionId));
        }
        
        if (typeof obj === 'object') {
            const sanitized = {};
            for (const [key, value] of Object.entries(obj)) {
                // Sanitize both key and value
                const cleanKey = sanitizeString(key, userId, sessionId);
                sanitized[cleanKey] = sanitizeObject(value, userId, sessionId);
            }
            
            // Pattern 2: User Activity Logs (only for successful operations)
            const endTime = new Date().toISOString();
            if (userId) {
                monitoringService.info('Object sanitization completed successfully', {
                    objectKeys: Object.keys(sanitized).length,
                    duration: new Date(endTime) - new Date(startTime),
                    timestamp: endTime
                }, 'security', null, userId);
            } else if (sessionId) {
                monitoringService.info('Object sanitization completed with session', {
                    sessionId: sessionId,
                    objectKeys: Object.keys(sanitized).length,
                    duration: new Date(endTime) - new Date(startTime),
                    timestamp: endTime
                }, 'security');
            }
            
            return sanitized;
        }
        
        return obj;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = errorService.createError(
            'security',
            'Failed to sanitize object',
            'error',
            {
                error: error.message,
                stack: error.stack,
                objectType: typeof obj,
                timestamp: new Date().toISOString()
            }
        );
        monitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            monitoringService.error('Object sanitization failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'security', null, userId);
        } else if (sessionId) {
            monitoringService.error('Object sanitization failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'security');
        }
        
        throw error;
    }
}

// Helper function to sanitize strings
function sanitizeString(str, userId, sessionId) {
    const startTime = new Date().toISOString();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            monitoringService.debug('Sanitizing string', {
                sessionId: sessionId,
                timestamp: startTime,
                stringType: typeof str,
                stringLength: typeof str === 'string' ? str.length : 0
            }, 'security');
        }
        
        if (typeof str !== 'string') {
            return str;
        }
        
        // Remove potential XSS patterns
        const sanitized = str
            .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '') // Remove script tags
            .replace(/javascript:/gi, '') // Remove javascript: protocol
            .replace(/on\w+\s*=/gi, '') // Remove event handlers like onclick=
            .replace(/data:text\/html/gi, '') // Remove data URLs with HTML
            .trim();
        
        // Pattern 2: User Activity Logs (only for successful operations with changes)
        const endTime = new Date().toISOString();
        if (sanitized !== str) {
            if (userId) {
                monitoringService.info('String sanitization completed with changes', {
                    originalLength: str.length,
                    sanitizedLength: sanitized.length,
                    duration: new Date(endTime) - new Date(startTime),
                    timestamp: endTime
                }, 'security', null, userId);
            } else if (sessionId) {
                monitoringService.info('String sanitization completed with changes', {
                    sessionId: sessionId,
                    originalLength: str.length,
                    sanitizedLength: sanitized.length,
                    duration: new Date(endTime) - new Date(startTime),
                    timestamp: endTime
                }, 'security');
            }
        }
        
        return sanitized;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = errorService.createError(
            'security',
            'Failed to sanitize string',
            'error',
            {
                error: error.message,
                stack: error.stack,
                stringType: typeof str,
                timestamp: new Date().toISOString()
            }
        );
        monitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            monitoringService.error('String sanitization failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'security', null, userId);
        } else if (sessionId) {
            monitoringService.error('String sanitization failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'security');
        }
        
        throw error;
    }
}

// Set up middleware for the main app
setupMiddleware(app);

// Health endpoints are now handled by the mainApiRouter

// API routes
const { registerRoutes } = require('../api/routes.cjs');
const statusRouter = require('../api/status.cjs');

// Create a main API router to handle all API endpoints
const mainApiRouter = express.Router();

// Mount the status router directly on the main API router
mainApiRouter.use('/status', statusRouter);

// Create a versioned API router for v1 endpoints
const apiRouter = express.Router();
registerRoutes(apiRouter);

// Mount the versioned API router on the main API router
mainApiRouter.use('/', apiRouter);

// Mount the main API router at /api
app.use('/api', mainApiRouter);

// Add a direct health endpoint for better discoverability
mainApiRouter.get('/health', (req, res) => {
    const startTime = new Date().toISOString();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            monitoringService.debug('Health check endpoint accessed', {
                sessionId: req.session?.id,
                userAgent: req.get('User-Agent'),
                timestamp: startTime,
                ip: req.ip
            }, 'server');
        }
        
        const healthStatus = { status: 'ok', ts: new Date().toISOString() };
        
        // Pattern 2: User Activity Logs
        const endTime = new Date().toISOString();
        const userId = req?.user?.userId;
        if (userId) {
            monitoringService.info('Health check completed successfully', {
                status: healthStatus.status,
                duration: new Date(endTime) - new Date(startTime),
                timestamp: endTime
            }, 'server', null, userId);
        } else if (req.session?.id) {
            monitoringService.info('Health check completed with session', {
                sessionId: req.session.id,
                status: healthStatus.status,
                duration: new Date(endTime) - new Date(startTime),
                timestamp: endTime
            }, 'server');
        }
        
        res.json(healthStatus);
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = errorService.createError(
            'server',
            'Health check endpoint failed',
            'error',
            {
                error: error.message,
                stack: error.stack,
                endpoint: '/health',
                timestamp: new Date().toISOString()
            }
        );
        monitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        const userId = req?.user?.userId;
        if (userId) {
            monitoringService.error('Health check failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'server', null, userId);
        } else if (req.session?.id) {
            monitoringService.error('Health check failed', {
                sessionId: req.session.id,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'server');
        }
        
        res.status(500).json({ error: 'Health check failed', ts: new Date().toISOString() });
    }
});

// Error handling middleware
app.use((err, req, res, next) => {
    const startTime = new Date().toISOString();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        monitoringService.debug('Global error handler triggered', {
            sessionId: req.session?.id,
            userAgent: req.get('User-Agent'),
            timestamp: startTime,
            errorType: err.name,
            url: req.url,
            method: req.method
        }, 'server');
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = errorService.createError(
        'server',
        err.message || 'Internal server error',
        'error',
        { 
            stack: err.stack, 
            url: req.url, 
            method: req.method,
            statusCode: err.statusCode || 500,
            timestamp: new Date().toISOString()
        }
    );
    monitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    const userId = req?.user?.userId;
    if (userId) {
        monitoringService.error('Request processing failed', {
            error: err.message || 'Internal server error',
            url: req.url,
            method: req.method,
            timestamp: new Date().toISOString()
        }, 'server', null, userId);
    } else if (req.session?.id) {
        monitoringService.error('Request processing failed', {
            sessionId: req.session.id,
            error: err.message || 'Internal server error',
            url: req.url,
            method: req.method,
            timestamp: new Date().toISOString()
        }, 'server');
    }
    
    res.status(err.statusCode || 500).json({ error: 'Internal server error' });
});

// Server lifecycle management
let server = null;
function startServer(port = 3001, userId, sessionId) {
    const startTime = new Date().toISOString();
    
    return new Promise((resolve, reject) => {
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Starting Express server', {
                    sessionId: sessionId,
                    timestamp: startTime,
                    port: port,
                    nodeEnv: process.env.NODE_ENV
                }, 'server');
            }
            
            server = app.listen(port, () => {
                // Pattern 2: User Activity Logs
                const endTime = new Date().toISOString();
                if (userId) {
                    monitoringService.info('API server started successfully', {
                        port: port,
                        duration: new Date(endTime) - new Date(startTime),
                        timestamp: endTime
                    }, 'server', null, userId);
                } else if (sessionId) {
                    monitoringService.info('API server started with session', {
                        sessionId: sessionId,
                        port: port,
                        duration: new Date(endTime) - new Date(startTime),
                        timestamp: endTime
                    }, 'server');
                } else {
                    monitoringService.info(`API server started on port ${port}`, {
                        port: port,
                        timestamp: endTime
                    }, 'server');
                }
                
                resolve(server);
            });
            
            server.on('error', (error) => {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = errorService.createError(
                    'server',
                    'Failed to start Express server',
                    'error',
                    {
                        error: error.message,
                        stack: error.stack,
                        port: port,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('Server startup failed', {
                        error: error.message,
                        port: port,
                        timestamp: new Date().toISOString()
                    }, 'server', null, userId);
                } else if (sessionId) {
                    monitoringService.error('Server startup failed', {
                        sessionId: sessionId,
                        error: error.message,
                        port: port,
                        timestamp: new Date().toISOString()
                    }, 'server');
                }
                
                reject(error);
            });
            
        } catch (error) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = errorService.createError(
                'server',
                'Exception during server startup',
                'error',
                {
                    error: error.message,
                    stack: error.stack,
                    port: port,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Server startup exception', {
                    error: error.message,
                    port: port,
                    timestamp: new Date().toISOString()
                }, 'server', null, userId);
            } else if (sessionId) {
                monitoringService.error('Server startup exception', {
                    sessionId: sessionId,
                    error: error.message,
                    port: port,
                    timestamp: new Date().toISOString()
                }, 'server');
            }
            
            reject(error);
        }
    });
}
function stopServer(userId, sessionId) {
    const startTime = new Date().toISOString();
    
    return new Promise((resolve, reject) => {
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Stopping Express server', {
                    sessionId: sessionId,
                    timestamp: startTime,
                    hasServer: !!server
                }, 'server');
            }
            
            if (server) {
                server.close((err) => {
                    if (err) {
                        // Pattern 3: Infrastructure Error Logging
                        const mcpError = errorService.createError(
                            'server',
                            'Failed to stop Express server',
                            'error',
                            {
                                error: err.message,
                                stack: err.stack,
                                timestamp: new Date().toISOString()
                            }
                        );
                        monitoringService.logError(mcpError);
                        
                        // Pattern 4: User Error Tracking
                        if (userId) {
                            monitoringService.error('Server shutdown failed', {
                                error: err.message,
                                timestamp: new Date().toISOString()
                            }, 'server', null, userId);
                        } else if (sessionId) {
                            monitoringService.error('Server shutdown failed', {
                                sessionId: sessionId,
                                error: err.message,
                                timestamp: new Date().toISOString()
                            }, 'server');
                        }
                        
                        reject(err);
                    } else {
                        // Pattern 2: User Activity Logs
                        const endTime = new Date().toISOString();
                        if (userId) {
                            monitoringService.info('Express server stopped successfully', {
                                duration: new Date(endTime) - new Date(startTime),
                                timestamp: endTime
                            }, 'server', null, userId);
                        } else if (sessionId) {
                            monitoringService.info('Express server stopped with session', {
                                sessionId: sessionId,
                                duration: new Date(endTime) - new Date(startTime),
                                timestamp: endTime
                            }, 'server');
                        } else {
                            monitoringService.info('Express server stopped', {
                                timestamp: endTime
                            }, 'server');
                        }
                        
                        resolve();
                    }
                });
            } else {
                // Pattern 2: User Activity Logs (server already stopped)
                const endTime = new Date().toISOString();
                if (userId) {
                    monitoringService.info('Server stop requested but server not running', {
                        duration: new Date(endTime) - new Date(startTime),
                        timestamp: endTime
                    }, 'server', null, userId);
                } else if (sessionId) {
                    monitoringService.info('Server stop requested but server not running', {
                        sessionId: sessionId,
                        duration: new Date(endTime) - new Date(startTime),
                        timestamp: endTime
                    }, 'server');
                }
                
                resolve();
            }
            
        } catch (error) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = errorService.createError(
                'server',
                'Exception during server shutdown',
                'error',
                {
                    error: error.message,
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Server shutdown exception', {
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'server', null, userId);
            } else if (sessionId) {
                monitoringService.error('Server shutdown exception', {
                    sessionId: sessionId,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'server');
            }
            
            reject(error);
        }
    });
}

module.exports = { app, startServer, stopServer, setupMiddleware };
