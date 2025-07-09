/**
 * @fileoverview Registers all API routes and controllers for MCP.
 * Handles versioning, middleware, and route registration.
 */

const express = require('express');
const queryControllerFactory = require('./controllers/query-controller.js');
const mailControllerFactory = require('./controllers/mail-controller.js');
const calendarControllerFactory = require('./controllers/calendar-controller.js');
const filesControllerFactory = require('./controllers/files-controller.js');
const peopleControllerFactory = require('./controllers/people-controller.cjs');
const logController = require('./controllers/log-controller.cjs');
const authController = require('./controllers/auth-controller.cjs');
const deviceAuthController = require('./controllers/device-auth-controller.cjs');
const adapterController = require('./controllers/adapter-controller.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const { requireAuth } = require('./middleware/auth-middleware.cjs');
const { routesLogger, controllerLogger } = require('./middleware/request-logger.cjs');
const apiContext = require('./api-context.cjs');
const statusRouter = require('./status.cjs');

/**
 * TODO: [Rate Limiting] Implement and configure rate limiting middleware
 * const rateLimiter = require('express-rate-limit');
 * const postLimiter = rateLimiter({ windowMs: 15 * 60 * 1000, max: 100 }); // Example: 100 requests per 15 mins
 */
const placeholderRateLimit = (req, res, next) => next(); // Placeholder

/**
 * Registers all API routes on the provided router.
 * @param {express.Router} router
 */
function registerRoutes(router) {
    // Add CORS headers for all routes to handle browser preflight requests
    router.use((req, res, next) => {
        res.header('Access-Control-Allow-Origin', '*');
        res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
        res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
        
        // Handle preflight requests
        if (req.method === 'OPTIONS') {
            return res.status(200).end();
        }
        next();
    });
    // MCP Tool Manifest for Claude Desktop
    router.get('/tools', async (req, res) => {
        try {
            // Extract user context
            const { userId, sessionId } = req.user || {};
            const actualSessionId = sessionId || req.session?.id;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Processing MCP tools manifest request', {
                    method: req.method,
                    path: req.path,
                    sessionId: actualSessionId,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId
                }, 'routes');
            }
            
            // Get tools dynamically from the tools service
            const tools = apiContext.toolsService.getAllTools();
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('MCP tools manifest retrieved successfully', {
                    toolCount: tools.length,
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.info('MCP tools manifest retrieved with session', {
                    sessionId: actualSessionId,
                    toolCount: tools.length,
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            res.json({ tools });
            
        } catch (error) {
            // Extract user context for error handling
            const { userId } = req.user || {};
            const actualSessionId = req.session?.id;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'routes',
                'Failed to retrieve MCP tools manifest',
                'error',
                {
                    endpoint: '/tools',
                    error: error.message,
                    stack: error.stack,
                    userId,
                    sessionId: actualSessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('MCP tools manifest retrieval failed', {
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.error('MCP tools manifest retrieval failed', {
                    sessionId: actualSessionId,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            res.status(500).json({
                error: 'TOOLS_MANIFEST_FAILED',
                error_description: 'Failed to retrieve tools manifest'
            });
        }
    });

    // Health check endpoint (on main router before v1 to avoid potential v1 middleware)
    router.get('/health', async (req, res) => {
        try {
            // Extract user context
            const { userId, sessionId } = req.user || {};
            const actualSessionId = sessionId || req.session?.id;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Processing health check request', {
                    method: req.method,
                    path: req.path,
                    sessionId: actualSessionId,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId
                }, 'routes');
            }
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Health check completed successfully', {
                    status: 'ok',
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.info('Health check completed with session', {
                    sessionId: actualSessionId,
                    status: 'ok',
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            res.json({ status: 'ok' });
            
        } catch (error) {
            // Extract user context for error handling
            const { userId } = req.user || {};
            const actualSessionId = req.session?.id;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'routes',
                'Health check failed',
                'error',
                {
                    endpoint: '/health',
                    error: error.message,
                    stack: error.stack,
                    userId,
                    sessionId: actualSessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Health check failed', {
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.error('Health check failed', {
                    sessionId: actualSessionId,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            res.status(500).json({
                error: 'HEALTH_CHECK_FAILED',
                error_description: 'Health check failed'
            });
        }
    });

    // Versioned API path
    const v1 = express.Router();
    
    // Apply routes logger middleware to all v1 routes
    v1.use(routesLogger());
    
    // Apply authentication middleware to ensure user context is available
    v1.use(requireAuth);

    // Create injected controller instances
    // TODO: Consider moving controller instantiation closer to where their routers are defined/used.
    const mailController = mailControllerFactory({ mailModule: apiContext.mailModule });
    const calendarController = calendarControllerFactory({ calendarModule: apiContext.calendarModule });
    const filesController = filesControllerFactory({ filesModule: apiContext.filesModule });
    const peopleController = peopleControllerFactory({ peopleModule: apiContext.peopleModule });
    const queryController = queryControllerFactory({
        nluAgent: apiContext.nluAgent,
        contextService: apiContext.contextService,
        errorService: apiContext.errorService
    });

    // --- Query Router --- 
    const queryRouter = express.Router();
    // Apply controller logger middleware
    queryRouter.use(controllerLogger());
    // TODO: Apply rate limiting
    queryRouter.post('/', placeholderRateLimit, queryController.handleQuery);
    v1.use('/query', queryRouter);

    // --- Mail Router --- 
    const mailRouter = express.Router();
    // Apply controller logger middleware
    mailRouter.use(controllerLogger());
    mailRouter.get('/', mailController.getMail); // Corresponds to /v1/mail
    // TODO: Apply rate limiting
    mailRouter.post('/send', placeholderRateLimit, mailController.sendMail); // Corresponds to /v1/mail/send
    mailRouter.get('/search', mailController.searchMail); // Corresponds to /v1/mail/search
    mailRouter.get('/attachments', mailController.getMailAttachments); // Corresponds to /v1/mail/attachments
    // IMPORTANT: Route order matters! Put specific routes before parametrized routes
    // Route order problem fixed: Specific routes now come before the :id pattern
    mailRouter.patch('/:id/read', placeholderRateLimit, mailController.markAsRead); // Corresponds to /v1/mail/:id/read
    // Flag/unflag email route
    mailRouter.post('/flag', placeholderRateLimit, mailController.flagMail); // Corresponds to /v1/mail/flag
    // Mail attachment routes
    mailRouter.post('/:id/attachments', placeholderRateLimit, mailController.addMailAttachment); // Corresponds to /v1/mail/:id/attachments
    mailRouter.delete('/:id/attachments/:attachmentId', mailController.removeMailAttachment); // Corresponds to /v1/mail/:id/attachments/:attachmentId
    mailRouter.get('/:id', mailController.getEmailDetails); // Corresponds to /v1/mail/:id
    v1.use('/mail', mailRouter);

    // --- Calendar Router --- 
    const calendarRouter = express.Router();
    // Apply controller logger middleware
    calendarRouter.use(controllerLogger());
    calendarRouter.get('/', calendarController.getEvents); // /v1/calendar
    // TODO: Apply rate limiting
    calendarRouter.post('/events', placeholderRateLimit, calendarController.createEvent); // /v1/calendar/events
    calendarRouter.put('/events/:id', calendarController.updateEvent); // /v1/calendar/events/:id 
    // TODO: Apply rate limiting
    calendarRouter.post('/availability', placeholderRateLimit, calendarController.getAvailability); // /v1/calendar/availability
    // TODO: Apply rate limiting
    // TODO: Apply rate limiting
    calendarRouter.post('/events/:id/accept', placeholderRateLimit, calendarController.acceptEvent);
    // TODO: Apply rate limiting
    calendarRouter.post('/events/:id/tentativelyAccept', placeholderRateLimit, calendarController.tentativelyAcceptEvent);
    // TODO: Apply rate limiting
    calendarRouter.post('/events/:id/decline', placeholderRateLimit, calendarController.declineEvent);
    // TODO: Apply rate limiting
    calendarRouter.post('/events/:id/cancel', placeholderRateLimit, calendarController.cancelEvent);
    // TODO: Apply rate limiting
    calendarRouter.post('/findMeetingTimes', placeholderRateLimit, calendarController.findMeetingTimes);
    calendarRouter.get('/rooms', calendarController.getRooms);
    calendarRouter.get('/calendars', calendarController.getCalendars);
    // TODO: Apply rate limiting
    calendarRouter.post('/events/:id/attachments', placeholderRateLimit, calendarController.addAttachment);
    calendarRouter.delete('/events/:id/attachments/:attachmentId', calendarController.removeAttachment);
    v1.use('/calendar', calendarRouter);

    // --- Files Router --- 
    const filesRouter = express.Router();
    // Apply controller logger middleware
    filesRouter.use(controllerLogger());
    filesRouter.get('/', filesController.listFiles); // /v1/files
    // TODO: Apply rate limiting
    filesRouter.post('/upload', placeholderRateLimit, filesController.uploadFile); // /v1/files/upload
    filesRouter.get('/search', filesController.searchFiles);
    filesRouter.get('/metadata', filesController.getFileMetadata);
    filesRouter.get('/content', filesController.downloadFile);
    // TODO: Apply rate limiting
    filesRouter.post('/content', placeholderRateLimit, filesController.setFileContent);
    // TODO: Apply rate limiting
    filesRouter.post('/content/update', placeholderRateLimit, filesController.updateFileContent);
    filesRouter.get('/download', filesController.downloadFile);
    // TODO: Apply rate limiting
    filesRouter.post('/share', placeholderRateLimit, filesController.createSharingLink);
    filesRouter.get('/sharing', filesController.getSharingLinks);
    // TODO: Apply rate limiting
    filesRouter.post('/sharing/remove', placeholderRateLimit, filesController.removeSharingPermission);
    v1.use('/files', filesRouter);

    // --- People Router --- 
    const peopleRouter = express.Router();
    // Apply controller logger middleware
    peopleRouter.use(controllerLogger());
    peopleRouter.get('/', peopleController.getRelevantPeople); // /v1/people
    peopleRouter.get('/find', peopleController.findPeople);
    peopleRouter.get('/:id', peopleController.getPersonById); // /v1/people/:id
    v1.use('/people', peopleRouter);

    // --- Log Router --- (No Auth required for logs)
    const logRouter = express.Router();
    // Apply controller logger middleware
    logRouter.use(controllerLogger());
    // TODO: Apply rate limiting
    logRouter.post('/', placeholderRateLimit, logController.addLogEntry); // /v1/logs
    logRouter.get('/', logController.getLogEntries); // /v1/logs
    // Convenience endpoints for specific log categories
    logRouter.get('/calendar', async (req, res) => {
        try {
            // Extract user context
            const { userId, sessionId } = req.user || {};
            const actualSessionId = sessionId || req.session?.id;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Processing calendar logs filter request', {
                    method: req.method,
                    path: req.path,
                    sessionId: actualSessionId,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId,
                    category: 'calendar'
                }, 'routes');
            }
            
            // Pre-filter for calendar logs
            req.query.category = 'calendar';
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Calendar logs filter applied successfully', {
                    category: 'calendar',
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.info('Calendar logs filter applied with session', {
                    sessionId: actualSessionId,
                    category: 'calendar',
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            return logController.getLogEntries(req, res);
            
        } catch (error) {
            // Extract user context for error handling
            const { userId } = req.user || {};
            const actualSessionId = req.session?.id;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'routes',
                'Failed to apply calendar logs filter',
                'error',
                {
                    endpoint: '/v1/logs/calendar',
                    error: error.message,
                    stack: error.stack,
                    userId,
                    sessionId: actualSessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Calendar logs filter failed', {
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes', null, userId);
            } else if (actualSessionId) {
                MonitoringService.error('Calendar logs filter failed', {
                    sessionId: actualSessionId,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'routes');
            }
            
            res.status(500).json({
                error: 'CALENDAR_LOGS_FILTER_FAILED',
                error_description: 'Failed to filter calendar logs'
            });
        }
    }); // /v1/logs/calendar
    // TODO: Apply rate limiting
    logRouter.post('/clear', placeholderRateLimit, logController.clearLogEntries); // /v1/logs/clear
    // RESTful DELETE endpoint for clearing logs
    logRouter.delete('/', placeholderRateLimit, logController.clearLogEntries); // /v1/logs
    v1.use('/logs', logRouter); // Mounted at /v1/logs

    // --- Auth Router ---
    const authRouter = express.Router();
    authRouter.use(controllerLogger());
    
    // Web-based authentication endpoints
    authRouter.get('/status', authController.getAuthStatus);
    authRouter.get('/login', authController.login);
    authRouter.get('/callback', authController.handleCallback);
    authRouter.post('/logout', authController.logout);
    
    // Device authentication endpoints (don't require authentication as they're part of the auth flow)
    authRouter.post('/device/register', deviceAuthController.registerDevice);
    authRouter.post('/device/authorize', deviceAuthController.authorizeDevice);
    authRouter.post('/device/token', deviceAuthController.pollForToken);
    authRouter.post('/device/refresh', deviceAuthController.refreshToken);
    
    // MCP token generation endpoint - requires authentication
    authRouter.post('/generate-mcp-token', requireAuth, deviceAuthController.generateMcpToken);
    
    // OAuth 2.0 discovery endpoint
    authRouter.get('/.well-known/oauth-protected-resource', deviceAuthController.getResourceServerInfo);
    
    // Register auth router at both /auth and /api/auth for compatibility
    router.use('/auth', authRouter);
    router.use('/api/auth', authRouter);

    // --- Adapter Router --- 
    const adapterRouter = express.Router();
    // Apply controller logger middleware
    adapterRouter.use(controllerLogger());
    adapterRouter.get('/download/:deviceId', adapterController.downloadAdapter); // /adapter/download/:deviceId
    adapterRouter.get('/package/:deviceId', adapterController.downloadPackageJson); // /adapter/package/:deviceId
    adapterRouter.get('/setup/:deviceId', adapterController.downloadSetupInstructions); // /adapter/setup/:deviceId
    router.use('/adapter', adapterRouter);

    // Serve MCP adapter directly at /mcp-adapter.cjs for easy distribution
    const path = require('path');
    const fs = require('fs');
    router.get('/mcp-adapter.cjs', (req, res) => {
        try {
            const adapterPath = path.join(__dirname, '../../mcp-adapter.cjs');
            res.setHeader('Content-Type', 'application/javascript');
            res.setHeader('Content-Disposition', 'attachment; filename="mcp-adapter.cjs"');
            res.sendFile(adapterPath);
        } catch (error) {
            res.status(500).json({ 
                error: 'ADAPTER_DOWNLOAD_FAILED', 
                error_description: 'Failed to serve MCP adapter' 
            });
        }
    });

    // Debug routes (development only)
    if (process.env.NODE_ENV === 'development') {
        router.get('/api/v1/debug/graph-token', requireAuth, async (req, res) => {
            try {
                // Extract user context
                const { userId, sessionId } = req.user || {};
                const actualSessionId = sessionId || req.session?.id;
                
                // Pattern 1: Development Debug Logs
                MonitoringService.debug('Processing debug graph token request', {
                    method: req.method,
                    path: req.path,
                    sessionId: actualSessionId,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId
                }, 'routes');
                
                const graphClientFactory = require('../graph/graph-client-factory.cjs');
                const client = await graphClientFactory.createClient(req);
                
                // Get the access token from the client
                const authProvider = client.authProvider || client._authProvider;
                if (authProvider && authProvider.getAccessToken) {
                    const accessToken = await authProvider.getAccessToken();
                    
                    // Pattern 2: User Activity Logs
                    if (userId) {
                        MonitoringService.info('Debug graph token retrieved successfully', {
                            hasToken: !!accessToken,
                            tokenLength: accessToken ? accessToken.length : 0,
                            timestamp: new Date().toISOString()
                        }, 'routes', null, userId);
                    } else if (actualSessionId) {
                        MonitoringService.info('Debug graph token retrieved with session', {
                            sessionId: actualSessionId,
                            hasToken: !!accessToken,
                            tokenLength: accessToken ? accessToken.length : 0,
                            timestamp: new Date().toISOString()
                        }, 'routes');
                    }
                    
                    res.json({ 
                        accessToken: accessToken,
                        hasToken: !!accessToken,
                        tokenLength: accessToken ? accessToken.length : 0
                    });
                } else {
                    // Pattern 4: User Error Tracking for auth provider issue
                    if (userId) {
                        MonitoringService.error('Debug graph token - auth provider not accessible', {
                            error: 'Could not access auth provider',
                            timestamp: new Date().toISOString()
                        }, 'routes', null, userId);
                    } else if (actualSessionId) {
                        MonitoringService.error('Debug graph token - auth provider not accessible', {
                            sessionId: actualSessionId,
                            error: 'Could not access auth provider',
                            timestamp: new Date().toISOString()
                        }, 'routes');
                    }
                    
                    res.json({ 
                        error: 'Could not access auth provider',
                        hasToken: false
                    });
                }
            } catch (error) {
                // Extract user context for error handling
                const { userId } = req.user || {};
                const actualSessionId = req.session?.id;
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'routes',
                    'Failed to get debug graph token',
                    'error',
                    {
                        endpoint: '/api/v1/debug/graph-token',
                        error: error.message,
                        stack: error.stack,
                        userId,
                        sessionId: actualSessionId,
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Debug graph token retrieval failed', {
                        error: error.message,
                        timestamp: new Date().toISOString()
                    }, 'routes', null, userId);
                } else if (actualSessionId) {
                    MonitoringService.error('Debug graph token retrieval failed', {
                        sessionId: actualSessionId,
                        error: error.message,
                        timestamp: new Date().toISOString()
                    }, 'routes');
                }
                
                res.status(500).json({ 
                    error: 'DEBUG_GRAPH_TOKEN_FAILED',
                    error_description: 'Failed to get access token'
                });
            }
        });
    }

    // Mount v1 under /v1 path
    router.use('/v1', v1);
}

module.exports = { registerRoutes };
