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
const searchControllerFactory = require('./controllers/search-controller.cjs');
const teamsControllerFactory = require('./controllers/teams-controller.cjs');
const todoControllerFactory = require('./controllers/todo-controller.cjs');
const contactsControllerFactory = require('./controllers/contacts-controller.cjs');
const groupsControllerFactory = require('./controllers/groups-controller.cjs');
const logController = require('./controllers/log-controller.cjs');
const authController = require('./controllers/auth-controller.cjs');
const deviceAuthController = require('./controllers/device-auth-controller.cjs');
const externalTokenController = require('./controllers/external-token-controller.cjs');
const adapterController = require('./controllers/adapter-controller.cjs');
const mcpTransportController = require('./controllers/mcp-transport-controller.cjs');
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
    const searchController = searchControllerFactory({ searchModule: apiContext.searchModule });
    const queryController = queryControllerFactory({
        nluAgent: apiContext.nluAgent,
        contextService: apiContext.contextService,
        errorService: apiContext.errorService
    });

    // --- Permissions Endpoint ---
    // Returns available scopes and tool capabilities based on user's Microsoft Graph token
    v1.get('/permissions', async (req, res) => {
        try {
            const msalService = require('../auth/msal-service.cjs');
            const externalTokenValidator = require('../auth/external-token-validator.cjs');

            // Get the Microsoft Graph token
            const accessToken = await msalService.getAccessToken(req);
            if (!accessToken) {
                return res.status(401).json({ error: 'No access token available' });
            }

            // Decode token and extract scopes
            const decoded = externalTokenValidator.decodeToken(accessToken);
            const scopes = decoded.payload.scp ? decoded.payload.scp.split(' ').sort() : [];

            // Map scopes to tool capabilities
            // Note: 'search' is the unified search tool - available with any read permission
            const scopeToTools = {
                'Mail.Read': ['getInbox', 'search', 'getEmailDetails', 'getMailAttachments'],
                'Mail.ReadWrite': ['getInbox', 'search', 'getEmailDetails', 'getMailAttachments', 'markAsRead', 'flagEmail'],
                'Mail.Send': ['sendEmail'],
                'Calendars.Read': ['getEvents', 'getAvailability', 'search'],
                'Calendars.ReadWrite': ['getEvents', 'getAvailability', 'search', 'createEvent', 'updateEvent', 'cancelEvent', 'acceptEvent', 'declineEvent', 'findMeetingTimes'],
                'Files.Read': ['listFiles', 'search', 'downloadFile', 'getFileMetadata'],
                'Files.ReadWrite': ['listFiles', 'search', 'downloadFile', 'getFileMetadata', 'uploadFile', 'createSharingLink'],
                'People.Read': ['findPeople', 'getRelevantPeople', 'search'],
                'Chat.Read': ['listChats', 'getChatMessages', 'search'],
                'Chat.ReadWrite': ['listChats', 'getChatMessages', 'sendChatMessage', 'search'],
                'OnlineMeetings.Read': ['listOnlineMeetings', 'getOnlineMeeting'],
                'OnlineMeetings.ReadWrite': ['listOnlineMeetings', 'getOnlineMeeting', 'createOnlineMeeting'],
                'OnlineMeetingTranscript.Read.All': ['getMeetingTranscripts', 'getMeetingTranscriptContent'],
                'Team.ReadBasic.All': ['listJoinedTeams'],
                'Channel.ReadBasic.All': ['listTeamChannels', 'getChannelMessages'],
                'ChannelMessage.Send': ['sendChannelMessage'],
                'Tasks.Read': ['listTaskLists', 'getTaskList', 'listTasks', 'getTask'],
                'Tasks.ReadWrite': ['listTaskLists', 'getTaskList', 'listTasks', 'getTask', 'createTaskList', 'updateTaskList', 'deleteTaskList', 'createTask', 'updateTask', 'deleteTask', 'completeTask'],
                'Contacts.Read': ['listContacts', 'getContact', 'searchContacts'],
                'Contacts.ReadWrite': ['listContacts', 'getContact', 'searchContacts', 'createContact', 'updateContact', 'deleteContact'],
                'Group.Read.All': ['listGroups', 'getGroup', 'listGroupMembers', 'listMyGroups'],
                'User.Read': ['getProfile']
            };

            // Build available tools from scopes
            const availableTools = new Set();
            for (const scope of scopes) {
                const tools = scopeToTools[scope] || [];
                tools.forEach(t => availableTools.add(t));
            }

            res.json({
                scopes,
                availableTools: Array.from(availableTools).sort(),
                scopeCount: scopes.length,
                toolCount: availableTools.size
            });
        } catch (error) {
            MonitoringService.error('Permissions check failed', { error: error.message }, 'routes');
            res.status(500).json({ error: 'Failed to check permissions', message: error.message });
        }
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

    // --- Search Router ---
    const searchRouter = express.Router();
    // Apply controller logger middleware
    searchRouter.use(controllerLogger());
    // Unified search supports both GET and POST
    searchRouter.get('/', searchController.search); // /v1/search?query=...
    // TODO: Apply rate limiting
    searchRouter.post('/', placeholderRateLimit, searchController.search); // /v1/search
    v1.use('/search', searchRouter);

    // --- Teams Router ---
    const teamsController = teamsControllerFactory({ teamsModule: apiContext.teamsModule });
    const teamsRouter = express.Router();
    // Apply controller logger middleware
    teamsRouter.use(controllerLogger());
    // Chat routes
    teamsRouter.get('/chats', teamsController.listChats); // /v1/teams/chats
    teamsRouter.get('/chats/:chatId/messages', teamsController.getChatMessages); // /v1/teams/chats/:chatId/messages
    teamsRouter.post('/chats/:chatId/messages', placeholderRateLimit, teamsController.sendChatMessage); // /v1/teams/chats/:chatId/messages
    // Teams and channel routes
    teamsRouter.get('/', teamsController.listJoinedTeams); // /v1/teams
    teamsRouter.get('/:teamId/channels', teamsController.listTeamChannels); // /v1/teams/:teamId/channels
    teamsRouter.get('/:teamId/channels/:channelId/messages', teamsController.getChannelMessages); // /v1/teams/:teamId/channels/:channelId/messages
    teamsRouter.post('/:teamId/channels/:channelId/messages', placeholderRateLimit, teamsController.sendChannelMessage); // /v1/teams/:teamId/channels/:channelId/messages
    teamsRouter.post('/:teamId/channels/:channelId/messages/:messageId/replies', placeholderRateLimit, teamsController.replyToMessage); // /v1/teams/:teamId/channels/:channelId/messages/:messageId/replies
    // Meeting routes
    teamsRouter.get('/meetings', teamsController.listOnlineMeetings); // /v1/teams/meetings
    teamsRouter.post('/meetings', placeholderRateLimit, teamsController.createOnlineMeeting); // /v1/teams/meetings
    teamsRouter.get('/meetings/findByJoinUrl', teamsController.getMeetingByJoinUrl); // /v1/teams/meetings/findByJoinUrl
    teamsRouter.get('/meetings/:meetingId', teamsController.getOnlineMeeting); // /v1/teams/meetings/:meetingId
    // Transcript routes
    teamsRouter.get('/meetings/:meetingId/transcripts', teamsController.getMeetingTranscripts); // /v1/teams/meetings/:meetingId/transcripts
    teamsRouter.get('/meetings/:meetingId/transcripts/:transcriptId', teamsController.getMeetingTranscriptContent); // /v1/teams/meetings/:meetingId/transcripts/:transcriptId
    v1.use('/teams', teamsRouter);

    // --- To-Do Router ---
    const todoController = todoControllerFactory({ todoModule: apiContext.todoModule });
    const todoRouter = express.Router();
    // Apply controller logger middleware
    todoRouter.use(controllerLogger());
    // Task list routes
    todoRouter.get('/lists', todoController.listTaskLists); // /v1/todo/lists
    todoRouter.post('/lists', placeholderRateLimit, todoController.createTaskList); // /v1/todo/lists
    todoRouter.get('/lists/:listId', todoController.getTaskList); // /v1/todo/lists/:listId
    todoRouter.patch('/lists/:listId', placeholderRateLimit, todoController.updateTaskList); // /v1/todo/lists/:listId
    todoRouter.delete('/lists/:listId', todoController.deleteTaskList); // /v1/todo/lists/:listId
    // Task routes
    todoRouter.get('/lists/:listId/tasks', todoController.listTasks); // /v1/todo/lists/:listId/tasks
    todoRouter.post('/lists/:listId/tasks', placeholderRateLimit, todoController.createTask); // /v1/todo/lists/:listId/tasks
    todoRouter.get('/lists/:listId/tasks/:taskId', todoController.getTask); // /v1/todo/lists/:listId/tasks/:taskId
    todoRouter.patch('/lists/:listId/tasks/:taskId', placeholderRateLimit, todoController.updateTask); // /v1/todo/lists/:listId/tasks/:taskId
    todoRouter.delete('/lists/:listId/tasks/:taskId', todoController.deleteTask); // /v1/todo/lists/:listId/tasks/:taskId
    todoRouter.post('/lists/:listId/tasks/:taskId/complete', placeholderRateLimit, todoController.completeTask); // /v1/todo/lists/:listId/tasks/:taskId/complete
    v1.use('/todo', todoRouter);

    // --- Contacts Router ---
    const contactsController = contactsControllerFactory({ contactsModule: apiContext.contactsModule });
    const contactsRouter = express.Router();
    // Apply controller logger middleware
    contactsRouter.use(controllerLogger());
    contactsRouter.get('/', contactsController.listContacts); // /v1/contacts
    contactsRouter.get('/search', contactsController.searchContacts); // /v1/contacts/search
    contactsRouter.post('/', placeholderRateLimit, contactsController.createContact); // /v1/contacts
    contactsRouter.get('/:contactId', contactsController.getContact); // /v1/contacts/:contactId
    contactsRouter.patch('/:contactId', placeholderRateLimit, contactsController.updateContact); // /v1/contacts/:contactId
    contactsRouter.delete('/:contactId', contactsController.deleteContact); // /v1/contacts/:contactId
    v1.use('/contacts', contactsRouter);

    // --- Groups Router ---
    const groupsController = groupsControllerFactory({ groupsModule: apiContext.groupsModule });
    const groupsRouter = express.Router();
    // Apply controller logger middleware
    groupsRouter.use(controllerLogger());
    groupsRouter.get('/', groupsController.listGroups); // /v1/groups
    groupsRouter.get('/my', groupsController.listMyGroups); // /v1/groups/my
    groupsRouter.get('/:groupId', groupsController.getGroup); // /v1/groups/:groupId
    groupsRouter.get('/:groupId/members', groupsController.listGroupMembers); // /v1/groups/:groupId/members
    v1.use('/groups', groupsRouter);

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

    // External token login endpoint - NO AUTH REQUIRED (this IS the login)
    authRouter.post('/external-token/login', externalTokenController.loginWithToken);

    // External token management endpoints - requires authentication
    authRouter.post('/external-token', requireAuth, externalTokenController.inject);
    authRouter.get('/external-token/status', requireAuth, externalTokenController.status);
    authRouter.delete('/external-token', requireAuth, externalTokenController.clear);
    authRouter.post('/external-token/switch', requireAuth, externalTokenController.switchSource);

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

    // --- MCP Transport Router (SSE transport for direct Claude Desktop connection) ---
    const mcpRouter = express.Router();
    mcpRouter.use(controllerLogger());

    // Middleware to extract token from query parameter (for SSE connections)
    // This allows: /api/mcp/sse?token=xxx
    mcpRouter.use((req, res, next) => {
        if (req.query.token && !req.headers.authorization) {
            req.headers.authorization = `Bearer ${req.query.token}`;
        }
        next();
    });

    // Info endpoint - no auth required
    mcpRouter.get('/info', mcpTransportController.getInfo);

    // SSE endpoint - requires authentication (via Bearer token, query param, or session)
    // GET: SSE connection for streaming
    // POST: Streamable HTTP transport (mcp-remote uses this first)
    mcpRouter.get('/sse', requireAuth, mcpTransportController.sseConnect);
    mcpRouter.post('/sse', requireAuth, mcpTransportController.handleMessage);

    // Message endpoint - requires authentication
    mcpRouter.post('/message', requireAuth, mcpTransportController.handleMessage);

    // Simple HTTP endpoint - requires authentication
    mcpRouter.post('/', requireAuth, mcpTransportController.handleSimpleMessage);

    router.use('/mcp', mcpRouter);

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
