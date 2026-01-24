/**
 * @fileoverview Registers all API routes and controllers for MCP.
 * Handles versioning, middleware, and route registration.
 */

const express = require('express');
const queryControllerFactory = require('./controllers/query-controller.cjs');
const mailControllerFactory = require('./controllers/mail-controller.cjs');
const calendarControllerFactory = require('./controllers/calendar-controller.cjs');
const filesControllerFactory = require('./controllers/files-controller.cjs');
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
const graphTokenExchangeController = require('./controllers/graph-token-exchange-controller.cjs');
const adapterController = require('./controllers/adapter-controller.cjs');
const mcpTransportController = require('./controllers/mcp-transport-controller.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const { requireAuth } = require('./middleware/auth-middleware.cjs');
const { routesLogger, controllerLogger } = require('./middleware/request-logger.cjs');
const apiContext = require('./api-context.cjs');
const statusRouter = require('./status.cjs');

// SEC-3: Rate limiting configuration
const RATE_LIMIT_WINDOW_MS = parseInt(process.env.RATE_LIMIT_WINDOW_MS || '900000', 10); // 15 minutes default
const RATE_LIMIT_MAX = parseInt(process.env.RATE_LIMIT_MAX || '100', 10); // 100 requests per window default
const RATE_LIMIT_AUTH_MAX = parseInt(process.env.RATE_LIMIT_AUTH_MAX || '20', 10); // 20 auth attempts per window

// Simple in-memory rate limiter (use Redis in production cluster)
const rateLimitStore = new Map();
function createRateLimiter(maxRequests, windowMs = RATE_LIMIT_WINDOW_MS) {
    return (req, res, next) => {
        const key = req.ip || req.connection.remoteAddress || 'unknown';
        const now = Date.now();

        let record = rateLimitStore.get(key);
        if (!record || now - record.start > windowMs) {
            record = { count: 1, start: now };
            rateLimitStore.set(key, record);
        } else {
            record.count++;
        }

        // Set rate limit headers
        res.set('X-RateLimit-Limit', maxRequests.toString());
        res.set('X-RateLimit-Remaining', Math.max(0, maxRequests - record.count).toString());
        res.set('X-RateLimit-Reset', Math.ceil((record.start + windowMs) / 1000).toString());

        if (record.count > maxRequests) {
            MonitoringService.warn('Rate limit exceeded', {
                ip: key,
                path: req.path,
                count: record.count,
                limit: maxRequests,
                timestamp: new Date().toISOString()
            }, 'security');
            return res.status(429).json({
                error: 'TOO_MANY_REQUESTS',
                error_description: 'Rate limit exceeded. Please try again later.',
                retryAfter: Math.ceil((record.start + windowMs - now) / 1000)
            });
        }
        next();
    };
}

// Cleanup old rate limit entries every 5 minutes
setInterval(() => {
    const now = Date.now();
    for (const [key, record] of rateLimitStore.entries()) {
        if (now - record.start > RATE_LIMIT_WINDOW_MS * 2) {
            rateLimitStore.delete(key);
        }
    }
}, 300000);

const apiRateLimit = createRateLimiter(RATE_LIMIT_MAX);
const authRateLimit = createRateLimiter(RATE_LIMIT_AUTH_MAX);
// Placeholder rate limit - used for endpoints that need rate limiting configured
const placeholderRateLimit = apiRateLimit;

// SEC-2: CORS allowlist configuration
const CORS_ALLOWED_ORIGINS = (process.env.CORS_ALLOWED_ORIGINS || '')
    .split(',')
    .map(o => o.trim())
    .filter(Boolean);
const CORS_ALLOW_ALL = process.env.NODE_ENV === 'development' && CORS_ALLOWED_ORIGINS.length === 0;

if (CORS_ALLOW_ALL) {
    console.warn('[SECURITY WARNING] CORS allows all origins in development. Set CORS_ALLOWED_ORIGINS for production.');
}

/**
 * Registers all API routes on the provided router.
 * @param {express.Router} router
 */
function registerRoutes(router) {
    // SEC-2: CORS with allowlist (production) or open (development only)
    router.use((req, res, next) => {
        const origin = req.get('Origin');

        // Check if origin is allowed
        let allowOrigin = null;
        if (CORS_ALLOW_ALL) {
            allowOrigin = '*';
        } else if (origin && CORS_ALLOWED_ORIGINS.includes(origin)) {
            allowOrigin = origin;
        } else if (!origin) {
            // No Origin header = same-origin request or non-browser client
            allowOrigin = null; // Don't set CORS headers for same-origin
        }

        if (allowOrigin) {
            res.header('Access-Control-Allow-Origin', allowOrigin);
            res.header('Access-Control-Allow-Credentials', 'true');
        }
        res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
        res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');

        // Handle preflight requests
        if (req.method === 'OPTIONS') {
            if (!allowOrigin && origin) {
                // Origin not in allowlist - deny preflight
                MonitoringService.warn('CORS preflight denied', {
                    origin,
                    path: req.path,
                    timestamp: new Date().toISOString()
                }, 'security');
                return res.status(403).json({ error: 'CORS_ORIGIN_NOT_ALLOWED' });
            }
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
                // Mail tools
                'Mail.Read': ['getInbox', 'readMail', 'readMailDetails', 'search', 'getEmailDetails', 'getMailAttachments'],
                'Mail.ReadWrite': ['getInbox', 'readMail', 'readMailDetails', 'search', 'getEmailDetails', 'getMailAttachments', 'markAsRead', 'markEmailRead', 'flagMail', 'flagEmail', 'addMailAttachment', 'removeMailAttachment'],
                'Mail.Send': ['sendEmail', 'sendMail', 'replyToMail', 'replyToEmail'],
                // Calendar tools
                'Calendars.Read': ['getEvents', 'getAvailability', 'getCalendars', 'getRooms', 'search'],
                'Calendars.ReadWrite': ['getEvents', 'getAvailability', 'getCalendars', 'getRooms', 'search', 'createEvent', 'updateEvent', 'cancelEvent', 'acceptEvent', 'tentativelyAcceptEvent', 'declineEvent', 'findMeetingTimes', 'addAttachment', 'removeAttachment'],
                // Files tools
                'Files.Read': ['listFiles', 'search', 'downloadFile', 'getFileMetadata', 'getFileContent', 'getSharingLinks'],
                'Files.ReadWrite': ['listFiles', 'search', 'downloadFile', 'getFileMetadata', 'getFileContent', 'getSharingLinks', 'uploadFile', 'createSharingLink', 'setFileContent', 'updateFileContent', 'removeSharingPermission'],
                // People tools
                'People.Read': ['findPeople', 'getRelevantPeople', 'getPersonById', 'search'],
                // Teams/Chat tools
                'Chat.Read': ['listChats', 'getChatMessages', 'search'],
                'Chat.ReadWrite': ['listChats', 'getChatMessages', 'sendChatMessage', 'search'],
                'OnlineMeetings.Read': ['listOnlineMeetings', 'getOnlineMeeting', 'getMeetingByJoinUrl'],
                'OnlineMeetings.ReadWrite': ['listOnlineMeetings', 'getOnlineMeeting', 'getMeetingByJoinUrl', 'createOnlineMeeting'],
                'OnlineMeetingTranscript.Read.All': ['getMeetingTranscripts', 'getMeetingTranscriptContent'],
                'Team.ReadBasic.All': ['listJoinedTeams'],
                'Channel.ReadBasic.All': ['listTeamChannels', 'getChannelMessages'],
                'ChannelMessage.Send': ['sendChannelMessage', 'replyToMessage'],
                // Tasks tools
                'Tasks.Read': ['listTaskLists', 'getTaskList', 'listTasks', 'getTask'],
                'Tasks.ReadWrite': ['listTaskLists', 'getTaskList', 'listTasks', 'getTask', 'createTaskList', 'updateTaskList', 'deleteTaskList', 'createTask', 'updateTask', 'deleteTask', 'completeTask'],
                // Contacts tools
                'Contacts.Read': ['listContacts', 'getContact', 'searchContacts'],
                'Contacts.ReadWrite': ['listContacts', 'getContact', 'searchContacts', 'createContact', 'updateContact', 'deleteContact'],
                // Groups tools
                'Group.Read.All': ['listGroups', 'getGroup', 'listGroupMembers', 'listMyGroups'],
                // User tools
                'User.Read': ['getProfile'],
                // Query tool (available with any permission)
                'User.ReadBasic.All': ['query']
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
    // Reply to email route
    mailRouter.post('/:id/reply', placeholderRateLimit, mailController.replyToMail); // Corresponds to /v1/mail/:id/reply
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
    // SEC-3: Apply rate limiting
    logRouter.post('/clear', apiRateLimit, logController.clearLogEntries); // /v1/logs/clear
    // RESTful DELETE endpoint for clearing logs
    logRouter.delete('/', apiRateLimit, logController.clearLogEntries); // /v1/logs
    v1.use('/logs', logRouter); // Mounted at /v1/logs

    // --- Auth Router ---
    const authRouter = express.Router();
    authRouter.use(controllerLogger());

    // Web-based authentication endpoints
    authRouter.get('/status', apiRateLimit, authController.getAuthStatus);
    authRouter.get('/login', authRateLimit, authController.login);
    authRouter.get('/callback', authRateLimit, authController.handleCallback);
    authRouter.post('/logout', apiRateLimit, authController.logout);

    // Device authentication endpoints - SEC-3: Rate limited to prevent brute-force
    authRouter.post('/device/register', authRateLimit, deviceAuthController.registerDevice);
    authRouter.post('/device/authorize', authRateLimit, deviceAuthController.authorizeDevice);
    authRouter.post('/device/token', authRateLimit, deviceAuthController.pollForToken);
    authRouter.post('/device/refresh', authRateLimit, deviceAuthController.refreshToken);

    // MCP token generation endpoint - requires authentication
    authRouter.post('/generate-mcp-token', authRateLimit, requireAuth, deviceAuthController.generateMcpToken);

    // External token login endpoint - SEC-3: Rate limited to prevent brute-force
    authRouter.post('/external-token/login', authRateLimit, externalTokenController.loginWithToken);

    // Graph token exchange endpoint - exchanges MS Graph tokens for MCP JWTs
    // Used by Synthetic Employees and other ROPC clients
    authRouter.post('/graph-token-exchange', authRateLimit, graphTokenExchangeController.exchange);

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
    mcpRouter.use(apiRateLimit);

    // SEC-4: Token in query parameter for SSE connections
    // SECURITY NOTE: Browser's EventSource API doesn't support Authorization headers,
    // so SSE connections must pass tokens via query params. This is a known security trade-off:
    // - Tokens may appear in server logs and browser history
    // - Only allowed for /sse endpoints, logged for audit trail
    // - Use short-lived tokens and HTTPS in production
    mcpRouter.use((req, res, next) => {
        if (req.query.token && !req.headers.authorization) {
            // Only allow for SSE endpoints
            if (req.path === '/sse' || req.path.startsWith('/sse')) {
                MonitoringService.info('Token passed via query parameter (SSE)', {
                    path: req.path,
                    ip: req.ip || req.connection?.remoteAddress,
                    userAgent: req.get('User-Agent')?.substring(0, 50),
                    timestamp: new Date().toISOString()
                }, 'security');
                req.headers.authorization = `Bearer ${req.query.token}`;
            } else {
                // Reject query param tokens for non-SSE endpoints
                MonitoringService.warn('Token in query param rejected (non-SSE)', {
                    path: req.path,
                    ip: req.ip || req.connection?.remoteAddress,
                    timestamp: new Date().toISOString()
                }, 'security');
                return res.status(400).json({
                    error: 'INVALID_AUTH_METHOD',
                    error_description: 'Use Authorization header for this endpoint'
                });
            }
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
