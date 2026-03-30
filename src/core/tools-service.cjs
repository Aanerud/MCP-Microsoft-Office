/**
 * @fileoverview ToolsService - Aggregates and manages MCP tools from modules.
 * Follows MCP modular, testable, and consistent API contract rules.
 * Handles tool definition, mapping, and parameter transformation.
 */

const ErrorService = require('./error-service.cjs');
const MonitoringService = require('./monitoring-service.cjs');

// Log service initialization
MonitoringService.info('Tools service factory initialized', {
    serviceName: 'tools-service',
    timestamp: new Date().toISOString()
}, 'tools');

/**
 * Creates a tools service with the module registry.
 * @param {object} deps - Service dependencies
 * @param {object} deps.moduleRegistry - The module registry instance
 * @param {object} [deps.logger=console] - Logger instance
 * @param {object} [deps.schemaValidator] - Schema validation service (optional)
 * @param {string} [userId] - User ID for multi-user context
 * @param {string} [sessionId] - Session ID for context
 * @returns {object} Tools service methods
 */
function createToolsService({ moduleRegistry, logger = console, schemaValidator = null }, userId = null, sessionId = null) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Creating tools service instance', {
            hasModuleRegistry: !!moduleRegistry,
            hasLogger: !!logger,
            hasSchemaValidator: !!schemaValidator,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        }, 'tools', null, userId);
    }
    
    try {
        if (!moduleRegistry) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'tools',
                'Module registry is required for ToolsService',
                'critical',
                {
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Tools service creation failed - missing module registry', {
                    error: 'Module registry is required',
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Tools service creation failed - missing module registry', {
                    sessionId,
                    error: 'Module registry is required',
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            throw mcpError;
        }
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Tools service instance created successfully', {
                hasModuleRegistry: !!moduleRegistry,
                hasLogger: !!logger,
                hasSchemaValidator: !!schemaValidator,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Tools service instance created with session', {
                sessionId,
                hasModuleRegistry: !!moduleRegistry,
                hasLogger: !!logger,
                hasSchemaValidator: !!schemaValidator,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'tools');
        }
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'tools',
            `Failed to create tools service: ${error.message}`,
            'error',
            {
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
            MonitoringService.error('Tools service creation failed', {
                error: error.message,
                duration: executionTime,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Tools service creation failed', {
                sessionId,
                error: error.message,
                duration: executionTime,
                timestamp: new Date().toISOString()
            }, 'tools');
        }
        
        MonitoringService.trackMetric('tools_service_creation_failure', executionTime, {
            errorType: error.code || 'unknown',
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        }, userId);
        
        throw error;
    }

    // Internal state for caching (will be used later)
    let cachedTools = null;

    // Comprehensive tool alias map for consistent module and method routing
    const toolAliases = {
        // Unified search tool (replaces searchMail, searchFiles)
        search: { moduleName: 'search', methodName: 'search' },

        // Mail module tools
        getMail: { moduleName: 'mail', methodName: 'getInbox' },
        readMail: { moduleName: 'mail', methodName: 'getInbox' },
        sendMail: { moduleName: 'mail', methodName: 'sendEmail' },
        flagMail: { moduleName: 'mail', methodName: 'flagEmail' },
        getMailDetails: { moduleName: 'mail', methodName: 'getEmailDetails' },
        markMailRead: { moduleName: 'mail', methodName: 'markAsRead' },
        markEmailRead: { moduleName: 'mail', methodName: 'markAsRead' },
        addMailAttachment: { moduleName: 'mail', methodName: 'addMailAttachment' },
        removeMailAttachment: { moduleName: 'mail', methodName: 'removeMailAttachment' },
        
        // Calendar module tools
        getCalendar: { moduleName: 'calendar', methodName: 'getEvents' },
        getEvents: { moduleName: 'calendar', methodName: 'getEvents' },
        createEvent: { moduleName: 'calendar', methodName: 'create' },
        updateEvent: { moduleName: 'calendar', methodName: 'update' },
        cancelEvent: { moduleName: 'calendar', methodName: 'cancelEvent' },
        acceptEvent: { moduleName: 'calendar', methodName: 'acceptEvent' },
        tentativelyAcceptEvent: { moduleName: 'calendar', methodName: 'tentativelyAcceptEvent' },
        declineEvent: { moduleName: 'calendar', methodName: 'declineEvent' },
        getAvailability: { moduleName: 'calendar', methodName: 'getAvailability' },
        findMeetingTimes: { moduleName: 'calendar', methodName: 'findMeetingTimes' },
        addAttachment: { moduleName: 'calendar', methodName: 'addAttachment' },
        removeAttachment: { moduleName: 'calendar', methodName: 'removeAttachment' },
        
        // Files module tools
        listFiles: { moduleName: 'files', methodName: 'listFiles' },
        downloadFile: { moduleName: 'files', methodName: 'downloadFile' },
        uploadFile: { moduleName: 'files', methodName: 'uploadFile' },
        getFileMetadata: { moduleName: 'files', methodName: 'getFileMetadata' },
        getFileContent: { moduleName: 'files', methodName: 'getFileContent' },
        setFileContent: { moduleName: 'files', methodName: 'setFileContent' },
        updateFileContent: { moduleName: 'files', methodName: 'updateFileContent' },
        createSharingLink: { moduleName: 'files', methodName: 'createSharingLink' },
        getSharingLinks: { moduleName: 'files', methodName: 'getSharingLinks' },
        removeSharingPermission: { moduleName: 'files', methodName: 'removeSharingPermission' },
        
        // People module tools
        findPeople: { moduleName: 'people', methodName: 'find' },
        getRelevantPeople: { moduleName: 'people', methodName: 'getRelevantPeople' },
        getPersonById: { moduleName: 'people', methodName: 'getPersonById' },

        // Teams module tools - Chat operations
        listChats: { moduleName: 'teams', methodName: 'listChats' },
        getChats: { moduleName: 'teams', methodName: 'listChats' },
        createChat: { moduleName: 'teams', methodName: 'createChat' },
        getChatMessages: { moduleName: 'teams', methodName: 'getChatMessages' },
        sendChatMessage: { moduleName: 'teams', methodName: 'sendChatMessage' },
        // Teams module tools - Team & channel operations
        listJoinedTeams: { moduleName: 'teams', methodName: 'listJoinedTeams' },
        getTeams: { moduleName: 'teams', methodName: 'listJoinedTeams' },
        listTeamChannels: { moduleName: 'teams', methodName: 'listTeamChannels' },
        getTeamChannels: { moduleName: 'teams', methodName: 'listTeamChannels' },
        getChannelMessages: { moduleName: 'teams', methodName: 'getChannelMessages' },
        sendChannelMessage: { moduleName: 'teams', methodName: 'sendChannelMessage' },
        replyToMessage: { moduleName: 'teams', methodName: 'replyToMessage' },
        // Teams module tools - Channel management operations
        createTeamChannel: { moduleName: 'teams', methodName: 'createTeamChannel' },
        addChannelMember: { moduleName: 'teams', methodName: 'addChannelMember' },
        // Teams module tools - Channel file operations
        listChannelFiles: { moduleName: 'teams', methodName: 'listChannelFiles' },
        uploadFileToChannel: { moduleName: 'teams', methodName: 'uploadFileToChannel' },
        readChannelFile: { moduleName: 'teams', methodName: 'readChannelFile' },
        // Teams module tools - Meeting operations
        createOnlineMeeting: { moduleName: 'teams', methodName: 'createOnlineMeeting' },
        createTeamsMeeting: { moduleName: 'teams', methodName: 'createOnlineMeeting' },
        getOnlineMeeting: { moduleName: 'teams', methodName: 'getOnlineMeeting' },
        getMeetingByJoinUrl: { moduleName: 'teams', methodName: 'getMeetingByJoinUrl' },
        listOnlineMeetings: { moduleName: 'teams', methodName: 'listOnlineMeetings' },
        // Teams module tools - Transcript operations
        getMeetingTranscripts: { moduleName: 'teams', methodName: 'getMeetingTranscripts' },
        getMeetingTranscriptContent: { moduleName: 'teams', methodName: 'getMeetingTranscriptContent' },

        // Query module
        query: { moduleName: 'query', methodName: 'processQuery' },

        // Excel consolidated tools (MCP-facing)
        excelSession: { moduleName: 'excel', methodName: 'excelSession' },
        excelWorksheet: { moduleName: 'excel', methodName: 'excelWorksheet' },
        excelRange: { moduleName: 'excel', methodName: 'excelRange' },
        excelTable: { moduleName: 'excel', methodName: 'excelTable' },
        excelFunction: { moduleName: 'excel', methodName: 'excelFunction' },
        // Word consolidated tool (MCP-facing)
        wordDocument: { moduleName: 'word', methodName: 'wordDocument' },
        // PowerPoint consolidated tool (MCP-facing)
        powerpointPresentation: { moduleName: 'powerpoint', methodName: 'powerpointPresentation' },

        // Legacy Excel module tools (kept for REST API backward compatibility)
        createWorkbookSession: { moduleName: 'excel', methodName: 'createWorkbookSession' },
        closeWorkbookSession: { moduleName: 'excel', methodName: 'closeWorkbookSession' },
        listWorksheets: { moduleName: 'excel', methodName: 'listWorksheets' },
        addWorksheet: { moduleName: 'excel', methodName: 'addWorksheet' },
        getWorksheet: { moduleName: 'excel', methodName: 'getWorksheet' },
        updateWorksheet: { moduleName: 'excel', methodName: 'updateWorksheet' },
        deleteWorksheet: { moduleName: 'excel', methodName: 'deleteWorksheet' },
        getRange: { moduleName: 'excel', methodName: 'getRange' },
        updateRange: { moduleName: 'excel', methodName: 'updateRange' },
        getRangeFormat: { moduleName: 'excel', methodName: 'getRangeFormat' },
        updateRangeFormat: { moduleName: 'excel', methodName: 'updateRangeFormat' },
        sortRange: { moduleName: 'excel', methodName: 'sortRange' },
        mergeRange: { moduleName: 'excel', methodName: 'mergeRange' },
        unmergeRange: { moduleName: 'excel', methodName: 'unmergeRange' },
        listTables: { moduleName: 'excel', methodName: 'listTables' },
        createTable: { moduleName: 'excel', methodName: 'createTable' },
        updateTable: { moduleName: 'excel', methodName: 'updateTable' },
        deleteTable: { moduleName: 'excel', methodName: 'deleteTable' },
        listTableRows: { moduleName: 'excel', methodName: 'listTableRows' },
        addTableRow: { moduleName: 'excel', methodName: 'addTableRow' },
        deleteTableRow: { moduleName: 'excel', methodName: 'deleteTableRow' },
        listTableColumns: { moduleName: 'excel', methodName: 'listTableColumns' },
        addTableColumn: { moduleName: 'excel', methodName: 'addTableColumn' },
        deleteTableColumn: { moduleName: 'excel', methodName: 'deleteTableColumn' },
        sortTable: { moduleName: 'excel', methodName: 'sortTable' },
        filterTable: { moduleName: 'excel', methodName: 'filterTable' },
        clearTableFilter: { moduleName: 'excel', methodName: 'clearTableFilter' },
        convertTableToRange: { moduleName: 'excel', methodName: 'convertTableToRange' },
        callWorkbookFunction: { moduleName: 'excel', methodName: 'callWorkbookFunction' },
        calculateWorkbook: { moduleName: 'excel', methodName: 'calculateWorkbook' },

        // Legacy Word module tools (kept for REST API backward compatibility)
        createWordDocument: { moduleName: 'word', methodName: 'createWordDocument' },
        readWordDocument: { moduleName: 'word', methodName: 'readWordDocument' },
        convertDocumentToPdf: { moduleName: 'word', methodName: 'convertDocumentToPdf' },
        getWordDocumentMetadata: { moduleName: 'word', methodName: 'getWordDocumentMetadata' },
        getWordDocumentAsHtml: { moduleName: 'word', methodName: 'getWordDocumentAsHtml' },

        // Legacy PowerPoint module tools (kept for REST API backward compatibility)
        createPresentation: { moduleName: 'powerpoint', methodName: 'createPresentation' },
        readPresentation: { moduleName: 'powerpoint', methodName: 'readPresentation' },
        convertPresentationToPdf: { moduleName: 'powerpoint', methodName: 'convertPresentationToPdf' },
        getPresentationMetadata: { moduleName: 'powerpoint', methodName: 'getPresentationMetadata' }
    };

    /**
     * Generates a tool definition from a module capability
     * @param {string} moduleName - Name of the module
     * @param {string} capability - Capability/tool name
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [sessionId] - Session ID for context
     * @returns {object} Tool definition
     */
    function generateToolDefinition(moduleName, capability, userId = null, sessionId = null) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Generating tool definition', {
                moduleName,
                capability,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        }
        
        try {
        // TODO: [generateToolDefinition] Ensure endpoints align with src/api/routes.cjs (HIGH).
        // This current endpoint generation logic is temporary and brittle.
        // It should be refactored to consume route definitions from routes.cjs or a shared config
        // once src/api/routes.cjs is refactored to export them cleanly.
        // IMPORTANT: All tool schemas must match backend validation exactly

        // Derive default HTTP method based on capability name convention
        let defaultMethod = 'GET';
        if (capability.startsWith('create') || capability.startsWith('add') || capability.startsWith('send') || capability.startsWith('search') || capability.startsWith('flag')) {
            defaultMethod = 'POST';
        } else if (capability.startsWith('update') || capability.startsWith('set')) {
            defaultMethod = 'PUT'; // Or PATCH, depending on API design
        } else if (capability.startsWith('delete') || capability.startsWith('remove')) {
            defaultMethod = 'DELETE';
        }

        // Default tool structure
        const toolDef = {
            name: capability,
            description: `${capability} operation for ${moduleName}`,
            endpoint: `/api/v1/${moduleName.toLowerCase()}/${capability}`, // Placeholder endpoint
            method: defaultMethod,
            parameters: {}
        };

        // Customize based on known capabilities
        switch (capability) {
            // Mail tools
            case 'getInbox':
            case 'getMail':
                toolDef.description = 'Fetch mail from Microsoft 365 inbox';
                toolDef.endpoint = '/api/v1/mail';
                toolDef.parameters = {
                    limit: { type: 'number', description: 'Maximum number of messages to retrieve', optional: true, default: 20 },
                    filter: { type: 'string', description: 'Filter string for messages', optional: true },
                    debug: { type: 'boolean', description: 'Enable debug mode to return raw message data', optional: true, default: false }
                };
                break;
            case 'sendEmail':
            case 'sendMail':
                toolDef.description = 'Send an email via Microsoft 365';
                toolDef.endpoint = '/api/v1/mail/send';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    to: { 
                        type: 'string', 
                        description: 'Recipient email address(es). Can be a single email, comma-separated list, or array of emails',
                        required: true
                    },
                    subject: { 
                        type: 'string', 
                        description: 'Email subject line', 
                        required: true,
                        minLength: 1
                    },
                    body: { 
                        type: 'string', 
                        description: 'Email body content', 
                        required: true,
                        minLength: 1
                    },
                    cc: { 
                        type: 'string', 
                        description: 'CC recipient email address(es). Can be a single email, comma-separated list, or array of emails', 
                        optional: true 
                    },
                    bcc: { 
                        type: 'string', 
                        description: 'BCC recipient email address(es). Can be a single email, comma-separated list, or array of emails', 
                        optional: true 
                    },
                    contentType: { 
                        type: 'string', 
                        description: 'Content type of the email body', 
                        optional: true, 
                        enum: ['Text', 'HTML'],
                        default: 'Text'
                    },
                    attachments: { 
                        type: 'array', 
                        description: 'File attachments', 
                        optional: true 
                    }
                };
                break;
            case 'search':
                toolDef.description = `Unified search across Microsoft 365: emails, calendar events, files, and people.

USE THIS TOOL when the user wants to find "everything about" a person, topic, or project. A single search returns results from ALL entity types at once.

PERSON DISAMBIGUATION WORKFLOW:
When searching for a person's name (e.g., "Krister"), you may get multiple matches in the 'person' results.
1. First search with the name to identify the right person (check jobTitle, department, email)
2. Then do a targeted search using their UPN/email for all their content:
   → "from:kristerm@microsoft.com OR to:kristerm@microsoft.com"

EXAMPLES:
- "Find everything about Krister" → Step 1: query: "Krister" to find the person
  → Step 2: query: "from:kristerm@microsoft.com OR to:kristerm@microsoft.com" for all communications
- "What do we have on Project X?" → query: "Project X" (searches all content)
- "Files shared by John" → query: "author:john@company.com", entityTypes: ["driveItem"]

⚠️ FOR DATE-BASED CALENDAR QUERIES: Use getEvents instead!
- "Today's meetings" → Use getEvents (supports date filtering)
- "Events this week" → Use getEvents (supports date filtering)
- "Meetings about budget" → Use search with entityTypes: ["event"] (keyword search only)

KQL QUERY SYNTAX:
- Simple search: "Krister" or "project budget" (searches all fields)
- From person: "from:john@company.com" or "to:john@company.com"
- Subject: "subject:quarterly report"
- Has attachments: "hasAttachments:true"
- Combine: "from:boss@company.com AND subject:urgent"
- File type: "filetype:pdf"
- Author: "author:john@company.com"

ENTITY TYPES (default: all):
- message: emails (from:, to:, subject:, body:, hasAttachments:)
- event: calendar events (keyword search only - use getEvents for dates)
- driveItem: files (filename:, filetype:, path:, author:)
- person: people directory (displayName, email, jobTitle)`;
                toolDef.endpoint = '/api/v1/search';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    query: {
                        type: 'string',
                        description: 'Search query. Use simple terms like "Krister" for broad search, or KQL syntax like "from:email@domain.com" for specific filters.',
                        required: true
                    },
                    entityTypes: {
                        type: 'array',
                        description: 'Entity types to search. Omit to search ALL types (recommended for "find everything about X" queries). Options: message, event, driveItem, person.',
                        optional: true,
                        default: ['message', 'event', 'driveItem', 'person'],
                        items: { type: 'string', enum: ['message', 'chatMessage', 'event', 'driveItem', 'person', 'site', 'list', 'listItem'] }
                    },
                    limit: {
                        type: 'number',
                        description: 'Maximum results per entity type (max 25). Default: 10.',
                        optional: true,
                        default: 10
                    },
                    includeAnswers: {
                        type: 'boolean',
                        description: 'Include enterprise knowledge answers (acronyms, bookmarks, Q&A). Returns definitions for company-specific terms.',
                        optional: true,
                        default: false
                    },
                    enableSpellingModification: {
                        type: 'boolean',
                        description: 'Auto-correct typos in search query. When true, misspelled words are automatically corrected.',
                        optional: true,
                        default: false
                    }
                };
                break;
            case 'flagEmail':
            case 'flagMail':
                toolDef.description = 'Flag or unflag an email';
                toolDef.endpoint = '/api/v1/mail/flag';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Email ID to flag or unflag',
                        required: true
                    },
                    flag: { 
                        type: 'boolean', 
                        description: 'Whether to flag (true) or unflag (false) the email',
                        optional: true,
                        default: true
                    }
                };
                break;
            case 'getAttachments':
            case 'getMailAttachments':
                toolDef.description = 'Get email attachments';
                toolDef.endpoint = '/api/v1/mail/attachments';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Email ID to get attachments for',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inQuery: true }
                };
                break;
            case 'getEmailDetails':
            case 'getMailDetails':
            case 'readMailDetails':
                toolDef.description = 'Get detailed information for a specific email';
                toolDef.endpoint = '/api/v1/mail/:id';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { type: 'string', description: 'Email ID to retrieve details for', required: true }
                };
                // Ensure this tool is properly registered with the /v1/mail/:id endpoint
                // Note: The :id in the path is a placeholder for the actual ID value
                toolDef.parameterMapping = {
                    id: { inPath: true }
                };
                break;
            case 'markAsRead':
            case 'markMailRead':
            case 'markEmailRead':
                toolDef.description = 'Mark an email as read or unread';
                toolDef.endpoint = '/api/v1/mail/:id/read';
                toolDef.method = 'PATCH';
                toolDef.parameters = {
                    id: { type: 'string', description: 'Email ID to mark as read/unread' },
                    isRead: { type: 'boolean', description: 'Whether to mark as read (true) or unread (false)', optional: true, default: true }
                };
                // Ensure this tool is properly registered with the /api/v1/mail/:id/read endpoint
                // Note: The :id in the path is a placeholder for the actual ID value
                toolDef.parameterMapping = {
                    id: { inPath: true },
                    isRead: { inBody: true }
                };
                break;

            case 'addMailAttachment':
                toolDef.description = 'Add an attachment to an existing email';
                toolDef.endpoint = '/api/v1/mail/:id/attachments';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { type: 'string', description: 'Email ID to add attachment to' },
                    name: { type: 'string', description: 'Name of the attachment file' },
                    contentBytes: { type: 'string', description: 'Base64 encoded content of the attachment' },
                    contentType: { type: 'string', description: 'MIME type of the attachment', optional: true },
                    isInline: { type: 'boolean', description: 'Whether the attachment is inline', optional: true, default: false }
                };
                toolDef.parameterMapping = {
                    id: { inPath: true },
                    name: { inBody: true },
                    contentBytes: { inBody: true },
                    contentType: { inBody: true },
                    isInline: { inBody: true }
                };
                break;
            case 'removeMailAttachment':
                toolDef.description = 'Remove an attachment from an existing email';
                toolDef.endpoint = '/api/v1/mail/:id/attachments/:attachmentId';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    id: { type: 'string', description: 'Email ID to remove attachment from' },
                    attachmentId: { type: 'string', description: 'ID of the attachment to remove' }
                };
                toolDef.parameterMapping = {
                    id: { inPath: true },
                    attachmentId: { inPath: true }
                };
                break;

            // Calendar tools
            case 'getEvents':
            case 'getCalendar':
                toolDef.description = `Get calendar events by DATE RANGE. This is the tool for "today's meetings", "this week's events", etc.

USE THIS TOOL FOR:
- "What's on my calendar today?" → start: today, end: today
- "Show me this week's meetings" → start: Monday, end: Friday
- "My schedule for tomorrow" → start: tomorrow, end: tomorrow

USE 'search' TOOL INSTEAD FOR:
- "Find meetings about budget" → search with entityTypes: ["event"]
- "Meetings with Krister" → search with query: "organizer:kristerm@microsoft.com"

This endpoint uses Microsoft Graph's calendarView which properly expands recurring events within the date range.`;
                toolDef.endpoint = '/api/v1/calendar';
                toolDef.parameters = {
                    start: {
                        type: 'string',
                        description: 'Start date (YYYY-MM-DD). Required for date-based queries like "today" or "this week".',
                        optional: true,
                        format: 'date'
                    },
                    end: {
                        type: 'string',
                        description: 'End date (YYYY-MM-DD). Required for date-based queries.',
                        optional: true,
                        format: 'date'
                    },
                    limit: {
                        type: 'number',
                        description: 'Max events to return',
                        optional: true,
                        default: 50
                    },
                    filter: {
                        type: 'string',
                        description: 'OData $filter query',
                        optional: true
                    },
                    select: {
                        type: 'string',
                        description: 'Properties to include (comma-separated)',
                        optional: true
                    },
                    orderby: {
                        type: 'string',
                        description: 'Sort by property',
                        optional: true,
                        default: 'start/dateTime'
                    },
                    subject: {
                        type: 'string',
                        description: 'Filter by subject text',
                        optional: true
                    },
                    organizer: {
                        type: 'string',
                        description: 'Filter by organizer name',
                        optional: true
                    },
                    location: {
                        type: 'string',
                        description: 'Filter by location text',
                        optional: true
                    },
                    
                    // Time-based filters
                    timeframe: { 
                        type: 'string', 
                        description: 'Predefined time range (today, tomorrow, this_week, next_week, this_month)', 
                        optional: true,
                        enum: ['today', 'tomorrow', 'this_week', 'next_week', 'this_month', 'next_month']
                    },
                    
                    // Response options
                    debug: { 
                        type: 'boolean', 
                        description: 'Enable debug mode to return additional metadata', 
                        optional: true, 
                        default: false 
                    }
                };
                break;
            case 'createEvent':
                toolDef.description = 'Create a new calendar event';
                toolDef.endpoint = '/api/v1/calendar/events';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    subject: { type: 'string', description: 'Event subject/title', required: true },
                    start: {
                        type: 'object',
                        description: 'Start time',
                        required: true,
                        properties: {
                            dateTime: { type: 'string', description: 'ISO date string', required: true },
                            timeZone: { type: 'string', description: 'Time zone', optional: true }
                        }
                    },
                    end: {
                        type: 'object',
                        description: 'End time',
                        required: true,
                        properties: {
                            dateTime: { type: 'string', description: 'ISO date string', required: true },
                            timeZone: { type: 'string', description: 'Time zone', optional: true }
                        }
                    },
                    location: {
                        type: 'object',
                        description: 'Event location (object with displayName, omit or do not send if null)',
                        optional: true,
                        properties: {
                            displayName: { type: 'string', description: 'Location display name', required: true }
                        }
                    },
                    body: {
                        type: 'object',
                        description: 'Event description/body content (object with content and optional contentType)',
                        optional: true,
                        properties: {
                            content: { type: 'string', description: 'Body content text', required: true },
                            contentType: { type: 'string', description: 'Content type (text or html)', optional: true, default: 'text' }
                        }
                    },
                    attendees: {
                        type: 'array',
                        description: 'Array of attendee objects (with emailAddress.address)',
                        optional: true,
                        items: {
                            type: 'object',
                            properties: {
                                emailAddress: { type: 'object', properties: { address: { type: 'string', required: true } }, required: true },
                                type: { type: 'string', description: 'Attendee type (required, optional, resource)', optional: true }
                            }
                        }
                    },
                    isOnlineMeeting: { type: 'boolean', description: 'Whether this is an online meeting', optional: true }
                };
                break;
            case 'updateEvent':
                toolDef.description = 'Update an existing calendar event';
                toolDef.endpoint = '/api/v1/calendar/events/:id';
                toolDef.method = 'PUT';
                toolDef.parameters = {
                    id: { type: 'string', description: 'Event ID to update', required: true },
                    subject: { type: 'string', description: 'Event subject/title', optional: true },
                    start: {
                        type: 'object',
                        description: 'Start time',
                        optional: true,
                        properties: {
                            dateTime: { type: 'string', description: 'ISO date string', required: true },
                            timeZone: { type: 'string', description: 'Time zone', optional: true }
                        }
                    },
                    end: {
                        type: 'object',
                        description: 'End time',
                        optional: true,
                        properties: {
                            dateTime: { type: 'string', description: 'ISO date string', required: true },
                            timeZone: { type: 'string', description: 'Time zone', optional: true }
                        }
                    },
                    location: {
                        type: 'object',
                        description: 'Event location (object with displayName, omit or do not send if null)',
                        optional: true,
                        properties: {
                            displayName: { type: 'string', description: 'Location display name', required: true }
                        }
                    },
                    body: {
                        type: 'object',
                        description: 'Event description/body content (object with content and optional contentType)',
                        optional: true,
                        properties: {
                            content: { type: 'string', description: 'Body content text', required: true },
                            contentType: { type: 'string', description: 'Content type (text or html)', optional: true, default: 'text' }
                        }
                    },
                    attendees: {
                        type: 'array',
                        description: 'Array of attendee objects (with emailAddress.address)',
                        optional: true,
                        items: {
                            type: 'object',
                            properties: {
                                emailAddress: { type: 'object', properties: { address: { type: 'string', required: true } }, required: true },
                                type: { type: 'string', description: 'Attendee type (required, optional, resource)', optional: true }
                            }
                        }
                    },
                    isAllDay: { type: 'boolean', description: 'Whether this is an all-day event', optional: true },
                    isOnlineMeeting: { type: 'boolean', description: 'Whether this is an online meeting', optional: true }
                };
                break;
            case 'deleteEvent':
            case 'cancelEvent': // Alias
                toolDef.description = 'Delete or cancel a calendar event';
                toolDef.endpoint = '/api/v1/calendar/events/:id/cancel'; // Correct endpoint
                toolDef.method = 'POST'; // Correct method (POST, not DELETE)
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Event ID to cancel', 
                        required: true,
                        inPath: true 
                    },
                    comment: { 
                        type: 'string', 
                        description: 'Optional cancellation comment', 
                        optional: true 
                    }
                };
                // Ensure this tool is properly registered with the /api/v1/calendar/events/:id/cancel endpoint
                // Note: The :id in the path is a placeholder for the actual ID value
                toolDef.parameterMapping = {
                    id: { inPath: true },
                    comment: { inBody: true }
                };
                break;
            case 'acceptEvent':
                toolDef.description = 'Accept a calendar event invitation. Note: This only works for events where the user is an attendee, not the organizer.';
                toolDef.endpoint = '/api/v1/calendar/events/:id/accept';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Event ID to accept', 
                        required: true,
                        inPath: true 
                    },
                    comment: { 
                        type: 'string', 
                        description: 'Optional comment to include with the acceptance', 
                        optional: true 
                    }
                };
                break;
            case 'declineEvent':
                toolDef.description = 'Decline a calendar event invitation. Note: This only works for events where the user is an attendee, not the organizer.';
                toolDef.endpoint = '/api/v1/calendar/events/:id/decline';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Event ID to decline', 
                        required: true,
                        inPath: true 
                    },
                    comment: { 
                        type: 'string', 
                        description: 'Optional comment to include with the decline', 
                        optional: true 
                    }
                };
                break;
            case 'tentativelyAcceptEvent':
                toolDef.description = 'Tentatively accept a calendar event invitation. Note: This only works for events where the user is an attendee, not the organizer.';
                toolDef.endpoint = '/api/v1/calendar/events/:id/tentativelyAccept';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Event ID to tentatively accept', 
                        required: true,
                        inPath: true 
                    },
                    comment: { 
                        type: 'string', 
                        description: 'Optional comment to include with the tentative acceptance', 
                        optional: true 
                    }
                };
                break;
            case 'getAvailability':
                toolDef.description = 'Get availability information for specified users and time slots. This tool helps identify when people are free or busy before scheduling meetings.';
                toolDef.endpoint = '/api/v1/calendar/availability';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    users: { 
                        type: 'array', 
                        description: 'Array of user email addresses to check availability for (must be valid email addresses)', 
                        required: true 
                    },
                    timeSlots: { 
                        type: 'array', 
                        description: 'Array of time slots to check availability within', 
                        required: true,
                        items: {
                            type: 'object',
                            properties: {
                                start: {
                                    type: 'object',
                                    description: 'Start time',
                                    required: true,
                                    properties: {
                                        dateTime: { 
                                            type: 'string', 
                                            format: 'date-time', 
                                            description: 'Start date/time in ISO format (e.g., 2025-05-02T14:00:00)', 
                                            required: true 
                                        },
                                        timeZone: { 
                                            type: 'string', 
                                            description: 'Time zone (e.g., UTC, Europe/Oslo)', 
                                            optional: true, 
                                            default: 'UTC' 
                                        }
                                    }
                                },
                                end: {
                                    type: 'object',
                                    description: 'End time',
                                    required: true,
                                    properties: {
                                        dateTime: { 
                                            type: 'string', 
                                            format: 'date-time', 
                                            description: 'End date/time in ISO format (e.g., 2025-05-02T15:00:00)', 
                                            required: true 
                                        },
                                        timeZone: { 
                                            type: 'string', 
                                            description: 'Time zone (e.g., UTC, Europe/Oslo)', 
                                            optional: true, 
                                            default: 'UTC' 
                                        }
                                    }
                                }
                            }
                        }
                    },
                    // Support for simpler API calls with direct start/end parameters
                    start: { 
                        type: 'string', 
                        format: 'date-time', 
                        description: 'Alternative to timeSlots: Start date/time in ISO format for a single time slot', 
                        optional: true 
                    },
                    end: { 
                        type: 'string', 
                        format: 'date-time', 
                        description: 'Alternative to timeSlots: End date/time in ISO format for a single time slot', 
                        optional: true 
                    }
                };
                toolDef.parameterMapping = {
                    users: { inBody: true },
                    timeSlots: { inBody: true },
                    start: { inBody: true },
                    end: { inBody: true }
                };
                break;
                case 'findMeetingTimes':
                    toolDef.description = 'Find suggested meeting times based on attendees and constraints';
                    toolDef.endpoint = '/api/v1/calendar/findMeetingTimes';
                    toolDef.method = 'POST';
                    toolDef.parameters = {
                        attendees: { 
                            type: 'array', 
                            description: 'Array of attendee email addresses', 
                            items: { type: 'string', format: 'email' },
                            required: true,
                            minItems: 1
                        },
                        timeConstraints: { 
                            type: 'object', 
                            description: 'Time constraints for the meeting',
                            required: true,
                            properties: {
                                activityDomain: { 
                                    type: 'string', 
                                    description: 'Activity domain (work/personal/unrestricted)', 
                                    optional: true, 
                                    default: 'work',
                                    enum: ['work', 'personal', 'unrestricted']
                                },
                                timeslots: { 
                                    type: 'array', 
                                    description: 'Array of time slots',
                                    required: true,
                                    items: {
                                        type: 'object',
                                        properties: {
                                            start: {
                                                type: 'object',
                                                description: 'Start time',
                                                required: true,
                                                properties: {
                                                    dateTime: { type: 'string', description: 'ISO date string (e.g. 2025-05-26T09:00:00)', required: true },
                                                    timeZone: { type: 'string', description: 'Time zone', optional: true, default: 'UTC' }
                                                }
                                            },
                                            end: {
                                                type: 'object',
                                                description: 'End time',
                                                required: true,
                                                properties: {
                                                    dateTime: { type: 'string', description: 'ISO date string (e.g. 2025-05-26T17:00:00)', required: true },
                                                    timeZone: { type: 'string', description: 'Time zone', optional: true, default: 'UTC' }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        },
                        locationConstraint: {
                            type: 'object',
                            description: 'Location constraints for the meeting',
                            optional: true,
                            properties: {
                                isRequired: { type: 'boolean', description: 'Whether a location is required', optional: true, default: false },
                                suggestLocation: { type: 'boolean', description: 'Whether to suggest a location', optional: true, default: false },
                                locations: {
                                    type: 'array',
                                    description: 'Array of potential locations',
                                    optional: true,
                                    items: {
                                        type: 'object',
                                        properties: {
                                            displayName: { type: 'string', description: 'Display name of the location', required: true },
                                            locationEmailAddress: { type: 'string', description: 'Email address of the location', optional: true }
                                        }
                                    }
                                }
                            }
                        },
                        meetingDuration: { 
                            type: 'string', 
                            description: 'Duration in ISO8601 format (e.g., PT1H for 1 hour, PT30M for 30 minutes)', 
                            optional: true, 
                            default: 'PT30M' 
                        },
                        maxCandidates: { 
                            type: 'number', 
                            description: 'Maximum number of meeting time suggestions', 
                            optional: true, 
                            min: 1, 
                            max: 100, 
                            default: 10 
                        }
                    };
                break;
            case 'getRooms':
                toolDef.description = 'Get available meeting rooms';
                toolDef.endpoint = '/api/v1/calendar/rooms';
                toolDef.parameters = { /* ... specific params ... */ };
                break;
            case 'addAttachment':
                toolDef.description = 'Add attachment to a calendar event';
                toolDef.endpoint = '/api/v1/calendar/events/:id/attachments';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'Event ID to add attachment to',
                        required: true
                    },
                    name: { 
                        type: 'string', 
                        description: 'Name of the attachment file',
                        required: true
                    },
                    contentBytes: { 
                        type: 'string', 
                        description: 'Base64-encoded file content',
                        required: true
                    },
                    contentType: { 
                        type: 'string', 
                        description: 'MIME type of the attachment',
                        optional: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inPath: true },
                    name: { inBody: true },
                    contentBytes: { inBody: true },
                    contentType: { inBody: true }
                };
                break;
            case 'removeAttachment':
                toolDef.description = 'Remove attachment from a calendar event';
                toolDef.endpoint = '/api/v1/calendar/events/:eventId/attachments/:attachmentId';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    eventId: { 
                        type: 'string', 
                        description: 'Event ID to remove attachment from',
                        required: true
                    },
                    attachmentId: { 
                        type: 'string', 
                        description: 'Attachment ID to remove',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    eventId: { inPath: true },
                    attachmentId: { inPath: true }
                };
                break;

            // File tools (OneDrive/SharePoint)
            case 'listFiles':
                toolDef.description = 'List files in a specific drive or folder';
                toolDef.endpoint = '/api/v1/files';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    parentId: { 
                        type: 'string', 
                        description: 'ID of the parent folder to list files from. If not provided, lists files from the root folder.',
                        optional: true
                    }
                };
                toolDef.parameterMapping = {
                    parentId: { inQuery: true }
                };
                break;
            case 'uploadFile':
                toolDef.description = 'Upload a file to OneDrive or SharePoint';
                toolDef.endpoint = '/api/v1/files/upload';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    name: { 
                        type: 'string', 
                        description: 'Name of the file to upload',
                        required: true
                    },
                    content: { 
                        type: 'string', 
                        description: 'Content of the file to upload',
                        required: true
                    }
                };
                break;
            case 'downloadFile':
                toolDef.description = 'Download a file from OneDrive or SharePoint';
                toolDef.endpoint = '/api/v1/files/download';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the file to download',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inQuery: true }
                };
                break;
            case 'getFileMetadata':
                toolDef.description = 'Get metadata for a specific file';
                toolDef.endpoint = '/api/v1/files/metadata';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the file to get metadata for',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inQuery: true }
                };
                break;
            case 'getFileContent':
                toolDef.description = 'Get the content of a specific file. Use search with entityTypes: ["driveItem"] or listFiles to find the file ID.';
                toolDef.endpoint = '/api/v1/files/content';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the file to get content for (required, must be obtained from search or listFiles)',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inQuery: true }
                };
                break;
            case 'setFileContent':
                toolDef.description = 'Set the content of a specific file';
                toolDef.endpoint = '/api/v1/files/content';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { 
                        type: 'string', 
                        description: 'ID of the file to set content for',
                        required: true
                    },
                    content: { 
                        type: 'string', 
                        description: 'New content for the file',
                        required: true
                    }
                };
                break;
            case 'updateFileContent':
                toolDef.description = 'Update the content of a specific file';
                toolDef.endpoint = '/api/v1/files/content/update';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { 
                        type: 'string', 
                        description: 'ID of the file to update content for',
                        required: true
                    },
                    content: { 
                        type: 'string', 
                        description: 'New content for the file',
                        required: true
                    }
                };
                break;
            case 'deleteFile':
                toolDef.description = 'Delete a file or folder';
                toolDef.endpoint = '/api/v1/files/:id';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the file or folder to delete',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inPath: true }
                };
                break;
            case 'createSharingLink':
                toolDef.description = 'Create a sharing link for a file';
                toolDef.endpoint = '/api/v1/files/share';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { 
                        type: 'string', 
                        description: 'ID of the file to create a sharing link for',
                        required: true
                    },
                    type: { 
                        type: 'string', 
                        description: 'Type of sharing link (view or edit)',
                        enum: ['view', 'edit'],
                        default: 'view'
                    }
                };
                break;
            case 'getSharingLinks':
                toolDef.description = 'Get sharing links for a file';
                toolDef.endpoint = '/api/v1/files/sharing';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { 
                        type: 'string', 
                        description: 'ID of the file to get sharing links for',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    fileId: { inQuery: true }
                };
                break;
            case 'removeSharingPermission':
                toolDef.description = 'Remove a sharing permission from a file';
                toolDef.endpoint = '/api/v1/files/sharing/remove';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { 
                        type: 'string', 
                        description: 'ID of the file to remove sharing permission from',
                        required: true
                    },
                    permissionId: { 
                        type: 'string', 
                        description: 'ID of the permission to remove',
                        required: true
                    }
                };
                break;

            // Query tool

            // People tools
            case 'findPeople':
                toolDef.description = 'Find people by name or email. Use this to get email addresses for scheduling or sending emails.';
                toolDef.endpoint = '/api/v1/people/find';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    query: { 
                        type: 'string', 
                        description: 'Search query to find a person',
                        optional: true
                    },
                    name: { 
                        type: 'string', 
                        description: 'Person name to search for', 
                        optional: true 
                    },
                    limit: { 
                        type: 'number', 
                        description: 'Maximum number of results', 
                        optional: true,
                        default: 10
                    }
                };
                toolDef.parameterMapping = {
                    query: { inQuery: true },
                    name: { inQuery: true },
                    limit: { inQuery: true }
                };
                break;
            case 'getRelevantPeople':
                toolDef.description = 'Get people relevant to the user';
                toolDef.endpoint = '/api/v1/people';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    limit: { 
                        type: 'number', 
                        description: 'Maximum number of people to return', 
                        optional: true,
                        default: 10
                    },
                    filter: { 
                        type: 'string', 
                        description: 'Filter criteria', 
                        optional: true 
                    },
                    orderby: { 
                        type: 'string', 
                        description: 'Order by field', 
                        optional: true 
                    }
                };
                toolDef.parameterMapping = {
                    limit: { inQuery: true },
                    filter: { inQuery: true },
                    orderby: { inQuery: true }
                };
                break;
            case 'getPersonById':
                toolDef.description = 'Get a specific person by ID';
                toolDef.endpoint = '/api/v1/people/:id';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the person to retrieve',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    id: { inPath: true }
                };
                break;

            // NOTE: 'search' case is defined earlier in this switch (line ~319)
            // Legacy search tools (searchMail, searchFiles) removed - use unified 'search' tool with entityTypes parameter

            // ================================================================
            // TEAMS MODULE TOOLS
            // ================================================================

            // Teams Chat Tools
            case 'listChats':
            case 'getChats':
                toolDef.description = 'List the user\'s Microsoft Teams chats. Returns recent chat conversations including one-on-one chats and group chats. Use this to find chat IDs before sending messages.';
                toolDef.endpoint = '/api/v1/teams/chats';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    limit: {
                        type: 'number',
                        description: 'Maximum number of chats to retrieve (default: 20, max: 50)',
                        optional: true,
                        default: 20
                    }
                };
                break;

            case 'getChatMessages':
                toolDef.description = 'Get messages from a specific Teams chat. Requires a chat ID which can be obtained from listChats.';
                toolDef.endpoint = '/api/v1/teams/chats/:chatId/messages';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    chatId: {
                        type: 'string',
                        description: 'The ID of the chat to retrieve messages from (required, get from listChats)',
                        required: true
                    },
                    limit: {
                        type: 'number',
                        description: 'Maximum number of messages to retrieve (default: 50)',
                        optional: true,
                        default: 50
                    }
                };
                toolDef.parameterMapping = { chatId: { inPath: true } };
                break;

            case 'sendChatMessage':
                toolDef.description = 'Send a message to a Teams chat. Requires a chat ID from listChats.';
                toolDef.endpoint = '/api/v1/teams/chats/:chatId/messages';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    chatId: {
                        type: 'string',
                        description: 'The ID of the chat to send the message to',
                        required: true
                    },
                    content: {
                        type: 'string',
                        description: 'The message content to send',
                        required: true
                    },
                    contentType: {
                        type: 'string',
                        description: 'Content type: text or html',
                        optional: true,
                        enum: ['text', 'html'],
                        default: 'text'
                    }
                };
                toolDef.parameterMapping = { chatId: { inPath: true }, content: { inBody: true }, contentType: { inBody: true } };
                break;

            // Teams & Channel Tools
            case 'listJoinedTeams':
            case 'getTeams':
                toolDef.description = 'List all Microsoft Teams that the user has joined. Use this to get team IDs for channel operations.';
                toolDef.endpoint = '/api/v1/teams';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    limit: {
                        type: 'number',
                        description: 'Maximum number of teams to retrieve',
                        optional: true,
                        default: 100
                    }
                };
                break;

            case 'listTeamChannels':
            case 'getTeamChannels':
                toolDef.description = 'List channels in a Microsoft Teams team. Requires a team ID which can be obtained from listJoinedTeams.';
                toolDef.endpoint = '/api/v1/teams/:teamId/channels';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    teamId: {
                        type: 'string',
                        description: 'The ID of the team to list channels from (required, get from listJoinedTeams)',
                        required: true
                    }
                };
                toolDef.parameterMapping = { teamId: { inPath: true } };
                break;

            case 'getChannelMessages':
                toolDef.description = 'Get messages from a Teams channel. Requires team ID and channel ID.';
                toolDef.endpoint = '/api/v1/teams/:teamId/channels/:channelId/messages';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    teamId: {
                        type: 'string',
                        description: 'The ID of the team',
                        required: true
                    },
                    channelId: {
                        type: 'string',
                        description: 'The ID of the channel to retrieve messages from',
                        required: true
                    },
                    limit: {
                        type: 'number',
                        description: 'Maximum number of messages to retrieve (default: 50)',
                        optional: true,
                        default: 50
                    }
                };
                toolDef.parameterMapping = { teamId: { inPath: true }, channelId: { inPath: true } };
                break;

            case 'sendChannelMessage':
                toolDef.description = 'Send a message to a Teams channel. Requires team ID and channel ID.';
                toolDef.endpoint = '/api/v1/teams/:teamId/channels/:channelId/messages';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    teamId: {
                        type: 'string',
                        description: 'The ID of the team',
                        required: true
                    },
                    channelId: {
                        type: 'string',
                        description: 'The ID of the channel to send the message to',
                        required: true
                    },
                    content: {
                        type: 'string',
                        description: 'The message content to send',
                        required: true
                    },
                    contentType: {
                        type: 'string',
                        description: 'Content type: text or html',
                        optional: true,
                        enum: ['text', 'html'],
                        default: 'text'
                    },
                    subject: {
                        type: 'string',
                        description: 'Subject line for the message (optional)',
                        optional: true
                    }
                };
                toolDef.parameterMapping = { teamId: { inPath: true }, channelId: { inPath: true }, content: { inBody: true }, contentType: { inBody: true }, subject: { inBody: true } };
                break;

            case 'replyToMessage':
                toolDef.description = 'Reply to a message in a Teams channel thread.';
                toolDef.endpoint = '/api/v1/teams/:teamId/channels/:channelId/messages/:messageId/replies';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    teamId: {
                        type: 'string',
                        description: 'The ID of the team',
                        required: true
                    },
                    channelId: {
                        type: 'string',
                        description: 'The ID of the channel',
                        required: true
                    },
                    messageId: {
                        type: 'string',
                        description: 'The ID of the message to reply to',
                        required: true
                    },
                    content: {
                        type: 'string',
                        description: 'The reply content',
                        required: true
                    },
                    contentType: {
                        type: 'string',
                        description: 'Content type: text or html',
                        optional: true,
                        enum: ['text', 'html'],
                        default: 'text'
                    }
                };
                toolDef.parameterMapping = { teamId: { inPath: true }, channelId: { inPath: true }, messageId: { inPath: true }, content: { inBody: true }, contentType: { inBody: true } };
                break;

            // Teams Meeting Tools
            case 'createOnlineMeeting':
            case 'createTeamsMeeting':
                toolDef.description = 'Create a new Microsoft Teams online meeting. Returns meeting details including the join URL. Use this for ad-hoc meetings. For scheduled meetings with calendar invite, use createEvent with isOnlineMeeting: true instead.';
                toolDef.endpoint = '/api/v1/teams/meetings';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    subject: {
                        type: 'string',
                        description: 'Subject/title of the meeting',
                        required: true
                    },
                    startDateTime: {
                        type: 'string',
                        description: 'Meeting start time in ISO 8601 format (e.g., 2025-01-15T10:00:00Z)',
                        required: true,
                        format: 'date-time'
                    },
                    endDateTime: {
                        type: 'string',
                        description: 'Meeting end time in ISO 8601 format',
                        required: true,
                        format: 'date-time'
                    },
                    participants: {
                        type: 'array',
                        description: 'Array of participant email addresses',
                        optional: true,
                        items: { type: 'string', format: 'email' }
                    },
                    lobbyBypassSettings: {
                        type: 'string',
                        description: 'Who can bypass the lobby: everyone, organization, organizationAndFederated, organizer',
                        optional: true,
                        enum: ['everyone', 'organization', 'organizationAndFederated', 'organizer'],
                        default: 'organization'
                    }
                };
                break;

            case 'getOnlineMeeting':
                toolDef.description = 'Get details of a specific Teams online meeting by ID.';
                toolDef.endpoint = '/api/v1/teams/meetings/:meetingId';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    meetingId: {
                        type: 'string',
                        description: 'The ID of the meeting to retrieve',
                        required: true
                    }
                };
                toolDef.parameterMapping = { meetingId: { inPath: true } };
                break;

            case 'getMeetingByJoinUrl':
                toolDef.description = 'Find a Teams meeting by its join URL.';
                toolDef.endpoint = '/api/v1/teams/meetings/findByJoinUrl';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    joinUrl: {
                        type: 'string',
                        description: 'The Teams meeting join URL',
                        required: true
                    }
                };
                toolDef.parameterMapping = { joinUrl: { inQuery: true } };
                break;

            case 'listOnlineMeetings':
                toolDef.description = 'List the user\'s scheduled Teams online meetings.';
                toolDef.endpoint = '/api/v1/teams/meetings';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    limit: {
                        type: 'number',
                        description: 'Maximum number of meetings to retrieve',
                        optional: true,
                        default: 20
                    }
                };
                break;

            case 'getMeetingTranscripts':
                toolDef.description = 'Get all transcripts for a Teams online meeting. Returns a list of available transcripts with metadata. Requires OnlineMeetingTranscript.Read.All permission.';
                toolDef.endpoint = '/api/v1/teams/meetings/{meetingId}/transcripts';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    meetingId: {
                        type: 'string',
                        description: 'The ID of the online meeting to get transcripts for',
                        required: true
                    }
                };
                break;

            case 'getMeetingTranscriptContent':
                toolDef.description = 'Get the full content of a specific meeting transcript. Returns parsed transcript entries with speaker attribution and timestamps. Use getMeetingTranscripts first to get available transcript IDs.';
                toolDef.endpoint = '/api/v1/teams/meetings/{meetingId}/transcripts/{transcriptId}';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    meetingId: {
                        type: 'string',
                        description: 'The ID of the online meeting',
                        required: true
                    },
                    transcriptId: {
                        type: 'string',
                        description: 'The ID of the transcript to retrieve',
                        required: true
                    }
                };
                break;

            // ========== Excel Workbook Tools ==========
            // ========== Consolidated compound tools (MCP-facing, 7 tools replace 39) ==========
            case 'excelSession':
                toolDef.description = 'Manage Excel workbook sessions for efficient batch operations.\n\nActions:\n- create: Create a persistent or temporary session (requires: fileId; optional: persistent, default true)\n- close: Close an active session (requires: fileId)\n\nSessions are cached per (user, fileId) with a 4-minute TTL. Only .xlsx files are supported.';
                toolDef.endpoint = '/api/v1/excel/session/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['create', 'close'] },
                    fileId: { type: 'string', description: 'OneDrive/SharePoint drive item ID of the .xlsx file', required: true },
                    persistent: { type: 'boolean', description: 'Whether changes should be saved (true) or temporary (false). Only for create action.', optional: true, default: true }
                };
                break;
            case 'excelWorksheet':
                toolDef.description = 'Manage worksheets in an Excel workbook.\n\nActions:\n- list: List all worksheets (requires: fileId)\n- add: Add a new worksheet (requires: fileId, name)\n- get: Get a specific worksheet (requires: fileId, sheetIdOrName)\n- update: Update worksheet properties like name, position, visibility (requires: fileId, sheetIdOrName, properties)\n- delete: Delete a worksheet (requires: fileId, sheetIdOrName)';
                toolDef.endpoint = '/api/v1/excel/worksheet/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['list', 'add', 'get', 'update', 'delete'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    name: { type: 'string', description: 'Name for new worksheet (add action)', optional: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID (get/update/delete actions)', optional: true },
                    properties: { type: 'object', description: 'Properties to update: { name?, position?, visibility? } (update action)', optional: true }
                };
                break;
            case 'excelRange':
                toolDef.description = 'Read, write, format, sort, and merge cell ranges in an Excel workbook.\n\nActions:\n- get: Read cell values, formulas, formatting (requires: fileId, sheetIdOrName, address)\n- update: Write values to cells (requires: fileId, sheetIdOrName, address, values as 2D array)\n- getFormat: Get formatting properties (requires: fileId, sheetIdOrName, address)\n- updateFormat: Set font/fill/borders (requires: fileId, sheetIdOrName, address, format)\n- sort: Sort range (requires: fileId, sheetIdOrName, address, fields array)\n- merge: Merge cells (requires: fileId, sheetIdOrName, address; optional: across)\n- unmerge: Unmerge cells (requires: fileId, sheetIdOrName, address)\n\nMax recommended: 10,000 cells per request. Range addresses use Excel notation: A1:C4, Sheet1!B2:D10.';
                toolDef.endpoint = '/api/v1/excel/range/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['get', 'update', 'getFormat', 'updateFormat', 'sort', 'merge', 'unmerge'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range in Excel notation (e.g., A1:C4)', required: true },
                    values: { type: 'array', description: '2D array of values for update action, e.g., [["Name","Age"],["Alice",30]]', optional: true },
                    format: { type: 'object', description: 'Format properties for updateFormat: { font?, fill?, borders?, horizontalAlignment?, numberFormat? }', optional: true },
                    fields: { type: 'array', description: 'Sort fields for sort action: [{ key: columnIndex, ascending: true/false }]', optional: true },
                    across: { type: 'boolean', description: 'Merge cells in each row separately (merge action)', optional: true }
                };
                break;
            case 'excelTable':
                toolDef.description = 'Manage tables, rows, columns, sorting, and filtering in an Excel workbook.\n\nActions:\n- list: List all tables (requires: fileId; optional: sheetIdOrName)\n- create: Create table from range (requires: fileId, sheetIdOrName, address, hasHeaders)\n- update: Update table properties (requires: fileId, tableIdOrName, properties)\n- delete: Delete a table (requires: fileId, tableIdOrName)\n- listRows: List table rows (requires: fileId, tableIdOrName)\n- addRow: Add a row (requires: fileId, tableIdOrName, values; optional: index)\n- deleteRow: Delete a row (requires: fileId, tableIdOrName, index)\n- listColumns: List table columns (requires: fileId, tableIdOrName)\n- addColumn: Add a column (requires: fileId, tableIdOrName, values; optional: index)\n- deleteColumn: Delete a column (requires: fileId, tableIdOrName, columnIdOrName)\n- sort: Sort table (requires: fileId, tableIdOrName, fields)\n- filter: Apply filter (requires: fileId, tableIdOrName, columnId, criteria)\n- clearFilter: Clear column filter (requires: fileId, tableIdOrName, columnId)\n- convertToRange: Convert table to range (requires: fileId, tableIdOrName)';
                toolDef.endpoint = '/api/v1/excel/table/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['list', 'create', 'update', 'delete', 'listRows', 'addRow', 'deleteRow', 'listColumns', 'addColumn', 'deleteColumn', 'sort', 'filter', 'clearFilter', 'convertToRange'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID (list/create actions)', optional: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', optional: true },
                    address: { type: 'string', description: 'Range address for create action (e.g., A1:D5)', optional: true },
                    hasHeaders: { type: 'boolean', description: 'Whether first row contains headers (create action)', optional: true },
                    values: { type: 'array', description: 'Values for addRow (1D array) or addColumn (2D array)', optional: true },
                    index: { type: 'number', description: 'Position index for addRow/addColumn, or row index for deleteRow', optional: true },
                    columnIdOrName: { type: 'string', description: 'Column name or ID (deleteColumn action)', optional: true },
                    columnId: { type: 'string', description: 'Column ID for filter/clearFilter actions', optional: true },
                    criteria: { type: 'object', description: 'Filter criteria: { filterOn, criterion1, operator?, criterion2? }', optional: true },
                    fields: { type: 'array', description: 'Sort fields: [{ key: columnIndex, ascending: true/false }]', optional: true },
                    properties: { type: 'object', description: 'Properties for update: { name?, style?, showHeaders?, showTotals? }', optional: true }
                };
                break;
            case 'excelFunction':
                toolDef.description = 'Call Excel workbook functions or recalculate formulas.\n\nActions:\n- call: Call any Excel function like SUM, VLOOKUP, PMT, MEDIAN, etc. (requires: fileId, functionName, args). Supports 300+ functions. Range arguments use { address: "Sheet1!A1:B5" } format.\n- calculate: Recalculate all formulas in the workbook (requires: fileId; optional: calculationType)';
                toolDef.endpoint = '/api/v1/excel/function/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['call', 'calculate'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    functionName: { type: 'string', description: 'Excel function name for call action (e.g., sum, vlookup)', optional: true },
                    args: { type: 'object', description: 'Function arguments for call action. For ranges use { address: "Sheet1!A1:B5" }', optional: true },
                    calculationType: { type: 'string', description: 'Calculation type for calculate action', optional: true, enum: ['Recalculate', 'Full', 'FullRebuild'], default: 'Full' }
                };
                break;
            case 'wordDocument':
                toolDef.description = 'Create, read, and convert Word documents.\n\nActions:\n- create: Create a new .docx from structured content (requires: fileName, content with sections array). Section types: heading, paragraph, table, list, image.\n- read: Read document as HTML and plain text (requires: fileId). Max 25MB.\n- metadata: Get document metadata — title, author, dates (requires: fileId)\n- html: Convert document to HTML for preview (requires: fileId). Max 25MB.\n- pdf: Convert document to PDF via Microsoft Graph (requires: fileId)';
                toolDef.endpoint = '/api/v1/word/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['create', 'read', 'metadata', 'html', 'pdf'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .docx file (read/metadata/html/pdf actions)', optional: true },
                    fileName: { type: 'string', description: 'Name for new document, e.g., "Report.docx" (create action)', optional: true },
                    content: { type: 'object', description: 'Document content for create: { sections: [{ type, ... }] }', optional: true }
                };
                break;
            case 'powerpointPresentation':
                toolDef.description = 'Create, read, and convert PowerPoint presentations.\n\nActions:\n- create: Create a new .pptx from structured slides (requires: fileName, slides array). Slide layouts: title, content, blank.\n- read: Read slide content as structured text (requires: fileId). Max 25MB.\n- metadata: Get presentation metadata — title, author, slide count (requires: fileId)\n- pdf: Convert presentation to PDF via Microsoft Graph (requires: fileId)';
                toolDef.endpoint = '/api/v1/powerpoint/action';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    action: { type: 'string', description: 'Action to perform', required: true, enum: ['create', 'read', 'metadata', 'pdf'] },
                    fileId: { type: 'string', description: 'Drive item ID of the .pptx file (read/metadata/pdf actions)', optional: true },
                    fileName: { type: 'string', description: 'Name for new presentation, e.g., "Deck.pptx" (create action)', optional: true },
                    slides: { type: 'array', description: 'Slides for create: [{ layout, title?, subtitle?, body?: [{type,...}] }]', optional: true }
                };
                break;

            // ========== Legacy granular tools (kept for REST API backward compat) ==========
            case 'createWorkbookSession':
                toolDef.description = 'Create a workbook session for an Excel file. Sessions enable efficient batch operations. Only .xlsx files are supported.';
                toolDef.endpoint = '/api/v1/excel/session';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'OneDrive/SharePoint drive item ID of the .xlsx file', required: true },
                    persistent: { type: 'boolean', description: 'Whether changes should be saved (true) or temporary (false)', optional: true, default: true }
                };
                break;
            case 'closeWorkbookSession':
                toolDef.description = 'Close an active workbook session for an Excel file.';
                toolDef.endpoint = '/api/v1/excel/session';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'OneDrive/SharePoint drive item ID of the .xlsx file', required: true }
                };
                break;
            case 'listWorksheets':
                toolDef.description = 'List all worksheets in an Excel workbook.';
                toolDef.endpoint = '/api/v1/excel/worksheets';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'OneDrive/SharePoint drive item ID of the .xlsx file', required: true }
                };
                break;
            case 'addWorksheet':
                toolDef.description = 'Add a new worksheet to an Excel workbook.';
                toolDef.endpoint = '/api/v1/excel/worksheets';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'OneDrive/SharePoint drive item ID of the .xlsx file', required: true },
                    name: { type: 'string', description: 'Name for the new worksheet', required: true }
                };
                break;
            case 'getWorksheet':
                toolDef.description = 'Get a specific worksheet by name or ID.';
                toolDef.endpoint = '/api/v1/excel/worksheets/detail';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true }
                };
                break;
            case 'updateWorksheet':
                toolDef.description = 'Update worksheet properties (name, position, visibility).';
                toolDef.endpoint = '/api/v1/excel/worksheets/update';
                toolDef.method = 'PATCH';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    properties: { type: 'object', description: 'Properties to update: { name?, position?, visibility? }', required: true }
                };
                break;
            case 'deleteWorksheet':
                toolDef.description = 'Delete a worksheet from an Excel workbook.';
                toolDef.endpoint = '/api/v1/excel/worksheets/delete';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID to delete', required: true }
                };
                break;
            case 'getRange':
                toolDef.description = 'Get cell values, formulas, and formatting from a range in an Excel worksheet. Max recommended: 10,000 cells per request.';
                toolDef.endpoint = '/api/v1/excel/range';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range in Excel notation (e.g., A1:C4, Sheet1!B2:D10)', required: true }
                };
                break;
            case 'updateRange':
                toolDef.description = 'Write values to a range in an Excel worksheet. Values should be a 2D array matching the range dimensions.';
                toolDef.endpoint = '/api/v1/excel/range';
                toolDef.method = 'PATCH';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range in Excel notation (e.g., A1:C4)', required: true },
                    values: { type: 'array', description: '2D array of values matching the range dimensions, e.g., [["Name","Age"],["Alice",30]]', required: true }
                };
                break;
            case 'getRangeFormat':
                toolDef.description = 'Get formatting properties of a range (font, fill, borders, alignment).';
                toolDef.endpoint = '/api/v1/excel/range/format';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range in Excel notation', required: true }
                };
                break;
            case 'updateRangeFormat':
                toolDef.description = 'Update formatting of a range (font, fill, borders, number format).';
                toolDef.endpoint = '/api/v1/excel/range/format';
                toolDef.method = 'PATCH';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range in Excel notation', required: true },
                    format: { type: 'object', description: 'Format properties: { font?, fill?, borders?, horizontalAlignment?, numberFormat? }', required: true }
                };
                break;
            case 'sortRange':
                toolDef.description = 'Sort a range of cells in an Excel worksheet.';
                toolDef.endpoint = '/api/v1/excel/range/sort';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range to sort', required: true },
                    fields: { type: 'array', description: 'Sort fields: [{ key: columnIndex, ascending: true/false }]', required: true }
                };
                break;
            case 'mergeRange':
                toolDef.description = 'Merge cells in a range.';
                toolDef.endpoint = '/api/v1/excel/range/merge';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range to merge', required: true },
                    across: { type: 'boolean', description: 'Merge cells in each row separately', optional: true, default: false }
                };
                break;
            case 'unmergeRange':
                toolDef.description = 'Unmerge previously merged cells in a range.';
                toolDef.endpoint = '/api/v1/excel/range/unmerge';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Cell range to unmerge', required: true }
                };
                break;
            case 'listTables':
                toolDef.description = 'List all tables in an Excel worksheet.';
                toolDef.endpoint = '/api/v1/excel/tables';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true }
                };
                break;
            case 'createTable':
                toolDef.description = 'Create a new table from a range in an Excel worksheet.';
                toolDef.endpoint = '/api/v1/excel/tables';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    sheetIdOrName: { type: 'string', description: 'Worksheet name or ID', required: true },
                    address: { type: 'string', description: 'Range address for the table (e.g., A1:D5)', required: true },
                    hasHeaders: { type: 'boolean', description: 'Whether the first row contains headers', required: true }
                };
                break;
            case 'updateTable':
                toolDef.description = 'Update table properties (name, style, showHeaders, showTotals).';
                toolDef.endpoint = '/api/v1/excel/tables/update';
                toolDef.method = 'PATCH';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    properties: { type: 'object', description: 'Properties: { name?, style?, showHeaders?, showTotals? }', required: true }
                };
                break;
            case 'deleteTable':
                toolDef.description = 'Delete a table from an Excel workbook. The data remains but table formatting is removed.';
                toolDef.endpoint = '/api/v1/excel/tables/delete';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID to delete', required: true }
                };
                break;
            case 'listTableRows':
                toolDef.description = 'List all rows in an Excel table.';
                toolDef.endpoint = '/api/v1/excel/tables/rows';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true }
                };
                break;
            case 'addTableRow':
                toolDef.description = 'Add a row to an Excel table.';
                toolDef.endpoint = '/api/v1/excel/tables/rows';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    values: { type: 'array', description: 'Row values as array (e.g., ["Alice", 30, "Engineer"])', required: true },
                    index: { type: 'number', description: 'Position to insert the row (null for end)', optional: true }
                };
                break;
            case 'deleteTableRow':
                toolDef.description = 'Delete a row from an Excel table by its index.';
                toolDef.endpoint = '/api/v1/excel/tables/rows/delete';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    index: { type: 'number', description: 'Zero-based row index to delete', required: true }
                };
                break;
            case 'listTableColumns':
                toolDef.description = 'List all columns in an Excel table.';
                toolDef.endpoint = '/api/v1/excel/tables/columns';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true }
                };
                break;
            case 'addTableColumn':
                toolDef.description = 'Add a column to an Excel table.';
                toolDef.endpoint = '/api/v1/excel/tables/columns';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    values: { type: 'array', description: 'Column values including header (e.g., [["Status"],["Open"],["Closed"]])', required: true },
                    index: { type: 'number', description: 'Column position index', optional: true }
                };
                break;
            case 'deleteTableColumn':
                toolDef.description = 'Delete a column from an Excel table.';
                toolDef.endpoint = '/api/v1/excel/tables/columns/delete';
                toolDef.method = 'DELETE';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    columnIdOrName: { type: 'string', description: 'Column name or ID to delete', required: true }
                };
                break;
            case 'sortTable':
                toolDef.description = 'Sort an Excel table by one or more columns.';
                toolDef.endpoint = '/api/v1/excel/tables/sort';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    fields: { type: 'array', description: 'Sort fields: [{ key: columnIndex, ascending: true/false }]', required: true }
                };
                break;
            case 'filterTable':
                toolDef.description = 'Apply a filter to a table column.';
                toolDef.endpoint = '/api/v1/excel/tables/filter';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    columnId: { type: 'string', description: 'Column ID to filter', required: true },
                    criteria: { type: 'object', description: 'Filter criteria: { filterOn, criterion1, operator?, criterion2? }', required: true }
                };
                break;
            case 'clearTableFilter':
                toolDef.description = 'Clear the filter on a table column.';
                toolDef.endpoint = '/api/v1/excel/tables/filter/clear';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID', required: true },
                    columnId: { type: 'string', description: 'Column ID to clear filter from', required: true }
                };
                break;
            case 'convertTableToRange':
                toolDef.description = 'Convert an Excel table to a regular range. Removes table formatting but keeps data.';
                toolDef.endpoint = '/api/v1/excel/tables/convert';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    tableIdOrName: { type: 'string', description: 'Table name or ID to convert', required: true }
                };
                break;
            case 'callWorkbookFunction':
                toolDef.description = 'Call any Excel workbook function (SUM, VLOOKUP, PMT, MEDIAN, etc.). Supports 300+ functions. Range arguments use { address: "Sheet1!A1:B5" } format.';
                toolDef.endpoint = '/api/v1/excel/functions';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    functionName: { type: 'string', description: 'Excel function name (e.g., sum, vlookup, pmt, median)', required: true },
                    args: { type: 'object', description: 'Function arguments as key-value pairs. For ranges use { address: "Sheet1!A1:B5" }', required: true }
                };
                break;
            case 'calculateWorkbook':
                toolDef.description = 'Recalculate all formulas in the workbook.';
                toolDef.endpoint = '/api/v1/excel/calculate';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .xlsx file', required: true },
                    calculationType: { type: 'string', description: 'Calculation type', optional: true, enum: ['Recalculate', 'Full', 'FullRebuild'], default: 'Full' }
                };
                break;

            // ========== Word Document Tools ==========
            case 'createWordDocument':
                toolDef.description = 'Create a new Word document (.docx) from structured content. Supports headings, paragraphs, tables, lists, and images.';
                toolDef.endpoint = '/api/v1/word/create';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileName: { type: 'string', description: 'Name for the document (e.g., "Report.docx"). Uploaded to OneDrive root.', required: true },
                    content: { type: 'object', description: 'Document content: { sections: [{ type: "heading"|"paragraph"|"table"|"list"|"image", ... }] }', required: true }
                };
                break;
            case 'readWordDocument':
                toolDef.description = 'Read a Word document and return its content as structured HTML and plain text. Max file size: 25MB.';
                toolDef.endpoint = '/api/v1/word/read';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .docx file', required: true }
                };
                break;
            case 'convertDocumentToPdf':
                toolDef.description = 'Convert a Word document to PDF using Microsoft Graph. Returns the PDF content.';
                toolDef.endpoint = '/api/v1/word/pdf';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .docx file', required: true }
                };
                break;
            case 'getWordDocumentMetadata':
                toolDef.description = 'Get metadata from a Word document (title, author, created date, modified date, keywords).';
                toolDef.endpoint = '/api/v1/word/metadata';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .docx file', required: true }
                };
                break;
            case 'getWordDocumentAsHtml':
                toolDef.description = 'Convert a Word document to HTML for preview or display. Max file size: 25MB.';
                toolDef.endpoint = '/api/v1/word/html';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .docx file', required: true }
                };
                break;

            // ========== PowerPoint Tools ==========
            case 'createPresentation':
                toolDef.description = 'Create a new PowerPoint presentation (.pptx) from structured slide data. Supports title slides, content slides with text and images.';
                toolDef.endpoint = '/api/v1/powerpoint/create';
                toolDef.method = 'POST';
                toolDef.parameters = {
                    fileName: { type: 'string', description: 'Name for the presentation (e.g., "Deck.pptx"). Uploaded to OneDrive root.', required: true },
                    slides: { type: 'array', description: 'Array of slides: [{ layout: "title"|"content"|"blank", title?, subtitle?, body?: [{type:"text"|"image",...}] }]', required: true }
                };
                break;
            case 'readPresentation':
                toolDef.description = 'Read a PowerPoint presentation and return structured slide content (text elements per slide). Max file size: 25MB.';
                toolDef.endpoint = '/api/v1/powerpoint/read';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .pptx file', required: true }
                };
                break;
            case 'convertPresentationToPdf':
                toolDef.description = 'Convert a PowerPoint presentation to PDF using Microsoft Graph.';
                toolDef.endpoint = '/api/v1/powerpoint/pdf';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .pptx file', required: true }
                };
                break;
            case 'getPresentationMetadata':
                toolDef.description = 'Get metadata from a PowerPoint presentation (title, author, slide count, created/modified dates).';
                toolDef.endpoint = '/api/v1/powerpoint/metadata';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    fileId: { type: 'string', description: 'Drive item ID of the .pptx file', required: true }
                };
                break;

            // Default for unknown capabilities
            default:
                MonitoringService.warn(`No specific definition found for capability '${capability}' in module '${moduleName}'. Using defaults.`, {
                    capability,
                    moduleName,
                    timestamp: new Date().toISOString()
                }, 'tools');
                // Use defaults with generic parameters
                break;
        }

        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Tool definition generated successfully', {
                moduleName,
                capability,
                toolName: toolDef.name,
                endpoint: toolDef.endpoint,
                method: toolDef.method,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Tool definition generated with session', {
                sessionId,
                moduleName,
                capability,
                toolName: toolDef.name,
                endpoint: toolDef.endpoint,
                method: toolDef.method,
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'tools');
        }

        return toolDef;
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'tools',
            `Failed to generate tool definition: ${error.message}`,
            'error',
            {
                moduleName,
                capability,
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
            MonitoringService.error('Tool definition generation failed', {
                moduleName,
                capability,
                error: error.message,
                duration: executionTime,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Tool definition generation failed', {
                sessionId,
                moduleName,
                capability,
                error: error.message,
                duration: executionTime,
                timestamp: new Date().toISOString()
            }, 'tools');
        }
        
        throw mcpError;
    }
}

    /**
     * Invalidates any internal caches, forcing regeneration on next access.
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [sessionId] - Session ID for context
     */
    function refresh(userId = null, sessionId = null) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Refreshing tools service cache', {
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        }
        
        try {
            const previousCacheSize = cachedTools ? cachedTools.length : 0;
            
            cachedTools = null; // Clear cache
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Tools service cache refreshed successfully', {
                    previousCacheSize,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Tools service cache refreshed with session', {
                    sessionId,
                    previousCacheSize,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            MonitoringService.trackMetric('tools_refresh_success', executionTime, {
                previousCacheSize,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, userId);
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'tools',
                `Tools service refresh failed: ${error.message}`,
                'error',
                {
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
                MonitoringService.error('Tools service refresh failed', {
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Tools service refresh failed', {
                    sessionId,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            MonitoringService.trackMetric('tools_refresh_failure', executionTime, {
                errorType: error.code || 'unknown',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, userId);
            
            throw mcpError;
        }
    }

    /**
     * Helper function to transform attendees from string or array to proper format
     * @param {string|Array} attendees - Attendees as string or array
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [sessionId] - Session ID for context
     * @returns {Array|undefined} - Transformed attendees or undefined if none
     */
    function transformAttendees(attendees, userId = null, sessionId = null) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Transforming attendees', {
                attendeesType: typeof attendees,
                attendeesLength: Array.isArray(attendees) ? attendees.length : (attendees ? attendees.length : 0),
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        }
        
        try {
            if (!attendees) {
                // Pattern 2: User Activity Logs (for empty attendees)
                if (userId) {
                    MonitoringService.info('Attendees transformation completed - no attendees provided', {
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Attendees transformation completed with session - no attendees', {
                        sessionId,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                return undefined;
            }
            
            let result;
            
            // If attendees is a string (comma-separated), convert to array
            if (typeof attendees === 'string') {
                result = attendees.split(',').map(email => email.trim());
            } else {
                // If already an array, return as is
                result = attendees;
            }
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Attendees transformation completed successfully', {
                    originalType: typeof attendees,
                    resultCount: result.length,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Attendees transformation completed with session', {
                    sessionId,
                    originalType: typeof attendees,
                    resultCount: result.length,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            return result;
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'tools',
                `Failed to transform attendees: ${error.message}`,
                'error',
                {
                    attendeesType: typeof attendees,
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
                MonitoringService.error('Attendees transformation failed', {
                    attendeesType: typeof attendees,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Attendees transformation failed', {
                    sessionId,
                    attendeesType: typeof attendees,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            throw mcpError;
        }
    }
    
    /**
     * Helper function to transform date/time to proper format
     * @param {string|object} dateTime - Date time string or object
     * @param {string} timeZone - Default timezone if not specified
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [sessionId] - Session ID for context
     * @returns {object|undefined} - Transformed date time object
     */
    function transformDateTime(dateTime, timeZone = 'UTC', userId = null, sessionId = null) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Transforming date/time', {
                dateTimeType: typeof dateTime,
                timeZone,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        }
        
        try {
            if (!dateTime) {
                // Pattern 2: User Activity Logs (for empty dateTime)
                if (userId) {
                    MonitoringService.info('DateTime transformation completed - no dateTime provided', {
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('DateTime transformation completed with session - no dateTime', {
                        sessionId,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                return undefined;
            }
            
            let result;
            
            // If already an object with dateTime, return as is
            if (typeof dateTime === 'object' && dateTime.dateTime) {
                result = dateTime;
            } else if (typeof dateTime === 'string') {
                // If a string, convert to object format
                result = {
                    dateTime: dateTime,
                    timeZone: timeZone
                };
            } else {
                result = dateTime;
            }
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('DateTime transformation completed successfully', {
                    originalType: typeof dateTime,
                    resultTimeZone: result.timeZone || timeZone,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.info('DateTime transformation completed with session', {
                    sessionId,
                    originalType: typeof dateTime,
                    resultTimeZone: result.timeZone || timeZone,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            return result;
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'tools',
                `Failed to transform dateTime: ${error.message}`,
                'error',
                {
                    dateTimeType: typeof dateTime,
                    timeZone,
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
                MonitoringService.error('DateTime transformation failed', {
                    dateTimeType: typeof dateTime,
                    timeZone,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('DateTime transformation failed', {
                    sessionId,
                    dateTimeType: typeof dateTime,
                    timeZone,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            throw mcpError;
        }
    }
    
    /**
     * Transforms parameters for a specific module and method with user context
     * @param {string} moduleName - Module name
     * @param {string} methodName - Method name
     * @param {object} params - Original parameters
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [deviceId] - Device ID for multi-user context
     * @param {string} [sessionId] - Session ID for context
     * @returns {object} - Transformed parameters
     */
    function transformParameters(moduleName, methodName, params, userId = null, deviceId = null, sessionId = null) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Transforming parameters', {
                moduleName,
                methodName,
                paramKeys: Object.keys(params),
                userId,
                deviceId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId);
        }
        
        try {
        
        // Create a copy of the parameters to avoid modifying the original
        const transformedParams = { ...params };
        
        // Add user context to all transformed parameters for isolation
        if (userId) {
            transformedParams._userId = userId;
        }
        if (deviceId) {
            transformedParams._deviceId = deviceId;
        }
        
        // Transform parameters based on the module and method
        switch (`${moduleName}.${methodName}`) {
            // Mail module methods
            case 'mail.sendEmail':
            case 'mail.sendMail':
                return {
                    to: transformAttendees(transformedParams.to),
                    subject: transformedParams.subject,
                    body: transformedParams.body,
                    cc: transformAttendees(transformedParams.cc),
                    bcc: transformAttendees(transformedParams.bcc)
                };
                
            case 'mail.searchEmails':
            case 'mail.searchMail':
                // Ensure query parameter is properly named
                if (transformedParams.query && !transformedParams.q) {
                    transformedParams.q = transformedParams.query;
                    delete transformedParams.query;
                }
                return transformedParams;
                
            // Calendar module methods
            case 'calendar.create':
            case 'calendar.createEvent':
                return {
                    subject: transformedParams.subject,
                    start: transformDateTime(transformedParams.start, transformedParams.timeZone),
                    end: transformDateTime(transformedParams.end, transformedParams.timeZone),
                    location: transformedParams.location,
                    body: transformedParams.body,
                    attendees: transformAttendees(transformedParams.attendees),
                    isOnlineMeeting: transformedParams.isOnlineMeeting
                };
                
            case 'calendar.update':
            case 'calendar.updateEvent':
                // Create an object with only the provided parameters
                const updateData = {};
                
                if (transformedParams.id !== undefined) {
                    updateData.id = transformedParams.id;
                }
                
                if (transformedParams.subject !== undefined) {
                    updateData.subject = transformedParams.subject;
                }
                
                if (transformedParams.start !== undefined) {
                    updateData.start = transformDateTime(transformedParams.start, transformedParams.timeZone);
                }
                
                if (transformedParams.end !== undefined) {
                    updateData.end = transformDateTime(transformedParams.end, transformedParams.timeZone);
                }
                
                if (transformedParams.attendees !== undefined) {
                    updateData.attendees = transformAttendees(transformedParams.attendees);
                }
                
                if (transformedParams.location !== undefined) {
                    updateData.location = transformedParams.location;
                }
                
                if (transformedParams.body !== undefined) {
                    updateData.body = transformedParams.body;
                }
                
                if (transformedParams.isOnlineMeeting !== undefined) {
                    updateData.isOnlineMeeting = transformedParams.isOnlineMeeting;
                }
                
                return updateData;
                
            case 'calendar.getAvailability':
                logger.debug(`transformParameters: Processing getAvailability parameters`, JSON.stringify(transformedParams, null, 2));
                
                // Check if we're receiving the new format with timeSlots array
                if (transformedParams.timeSlots && Array.isArray(transformedParams.timeSlots)) {
                    logger.debug(`transformParameters: getAvailability received timeSlots format with ${transformedParams.timeSlots.length} slots`);
                    
                    // Validate the structure of each time slot
                    const timeSlots = transformedParams.timeSlots.map((slot, index) => {
                        logger.debug(`transformParameters: Processing time slot ${index}:`, JSON.stringify(slot, null, 2));
                        
                        // Handle different possible formats for start/end
                        // Case 1: slot has start/end objects with dateTime property
                        if (slot.start?.dateTime && slot.end?.dateTime) {
                            return {
                                start: {
                                    dateTime: slot.start.dateTime,
                                    timeZone: slot.start.timeZone || transformedParams.timeZone || 'UTC'
                                },
                                end: {
                                    dateTime: slot.end.dateTime,
                                    timeZone: slot.end.timeZone || transformedParams.timeZone || 'UTC'
                                }
                            };
                        }
                        
                        // Case 2: slot has start/end as simple strings
                        if (typeof slot.start === 'string' && typeof slot.end === 'string') {
                            return {
                                start: {
                                    dateTime: slot.start,
                                    timeZone: transformedParams.timeZone || 'UTC'
                                },
                                end: {
                                    dateTime: slot.end,
                                    timeZone: transformedParams.timeZone || 'UTC'
                                }
                            };
                        }
                        
                        // Case 3: slot itself is malformed, try to extract what we can
                        logger.warn(`transformParameters: Malformed time slot at index ${index}:`, JSON.stringify(slot, null, 2));
                        return {
                            start: {
                                dateTime: slot.start?.dateTime || slot.start || new Date().toISOString(),
                                timeZone: slot.start?.timeZone || transformedParams.timeZone || 'UTC'
                            },
                            end: {
                                dateTime: slot.end?.dateTime || slot.end || new Date(Date.now() + 3600000).toISOString(),
                                timeZone: slot.end?.timeZone || transformedParams.timeZone || 'UTC'
                            }
                        };
                    });
                    
                    // Transform parameters to match the controller's expectations
                    return {
                        users: transformAttendees(transformedParams.users) || [],
                        timeSlots: timeSlots
                    };
                } else {
                    // Original format with start/end fields
                    // Ensure we have start/end times for availability check
                    if (!transformedParams.start && !transformedParams.startTime) {
                        logger.warn('getAvailability requires start time');
                        throw new Error('Start time is required for getAvailability');
                    }
                    
                    if (!transformedParams.end && !transformedParams.endTime) {
                        logger.warn('getAvailability requires end time');
                        throw new Error('End time is required for getAvailability');
                    }
                    
                    // Transform start/end to the format expected by the API
                    const availStartTime = typeof transformedParams.start === 'object' && transformedParams.start.dateTime 
                        ? transformedParams.start.dateTime 
                        : transformedParams.start || transformedParams.startTime;
                    
                    const availEndTime = typeof transformedParams.end === 'object' && transformedParams.end.dateTime 
                        ? transformedParams.end.dateTime 
                        : transformedParams.end || transformedParams.endTime;
                    
                    logger.debug(`transformParameters: Extracted start/end times:`, { availStartTime, availEndTime });
                    
                    return {
                        users: transformAttendees(transformedParams.users) || transformAttendees(transformedParams.attendees) || [],
                        timeSlots: [
                            {
                                start: {
                                    dateTime: availStartTime,
                                    timeZone: transformedParams.timeZone || 'UTC'
                                },
                                end: {
                                    dateTime: availEndTime,
                                    timeZone: transformedParams.timeZone || 'UTC'
                                }
                            }
                        ]
                    };
                }
                
                
            case 'calendar.findMeetingTimes':
                // Extract time constraints from different possible input formats
                let timeConstraints = transformedParams.timeConstraints;
                if (!timeConstraints && (transformedParams.startTime || transformedParams.start)) {
                    timeConstraints = {
                        startTime: transformedParams.startTime || transformedParams.start,
                        endTime: transformedParams.endTime || transformedParams.end,
                        timeZone: transformedParams.timeZone || 'UTC'
                    };
                }
                
                return {
                    attendees: transformAttendees(transformedParams.attendees) || [],
                    timeConstraint: {
                        start: timeConstraints?.startTime || timeConstraints?.start,
                        end: timeConstraints?.endTime || timeConstraints?.end,
                        timeZone: timeConstraints?.timeZone || transformedParams.timeZone || 'UTC'
                    },
                    meetingDuration: transformedParams.meetingDuration || transformedParams.duration || 60,
                    maxCandidates: transformedParams.maxCandidates || 10,
                    minimumAttendeePercentage: transformedParams.minimumAttendeePercentage || 100
                };

            // People module methods
            case 'people.find':
            case 'people.findPeople':
                // Make sure query parameter is preserved
                if (!transformedParams.query && transformedParams.q) {
                    transformedParams.query = transformedParams.q;
                    delete transformedParams.q;
                }
                
                // Ensure limit is a number
                if (transformedParams.limit) {
                    transformedParams.limit = parseInt(transformedParams.limit, 10);
                }
                
                return transformedParams;
                
            // Query module methods
            case 'query.processQuery':
                return { 
                    query: transformedParams.query,
                    context: transformedParams.context
                };
                
            // Default case - return original parameters
            default:
                const executionTime = Date.now() - startTime;
                MonitoringService.trackMetric('tools_transform_params_success', executionTime, {
                    moduleName,
                    methodName,
                    hasTransform: false,
                    userId,
                    deviceId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Parameter transformation completed successfully', {
                        moduleName,
                        methodName,
                        paramCount: Object.keys(transformedParams).length,
                        hasTransform: false,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Parameter transformation completed with session', {
                        sessionId,
                        moduleName,
                        methodName,
                        paramCount: Object.keys(transformedParams).length,
                        hasTransform: false,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                return transformedParams;
        }
        
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'tools',
                `Parameter transformation failed: ${error.message}`,
                'error',
                {
                    moduleName,
                    methodName,
                    paramKeys: Object.keys(params),
                    error: error.message,
                    stack: error.stack,
                    userId,
                    deviceId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Parameter transformation failed', {
                    moduleName,
                    methodName,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Parameter transformation failed', {
                    sessionId,
                    moduleName,
                    methodName,
                    error: error.message,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            MonitoringService.trackMetric('tools_transform_params_failure', executionTime, {
                moduleName,
                methodName,
                errorType: error.code || 'unknown',
                userId,
                deviceId,
                sessionId,
                timestamp: new Date().toISOString()
            }, userId);
            
            throw mcpError;
        }
    }

    return {
        /**
         * Gets all available tools from registered modules
         * @param {string} [userId] - User ID for multi-user context
         * @param {string} [sessionId] - Session ID for context
         * @returns {Array<object>} Tool definitions
         */
        getAllTools(userId = null, sessionId = null) {
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting all tools', {
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            }
            
            try {
                // Check cache first
                if (cachedTools) {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_get_all_cache_hit', executionTime, {
                        toolCount: cachedTools.length,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, userId);
                    
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Returning cached tool definitions', {
                            toolCount: cachedTools.length,
                            userId,
                            sessionId,
                            timestamp: new Date().toISOString()
                        }, 'tools', null, userId);
                    }
                    
                    // Pattern 2: User Activity Logs (cache hit)
                    if (userId) {
                        MonitoringService.info('All tools retrieved from cache successfully', {
                            toolCount: cachedTools.length,
                            duration: executionTime,
                            timestamp: new Date().toISOString()
                        }, 'tools', null, userId);
                    } else if (sessionId) {
                        MonitoringService.info('All tools retrieved from cache with session', {
                            sessionId,
                            toolCount: cachedTools.length,
                            duration: executionTime,
                            timestamp: new Date().toISOString()
                        }, 'tools');
                    }
                    
                    return cachedTools;
                }

                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Cache miss, generating tool definitions', {
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                }
                const tools = [];
            const modules = moduleRegistry.getAllModules();
            
            // First, add the findPeople tool at the beginning of the list
            // This is critical because person resolution must happen before scheduling or sending invites
            const peopleModule = modules.find(m => m.name === 'people' || m.capabilities?.includes('findPeople'));
            if (peopleModule) {
                const findPeopleTool = {
                    name: 'findPeople',
                    description: 'Find people by name or email. Use this to get email addresses for scheduling or sending emails.',
                    endpoint: '/api/v1/people/find',
                    method: 'GET',
                    parameters: {
                        query: { type: 'string', description: 'Search query to find a person' },
                        name: { type: 'string', description: 'Person name to search for', optional: true },
                        limit: { type: 'number', description: 'Maximum number of results', optional: true }
                    }
                };
                tools.push(findPeopleTool);
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Added findPeople tool with high priority', {
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                }
            }

            // For each module, generate tool definitions for each capability (except findPeople which we already added)
            for (const module of modules) {
                if (Array.isArray(module.capabilities)) {
                    for (const capability of module.capabilities) {
                        // Skip findPeople since we already added it
                        if (capability === 'findPeople') continue;
                        tools.push(generateToolDefinition(module.name, capability, userId, sessionId));
                    }
                }
            }

            // Add query tool (special case)
            tools.push({
                name: 'query',
                description: 'Submit a natural language query to Microsoft 365',
                endpoint: '/api/v1/query',
                method: 'POST',
                parameters: {
                    query: { type: 'string', description: 'The user\'s natural language question' },
                    context: { type: 'object', description: 'Conversation context', optional: true }
                }
            });

            // Store in cache before returning
            cachedTools = tools;
            const executionTime = Date.now() - startTime;
            
            MonitoringService.trackMetric('tools_get_all_cache_miss', executionTime, {
                toolCount: tools.length,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, userId);
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('All tools generated and cached successfully', {
                    toolCount: tools.length,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.info('All tools generated and cached with session', {
                    sessionId,
                    toolCount: tools.length,
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }

            return tools;
            
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'tools',
                    `Failed to get all tools: ${error.message}`,
                    'error',
                    {
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
                    MonitoringService.error('Failed to get all tools', {
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get all tools', {
                        sessionId,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_get_all_failure', executionTime, {
                    errorType: error.code || 'unknown',
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                throw mcpError;
            }
        },

        /**
         * Gets a tool definition by name
         * @param {string} toolName - Name of the tool
         * @param {string} [userId] - User ID for multi-user context
         * @param {string} [sessionId] - Session ID for context
         * @returns {object|null} Tool definition or null if not found
         */
        getToolByName(toolName, userId = null, sessionId = null) {
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting tool by name', {
                    toolName,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            }
            
            try {
                const allTools = this.getAllTools(userId, sessionId); // Uses cache if available
                const lowerCaseToolName = toolName.toLowerCase();
                const foundTool = allTools.find(tool => tool.name.toLowerCase() === lowerCaseToolName);
                
                const executionTime = Date.now() - startTime;
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Tool retrieval by name completed successfully', {
                        toolName,
                        found: !!foundTool,
                        totalTools: allTools.length,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Tool retrieval by name completed with session', {
                        sessionId,
                        toolName,
                        found: !!foundTool,
                        totalTools: allTools.length,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_get_by_name_success', executionTime, {
                    toolName,
                    found: !!foundTool,
                    totalTools: allTools.length,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                return foundTool || null;
                
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'tools',
                    `Failed to get tool by name: ${error.message}`,
                    'error',
                    {
                        toolName,
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
                    MonitoringService.error('Failed to get tool by name', {
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get tool by name', {
                        sessionId,
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_get_by_name_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                throw mcpError;
            }
        },

        /**
         * Maps a tool name to a module and method
         * @param {string} toolName - Name of the tool
         * @param {string} [userId] - User ID for multi-user context
         * @param {string} [sessionId] - Session ID for context
         * @returns {object|null} Module and method mapping or null if not found
         */
        mapToolToModule(toolName, userId = null, sessionId = null) {
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Mapping tool to module', {
                    toolName,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            }
            
            try {
                // Special case for query
                if (toolName === 'query') {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_map_to_module_success', executionTime, {
                        toolName,
                        mappingType: 'special_case_query',
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, userId);
                    MonitoringService.info('Tool mapped to module successfully', {
                        toolName,
                        mappingType: 'special_case_query',
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                    return { moduleName: 'query', methodName: 'processQuery' };
                }
            
            // Special case for calendar.getAvailability
            if (toolName === 'calendar.getAvailability') {
                const executionTime = Date.now() - startTime;
                MonitoringService.trackMetric('tools_map_to_module_success', executionTime, {
                    toolName,
                    mappingType: 'special_case_calendar_get_availability',
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                MonitoringService.info('Tool mapped to module successfully', {
                    toolName,
                    mappingType: 'special_case_calendar_get_availability',
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
                return { moduleName: 'calendar', methodName: 'getAvailability' };
            }

            const lowerCaseToolName = toolName.toLowerCase();

            // Find modules that have this capability (case-insensitive)
            const modules = moduleRegistry.getAllModules();
            for (const module of modules) {
                if (Array.isArray(module.capabilities)) {
                    const lowerCaseCapabilities = module.capabilities.map(c => c.toLowerCase());
                    const capabilityIndex = lowerCaseCapabilities.indexOf(lowerCaseToolName);
                    if (capabilityIndex > -1) {
                        const executionTime = Date.now() - startTime;
                        MonitoringService.trackMetric('tools_map_to_module_success', executionTime, {
                            toolName,
                            mappingType: 'direct_capability',
                            moduleName: module.id,
                            timestamp: new Date().toISOString()
                        });
                        // Return the original capability name casing from the module definition
                        return { moduleName: module.id, methodName: module.capabilities[capabilityIndex] };
                    }
                }
            }

            // Check aliases if no direct capability match found
            const aliasTarget = toolAliases[toolName]; // Use original case for alias lookup

            if (aliasTarget) {
                // Validate the alias target
                const targetModule = moduleRegistry.getModule(aliasTarget.moduleName);
                if (!targetModule) {
                    MonitoringService.error(`Alias points to non-existent module`, {
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                    return null;
                }
                if (!Array.isArray(targetModule.capabilities) || !targetModule.capabilities.includes(aliasTarget.methodName)) {
                    MonitoringService.error(`Alias points to module without required capability`, {
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        targetMethod: aliasTarget.methodName,
                        availableCapabilities: targetModule.capabilities || [],
                        timestamp: new Date().toISOString()
                    }, 'tools');
                    return null;
                }
                
                const executionTime = Date.now() - startTime;
                MonitoringService.trackMetric('tools_map_to_module_success', executionTime, {
                    toolName,
                    mappingType: 'alias',
                    moduleName: aliasTarget.moduleName,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Tool mapped to module successfully via alias', {
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        targetMethod: aliasTarget.methodName,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Tool mapped to module with session via alias', {
                        sessionId,
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        targetMethod: aliasTarget.methodName,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Mapping tool to alias target', {
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        targetMethod: aliasTarget.methodName,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                }
                
                return aliasTarget;
            }

            MonitoringService.warn(`No module or valid alias found for tool`, {
                toolName,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'tools');
            
            const executionTime = Date.now() - startTime;
            MonitoringService.trackMetric('tools_map_to_module_not_found', executionTime, {
                toolName,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, userId);
            
            // Pattern 4: User Error Tracking (for not found)
            if (userId) {
                MonitoringService.error('Tool mapping not found', {
                    toolName,
                    error: 'No module or valid alias found',
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Tool mapping not found', {
                    sessionId,
                    toolName,
                    error: 'No module or valid alias found',
                    duration: executionTime,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            return null; // Not found
            
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'tools',
                    `Failed to map tool to module: ${error.message}`,
                    'error',
                    {
                        toolName,
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
                    MonitoringService.error('Failed to map tool to module', {
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to map tool to module', {
                        sessionId,
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_map_to_module_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                throw mcpError;
            }
        },

        /**
         * Transforms parameters for a specific tool with user context
         * @param {string} toolName - Name of the tool
         * @param {object} params - Original parameters
         * @param {string} [userId] - User ID for multi-user context
         * @param {string} [deviceId] - Device ID for multi-user context
         * @param {string} [sessionId] - Session ID for context
         * @returns {object} - Transformed parameters and module/method mapping
         */
        transformToolParameters(toolName, params, userId = null, deviceId = null, sessionId = null) {
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Transforming tool parameters', {
                    toolName,
                    paramKeys: Object.keys(params),
                    userId,
                    deviceId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId);
            }
            
            try {
                // First, map the tool to a module and method
                const mapping = this.mapToolToModule(toolName, userId, sessionId);
                
                if (!mapping) {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_transform_tool_params_no_mapping', executionTime, {
                        toolName,
                        userId,
                        deviceId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, userId);
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('No mapping found for tool', {
                            toolName,
                            error: 'Tool mapping not found',
                            duration: executionTime,
                            timestamp: new Date().toISOString()
                        }, 'tools', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('No mapping found for tool', {
                            sessionId,
                            toolName,
                            error: 'Tool mapping not found',
                            duration: executionTime,
                            timestamp: new Date().toISOString()
                        }, 'tools');
                    }
                    
                    return { 
                        mapping: null, 
                        params: params
                    };
                }
                
                // Then transform the parameters based on the module and method with user context
                const transformedParams = transformParameters(mapping.moduleName, mapping.methodName, params, userId, deviceId, sessionId);
                
                const executionTime = Date.now() - startTime;
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Tool parameters transformed successfully', {
                        toolName,
                        moduleName: mapping.moduleName,
                        methodName: mapping.methodName,
                        paramCount: Object.keys(params).length,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Tool parameters transformed with session', {
                        sessionId,
                        toolName,
                        moduleName: mapping.moduleName,
                        methodName: mapping.methodName,
                        paramCount: Object.keys(params).length,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_transform_tool_params_success', executionTime, {
                    toolName,
                    moduleName: mapping.moduleName,
                    methodName: mapping.methodName,
                    paramCount: Object.keys(params).length,
                    userId,
                    deviceId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                return {
                    mapping,
                    params: transformedParams
                };
                
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'tools',
                    `Failed to transform tool parameters: ${error.message}`,
                    'error',
                    {
                        toolName,
                        paramKeys: Object.keys(params),
                        error: error.message,
                        stack: error.stack,
                        userId,
                        deviceId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to transform tool parameters', {
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to transform tool parameters', {
                        sessionId,
                        toolName,
                        error: error.message,
                        duration: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                
                MonitoringService.trackMetric('tools_transform_tool_params_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    userId,
                    deviceId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, userId);
                
                throw mcpError;
            }
        },
        
        refresh // Expose the refresh method
    };
}

module.exports = createToolsService;
