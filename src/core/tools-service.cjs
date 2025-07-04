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
 * @returns {object} Tools service methods
 */
function createToolsService({ moduleRegistry, logger = console, schemaValidator = null }) {
    const startTime = Date.now();
    
    try {
        if (!moduleRegistry) {
            const mcpError = ErrorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                'Module registry is required for ToolsService',
                ErrorService.SEVERITIES.CRITICAL,
                {
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }
        
        MonitoringService.info('Tools service instance created', {
            hasModuleRegistry: !!moduleRegistry,
            hasLogger: !!logger,
            hasSchemaValidator: !!schemaValidator,
            timestamp: new Date().toISOString()
        }, 'tools');
        
    } catch (error) {
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('tools_service_creation_failure', executionTime, {
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
        throw error;
    }

    // Internal state for caching (will be used later)
    let cachedTools = null;

    // Comprehensive tool alias map for consistent module and method routing
    const toolAliases = {
        // Mail module tools
        getMail: { moduleName: 'mail', methodName: 'getInbox' },
        readMail: { moduleName: 'mail', methodName: 'getInbox' },
        sendMail: { moduleName: 'mail', methodName: 'sendEmail' },
        searchMail: { moduleName: 'mail', methodName: 'searchEmails' },
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
        searchFiles: { moduleName: 'files', methodName: 'searchFiles' },
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
        
        // Query module
        query: { moduleName: 'query', methodName: 'processQuery' }
    };

    /**
     * Generates a tool definition from a module capability
     * @param {string} moduleName - Name of the module
     * @param {string} capability - Capability/tool name
     * @returns {object} Tool definition
     */
    function generateToolDefinition(moduleName, capability) {
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
            case 'searchEmails':
            case 'searchMail':
                toolDef.description = 'Search emails using Microsoft Graph KQL (Keyword Query Language) syntax';
                toolDef.endpoint = '/api/v1/mail/search';
                toolDef.parameters = {
                    query: { 
                        type: 'string', 
                        description: 'KQL search query string. Examples: "from:user@domain.com", "subject:meeting", "from:john AND subject:report", "hasattachments:true". Do not wrap the entire query in quotes.', 
                        required: true,
                        aliases: ['q'] // Support both 'query' and 'q' parameters
                    },
                    limit: { 
                        type: 'number', 
                        description: 'Maximum number of results to return', 
                        optional: true,
                        default: 20
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
                toolDef.description = 'Fetch calendar events from Microsoft 365. IMPORTANT: Microsoft Graph has strict filter limitations - organizer email filtering often fails with HTTP 501 errors. Use convenience parameters (subject, organizer, attendee) instead of complex $filter expressions.';
                toolDef.endpoint = '/api/v1/calendar';
                toolDef.parameters = {
                    // Date range parameters
                    start: { 
                        type: 'string', 
                        description: 'Start date in ISO format (YYYY-MM-DD) to filter events from', 
                        optional: true,
                        format: 'date'
                    },
                    end: { 
                        type: 'string', 
                        description: 'End date in ISO format (YYYY-MM-DD) to filter events until', 
                        optional: true,
                        format: 'date'
                    },
                    
                    // Pagination and limits
                    limit: { 
                        type: 'number', 
                        description: 'Maximum number of events to return (1-999)', 
                        optional: true, 
                        default: 50,
                        minimum: 1,
                        maximum: 999
                    },
                    top: { 
                        type: 'number', 
                        description: 'Alias for limit - maximum number of events to return', 
                        optional: true
                    },
                    
                    // Microsoft Graph query parameters
                    filter: { 
                        type: 'string', 
                        description: 'OData $filter query - WARNING: Microsoft Graph has severe limitations. Many expressions cause HTTP 501 errors. AVOID organizer/emailAddress filters. Use simple expressions only like "contains(subject,\'text\')" or "start/dateTime ge \'2025-01-01T00:00:00Z\'"', 
                        optional: true
                    },
                    select: { 
                        type: 'string', 
                        description: 'Comma-separated list of properties to include in response (e.g., "subject,start,end,organizer")', 
                        optional: true
                    },
                    orderby: { 
                        type: 'string', 
                        description: 'Property to sort results by (e.g., "start/dateTime desc", "subject asc")', 
                        optional: true, 
                        default: 'start/dateTime'
                    },
                    expand: { 
                        type: 'string', 
                        description: 'Comma-separated list of related properties to expand (e.g., "attendees,calendar")', 
                        optional: true
                    },
                    
                    // Convenience filters (converted to $filter queries) - RECOMMENDED APPROACH
                    subject: { 
                        type: 'string', 
                        description: 'Filter events by subject containing this text (SAFE - uses contains() filter)', 
                        optional: true
                    },
                    organizer: { 
                        type: 'string', 
                        description: 'Filter events by organizer display name (e.g., "John Doe"). NOTE: Email addresses are NOT supported - use the person\'s display name instead. Microsoft Graph API supports organizer/emailAddress/name but not organizer/emailAddress/address filtering.', 
                        optional: true
                    },
                    attendee: { 
                        type: 'string', 
                        description: 'Filter events where this email address is an attendee (RISKY - uses lambda expressions that may fail)', 
                        optional: true
                    },
                    location: { 
                        type: 'string', 
                        description: 'Filter events by location containing this text (SAFE - uses contains() filter)', 
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
            case 'searchFiles':
                toolDef.description = 'Search for files by name or content. This tool must be used to find files before performing operations on them.';
                toolDef.endpoint = '/api/v1/files/search';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    q: { 
                        type: 'string', 
                        description: 'Search query to find files by name or content',
                        required: true
                    }
                };
                toolDef.parameterMapping = {
                    q: { inQuery: true }
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
                toolDef.description = 'Get the content of a specific file. Use searchFiles first to find the file ID.';
                toolDef.endpoint = '/api/v1/files/content';
                toolDef.method = 'GET';
                toolDef.parameters = {
                    id: { 
                        type: 'string', 
                        description: 'ID of the file to get content for (required, must be obtained from searchFiles or listFiles)',
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
                toolDef.description = 'IMPORTANT: Find and resolve people by name or email before scheduling meetings or sending emails. This tool MUST be used to resolve any person references before creating calendar events or sending mail.';
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
            case 'findPeople':
                toolDef.description = 'IMPORTANT: Find and resolve people by name or email before scheduling meetings or sending emails. This tool MUST be used to resolve any person references before creating calendar events or sending mail.';
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

        return toolDef;
    }

    /**
     * Invalidates any internal caches, forcing regeneration on next access.
     */
    function refresh() {
        const startTime = Date.now();
        
        try {
            const previousCacheSize = cachedTools ? cachedTools.length : 0;
            
            MonitoringService.info('Tools service refresh triggered, clearing internal cache', {
                previousCacheSize,
                timestamp: new Date().toISOString()
            }, 'tools');
            
            cachedTools = null; // Clear cache
            
            const executionTime = Date.now() - startTime;
            MonitoringService.trackMetric('tools_refresh_success', executionTime, {
                previousCacheSize,
                timestamp: new Date().toISOString()
            });
            
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            const mcpError = ErrorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Tools service refresh failed: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                {
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            MonitoringService.trackMetric('tools_refresh_failure', executionTime, {
                errorType: error.code || 'unknown',
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    /**
     * Helper function to transform attendees from string or array to proper format
     * @param {string|Array} attendees - Attendees as string or array
     * @returns {Array|undefined} - Transformed attendees or undefined if none
     */
    function transformAttendees(attendees) {
        if (!attendees) return undefined;
        
        // If attendees is a string (comma-separated), convert to array
        if (typeof attendees === 'string') {
            return attendees.split(',').map(email => email.trim());
        }
        
        // If already an array, return as is
        return attendees;
    }
    
    /**
     * Helper function to transform date/time to proper format
     * @param {string|object} dateTime - Date time string or object
     * @param {string} timeZone - Default timezone if not specified
     * @returns {object|undefined} - Transformed date time object
     */
    function transformDateTime(dateTime, timeZone = 'UTC') {
        if (!dateTime) return undefined;
        
        // If already an object with dateTime, return as is
        if (typeof dateTime === 'object' && dateTime.dateTime) {
            return dateTime;
        }
        
        // If a string, convert to object format
        if (typeof dateTime === 'string') {
            return {
                dateTime: dateTime,
                timeZone: timeZone
            };
        }
        
        return dateTime;
    }
    
    /**
     * Transforms parameters for a specific module and method with user context
     * @param {string} moduleName - Module name
     * @param {string} methodName - Method name
     * @param {object} params - Original parameters
     * @param {string} [userId] - User ID for multi-user context
     * @param {string} [deviceId] - Device ID for multi-user context
     * @returns {object} - Transformed parameters
     */
    function transformParameters(moduleName, methodName, params, userId = null, deviceId = null) {
        const startTime = Date.now();
        
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Transforming parameters', {
                moduleName,
                methodName,
                paramKeys: Object.keys(params),
                userId,
                deviceId,
                timestamp: new Date().toISOString()
            }, 'tools', null, userId, deviceId);
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
                    timestamp: new Date().toISOString()
                });
                return transformedParams;
        }
        
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            const mcpError = ErrorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Parameter transformation failed: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                {
                    moduleName,
                    methodName,
                    paramKeys: Object.keys(params),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(mcpError);
            MonitoringService.trackMetric('tools_transform_params_failure', executionTime, {
                moduleName,
                methodName,
                errorType: error.code || 'unknown',
                timestamp: new Date().toISOString()
            });
            
            throw mcpError;
        }
    }

    return {
        /**
         * Gets all available tools from registered modules
         * @returns {Array<object>} Tool definitions
         */
        getAllTools() {
            const startTime = Date.now();
            
            try {
                // Check cache first
                if (cachedTools) {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_get_all_cache_hit', executionTime, {
                        toolCount: cachedTools.length,
                        timestamp: new Date().toISOString()
                    });
                    
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Returning cached tool definitions', {
                            toolCount: cachedTools.length,
                            timestamp: new Date().toISOString()
                        }, 'tools');
                    }
                    return cachedTools;
                }

                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Cache miss, generating tool definitions', {
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                const tools = [];
            const modules = moduleRegistry.getAllModules();
            
            // First, add the findPeople tool at the beginning of the list
            // This is critical because person resolution must happen before scheduling or sending invites
            const peopleModule = modules.find(m => m.name === 'people' || m.capabilities?.includes('findPeople'));
            if (peopleModule) {
                const findPeopleTool = {
                    name: 'findPeople',
                    description: 'IMPORTANT: Find and resolve people by name or email before scheduling meetings or sending emails. This tool MUST be used to resolve any person references before creating calendar events or sending mail.',
                    endpoint: '/api/v1/people/find',
                    method: 'GET',
                    parameters: {
                        query: { type: 'string', description: 'Search query to find a person' },
                        name: { type: 'string', description: 'Person name to search for', optional: true },
                        limit: { type: 'number', description: 'Maximum number of results', optional: true }
                    }
                };
                tools.push(findPeopleTool);
                logger.debug('getAllTools: Added findPeople tool with high priority');
            }

            // For each module, generate tool definitions for each capability (except findPeople which we already added)
            for (const module of modules) {
                if (Array.isArray(module.capabilities)) {
                    for (const capability of module.capabilities) {
                        // Skip findPeople since we already added it
                        if (capability === 'findPeople') continue;
                        tools.push(generateToolDefinition(module.name, capability));
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
                timestamp: new Date().toISOString()
            });
            
            MonitoringService.info('Generated and cached tool definitions', {
                toolCount: tools.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'tools');

            return tools;
            
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                const mcpError = ErrorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    `Failed to get all tools: ${error.message}`,
                    ErrorService.SEVERITIES.ERROR,
                    {
                        stack: error.stack,
                        timestamp: new Date().toISOString()
                    }
                );
                
                MonitoringService.logError(mcpError);
                MonitoringService.trackMetric('tools_get_all_failure', executionTime, {
                    errorType: error.code || 'unknown',
                    timestamp: new Date().toISOString()
                });
                
                throw mcpError;
            }
        },

        /**
         * Gets a tool definition by name
         * @param {string} toolName - Name of the tool
         * @returns {object|null} Tool definition or null if not found
         */
        getToolByName(toolName) {
            const startTime = Date.now();
            
            try {
                const allTools = this.getAllTools(); // Uses cache if available
                const lowerCaseToolName = toolName.toLowerCase();
                const foundTool = allTools.find(tool => tool.name.toLowerCase() === lowerCaseToolName);
                
                const executionTime = Date.now() - startTime;
                MonitoringService.trackMetric('tools_get_by_name_success', executionTime, {
                    toolName,
                    found: !!foundTool,
                    totalTools: allTools.length,
                    timestamp: new Date().toISOString()
                });
                
                return foundTool || null;
                
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                const mcpError = ErrorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    `Failed to get tool by name: ${error.message}`,
                    ErrorService.SEVERITIES.ERROR,
                    {
                        toolName,
                        stack: error.stack,
                        timestamp: new Date().toISOString()
                    }
                );
                
                MonitoringService.logError(mcpError);
                MonitoringService.trackMetric('tools_get_by_name_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    timestamp: new Date().toISOString()
                });
                
                throw mcpError;
            }
        },

        /**
         * Maps a tool name to a module and method
         * @param {string} toolName - Name of the tool
         * @returns {object|null} Module and method mapping or null if not found
         */
        mapToolToModule(toolName) {
            const startTime = Date.now();
            
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Mapping tool to module', {
                    toolName,
                    timestamp: new Date().toISOString()
                }, 'tools');
            }
            
            try {
                // Special case for query
                if (toolName === 'query') {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_map_to_module_success', executionTime, {
                        toolName,
                        mappingType: 'special_case_query',
                        timestamp: new Date().toISOString()
                    });
                    return { moduleName: 'query', methodName: 'processQuery' };
                }
            
            // Special case for calendar.getAvailability
            if (toolName === 'calendar.getAvailability') {
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
                    timestamp: new Date().toISOString()
                });
                
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Mapping tool to alias target', {
                        toolName,
                        targetModule: aliasTarget.moduleName,
                        targetMethod: aliasTarget.methodName,
                        timestamp: new Date().toISOString()
                    }, 'tools');
                }
                return aliasTarget;
            }

            MonitoringService.warn(`No module or valid alias found for tool`, {
                toolName,
                timestamp: new Date().toISOString()
            }, 'tools');
            
            const executionTime = Date.now() - startTime;
            MonitoringService.trackMetric('tools_map_to_module_not_found', executionTime, {
                toolName,
                timestamp: new Date().toISOString()
            });
            
            return null; // Not found
            
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                const mcpError = ErrorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    `Failed to map tool to module: ${error.message}`,
                    ErrorService.SEVERITIES.ERROR,
                    {
                        toolName,
                        stack: error.stack,
                        timestamp: new Date().toISOString()
                    }
                );
                
                MonitoringService.logError(mcpError);
                MonitoringService.trackMetric('tools_map_to_module_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    timestamp: new Date().toISOString()
                });
                
                throw mcpError;
            }
        },

        /**
         * Transforms parameters for a specific tool with user context
         * @param {string} toolName - Name of the tool
         * @param {object} params - Original parameters
         * @param {string} [userId] - User ID for multi-user context
         * @param {string} [deviceId] - Device ID for multi-user context
         * @returns {object} - Transformed parameters and module/method mapping
         */
        transformToolParameters(toolName, params, userId = null, deviceId = null) {
            const startTime = Date.now();
            
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Transforming tool parameters', {
                    toolName,
                    paramKeys: Object.keys(params),
                    userId,
                    deviceId,
                    timestamp: new Date().toISOString()
                }, 'tools', null, userId, deviceId);
            }
            
            try {
                // First, map the tool to a module and method
                const mapping = this.mapToolToModule(toolName);
                
                if (!mapping) {
                    const executionTime = Date.now() - startTime;
                    MonitoringService.trackMetric('tools_transform_tool_params_no_mapping', executionTime, {
                        toolName,
                        userId,
                        deviceId,
                        timestamp: new Date().toISOString()
                    }, userId, deviceId);
                    
                    MonitoringService.error(`No mapping found for tool`, {
                        toolName,
                        userId,
                        deviceId,
                        timestamp: new Date().toISOString()
                    }, 'tools', null, userId, deviceId);
                    
                    return { 
                        mapping: null, 
                        params: params
                    };
                }
                
                // Then transform the parameters based on the module and method with user context
                const transformedParams = transformParameters(mapping.moduleName, mapping.methodName, params, userId, deviceId);
                
                const executionTime = Date.now() - startTime;
                MonitoringService.trackMetric('tools_transform_tool_params_success', executionTime, {
                    toolName,
                    moduleName: mapping.moduleName,
                    methodName: mapping.methodName,
                    paramCount: Object.keys(params).length,
                    userId,
                    deviceId,
                    timestamp: new Date().toISOString()
                }, userId, deviceId);
                
                return {
                    mapping,
                    params: transformedParams
                };
                
            } catch (error) {
                const executionTime = Date.now() - startTime;
                
                const mcpError = ErrorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    `Failed to transform tool parameters: ${error.message}`,
                    ErrorService.SEVERITIES.ERROR,
                    {
                        toolName,
                        paramKeys: Object.keys(params),
                        stack: error.stack,
                        userId,
                        deviceId,
                        timestamp: new Date().toISOString()
                    },
                    null,
                    userId,
                    deviceId
                );
                
                MonitoringService.logError(mcpError);
                MonitoringService.trackMetric('tools_transform_tool_params_failure', executionTime, {
                    toolName,
                    errorType: error.code || 'unknown',
                    userId,
                    deviceId,
                    timestamp: new Date().toISOString()
                }, userId, deviceId);
                
                throw mcpError;
            }
        },
        
        refresh // Expose the refresh method
    };
}

module.exports = createToolsService;
