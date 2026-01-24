/**
 * @fileoverview Handles /api/calendar endpoints for calendar operations.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Helper function to validate request against schema and log validation errors
 * @param {object} req - Express request object
 * @param {object} schema - Joi schema to validate against
 * @param {string} endpoint - Endpoint path for error context
 * @param {object} [additionalContext] - Additional context for validation errors
 * @returns {object} Object with error and value properties
 */
const validateAndLog = (req, schema, endpoint, additionalContext = {}) => {
    const result = schema.validate(req.body);
    
    if (result.error) {
        const { userId = null, deviceId = null } = additionalContext;
        const validationError = ErrorService?.createError(
            'api', 
            `${endpoint} validation error`, 
            'warning', 
            { 
                details: result.error.details,
                endpoint,
                ...additionalContext
            },
            null,
            userId,
            deviceId
        );
        MonitoringService?.logError(validationError);
    }
    
    return result;
};

/**
 * Helper function to check if a module method is available
 * @param {string} methodName - Name of the method to check
 * @param {object} module - Module to check for method availability
 * @returns {boolean} Whether the method exists on the module
 */
const isModuleMethodAvailable = (methodName, module) => {
    return typeof module[methodName] === 'function';
};

/**
 * Factory for calendar controller with dependency injection.
 * @param {object} deps - { calendarModule }
 */
module.exports = ({ calendarModule }) => ({
    /**
     * GET /api/calendar
     * Retrieves calendar events with optional filtering and debug information
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async getEvents(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar events request', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    query: req.query,
                    ip: req.ip,
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Log request with user context
            MonitoringService?.info(`Processing ${req.method} ${req.path}`, {
                method: req.method,
                path: req.path,
                query: req.query,
                ip: req.ip,
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id
            }, 'calendar', null, actualUserId, deviceId);
            
            // Validate query params with Joi
            const querySchema = Joi.object({
                limit: Joi.number().integer().min(1).max(100).default(20),
                filter: Joi.string().optional(),
                debug: Joi.boolean().default(false).optional(),
                startDateTime: Joi.date().iso().optional(),
                endDateTime: Joi.date().iso().optional(),
                // Convenience filter parameters
                organizer: Joi.string().optional(),
                subject: Joi.string().optional(),
                location: Joi.string().optional(),
                attendee: Joi.string().optional()
            });
            
            // Convert query parameters for validation
            const queryParams = {
                limit: req.query.limit ? Number(req.query.limit) : undefined,
                filter: req.query.filter,
                debug: req.query.debug === 'true',
                startDateTime: req.query.startDateTime,
                endDateTime: req.query.endDateTime,
                // Convenience filter parameters
                organizer: req.query.organizer,
                subject: req.query.subject,
                location: req.query.location,
                attendee: req.query.attendee
            };
            
            const { error, value } = querySchema.validate(queryParams);
            if (error) {
                const validationError = ErrorService?.createError('api', 'Query parameter validation error', 'warning', { 
                    details: error.details,
                    endpoint,
                    query: req.query,
                    userId: actualUserId,
                    deviceId
                }, null, actualUserId, deviceId);
                MonitoringService?.logError(validationError);
                return res.status(400).json({ 
                    error: 'Invalid query parameters', 
                    details: error.details 
                });
            }
            
            const { limit: top, filter, debug, organizer, subject, location, attendee } = value;
            let rawEvents = null;
            
            // For debugging, get raw events if requested
            if (debug && isModuleMethodAvailable('getEventsRaw', calendarModule)) {
                try {
                    MonitoringService?.info('Fetching raw events for debug', { 
                        top, 
                        filter,
                        organizer,
                        subject,
                        location,
                        attendee,
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                    // Pass req object for user-scoped token selection, but don't pass internal userId to Graph API
                    // The internal userId is only for token storage - Graph API should use 'me' (default)
                    rawEvents = await calendarModule.getEventsRaw({ top, filter, organizer, subject, location, attendee }, req);
                    MonitoringService?.info(`Retrieved ${rawEvents.length} raw events`, { 
                        count: rawEvents.length, 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                } catch (fetchError) {
                    const error = ErrorService?.createError('api', 'Error fetching raw events', 'error', { 
                        error: fetchError.message,
                        userId: actualUserId,
                        deviceId
                    }, null, actualUserId, deviceId);
                    MonitoringService?.logError(error);
                    // Continue even if raw fetch fails
                }
            }
            
            // Try to get events from the module, or return mock data if it fails
            let events = [];
            let isMock = false;
            try {
                MonitoringService?.info('Attempting to get real calendar events from module', { 
                    top, 
                    filter,
                    organizer,
                    subject,
                    location,
                    attendee,
                    userId: actualUserId, 
                    deviceId 
                }, 'calendar', null, actualUserId, deviceId);
                
                if (isModuleMethodAvailable('getEvents', calendarModule)) {
                    // Pass req object for user-scoped token selection, but don't pass internal userId to Graph API
                    // The internal userId is only for token storage - Graph API should use 'me' (default)
                    events = await calendarModule.getEvents({ top, filter, organizer, subject, location, attendee }, req);
                    MonitoringService?.info(`Successfully retrieved ${events.length} real calendar events`, { 
                        count: events.length, 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                } else if (isModuleMethodAvailable('handleIntent', calendarModule)) {
                    // Try using the module's handleIntent method instead
                    MonitoringService?.info('Falling back to handleIntent method', { 
                        intent: 'readCalendar', 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                    const result = await calendarModule.handleIntent('readCalendar', { count: top, filter, organizer, subject, location, attendee }, { req });
                    events = result && result.items ? result.items : [];
                    MonitoringService?.info(`Retrieved ${events.length} events via handleIntent`, { 
                        count: events.length, 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                } else {
                    throw new Error('No calendar module method available for getting events');
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', 'Error calling calendar module for events', 'error', { 
                    error: moduleError.message,
                    userId: actualUserId,
                    deviceId
                }, null, actualUserId, deviceId);
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock calendar data', {
                    userId: actualUserId,
                    deviceId
                }, 'calendar', null, actualUserId, deviceId);
                
                // Return mock data only if the real module call fails
                const today = new Date();
                const tomorrow = new Date(today);
                tomorrow.setDate(tomorrow.getDate() + 1);
                
                events = [
                    { 
                        id: 'mock1', 
                        subject: 'Team Meeting', 
                        start: { dateTime: today.toISOString(), timeZone: 'UTC' },
                        end: { dateTime: new Date(today.getTime() + 3600000).toISOString(), timeZone: 'UTC' },
                        isMock: true
                    },
                    { 
                        id: 'mock2', 
                        subject: 'Project Review', 
                        start: { dateTime: tomorrow.toISOString(), timeZone: 'UTC' },
                        end: { dateTime: new Date(tomorrow.getTime() + 3600000).toISOString(), timeZone: 'UTC' },
                        isMock: true
                    }
                ];
                isMock = true;
            }
            
            // Track performance
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.getEvents.duration', duration, {
                eventCount: events.length,
                hasFilter: !!filter,
                debug,
                success: true,
                isMock,
                userId: actualUserId,
                deviceId
            });
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar events retrieved successfully', {
                    eventCount: events.length,
                    hasRawEvents: !!rawEvents,
                    debugMode: debug,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar events retrieved with session', {
                    sessionId: req.session.id,
                    eventCount: events.length,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            if (debug) {
                res.json({ events, rawEvents, debug: true });
            } else {
                res.json(events);
            }
        } catch (err) {
            // Track error metrics
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.getEvents.error', 1, {
                errorMessage: err.message,
                duration,
                success: false,
                userId: actualUserId,
                deviceId
            });
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to retrieve calendar events',
                'error',
                {
                    endpoint: '/api/calendar',
                    error: err.message,
                    stack: err.stack,
                    operation: 'getEvents',
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar events retrieval failed', {
                    error: err.message,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar events retrieval failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            res.status(500).json({ 
                error: 'calendar_events_error',
                error_description: 'Unable to retrieve calendar events',
                errorId: mcpError.id
            });
        }
    },
    /**
     * POST /api/calendar/create
     * Creates a new calendar event
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async createEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/create';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event creation', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    hasBody: !!req.body,
                    bodyKeys: req.body ? Object.keys(req.body) : [],
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Log request with user context
            MonitoringService?.info(`Processing ${req.method} ${req.path}`, {
                method: req.method,
                path: req.path,
                ip: req.ip,
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id
            }, 'calendar', null, actualUserId, deviceId);
            
            // Joi schema for createEvent with standardized dateTime validation
            const eventSchema = Joi.object({
                subject: Joi.string().min(1).required(),
                start: Joi.object({
                    dateTime: Joi.date().iso().required(),
                    timeZone: Joi.string().default('UTC')
                }).required(),
                end: Joi.object({
                    dateTime: Joi.date().iso().required(),
                    timeZone: Joi.string().default('UTC')
                }).required(),
                // Support multiple attendee formats
                attendees: Joi.alternatives().try(
                    // Simple email strings
                    Joi.array().items(Joi.string().email()),
                    
                    // Simple objects with email property
                    Joi.array().items(Joi.object({
                        email: Joi.string().email().required(),
                        name: Joi.string().optional(),
                        type: Joi.string().valid('required', 'optional').optional()
                    })),
                    
                    // Full Graph API format with emailAddress
                    Joi.array().items(Joi.object({
                        emailAddress: Joi.object({
                            address: Joi.string().email().required(),
                            name: Joi.string().optional()
                        }).required(),
                        type: Joi.string().valid('required', 'optional').optional()
                    }))
                ).optional(),
                // Body can be string or object with contentType and content
                body: Joi.alternatives().try(
                    Joi.string(),
                    Joi.object({
                        contentType: Joi.string().valid('text', 'html', 'HTML').optional(),
                        content: Joi.string().required()
                    })
                ).optional(),
                // Location can be string or object with displayName
                location: Joi.alternatives().try(
                    Joi.string(),
                    Joi.object({
                        displayName: Joi.string().required(),
                        address: Joi.object().optional()
                    })
                ).optional(),
                isAllDay: Joi.boolean().optional(),
                isOnlineMeeting: Joi.boolean().optional(),
                recurrence: Joi.object().optional()
            });
            
            // Log the incoming request body for debugging (redacting sensitive data)
            const safeReqBody = { ...req.body };
            if (safeReqBody.attendees) {
                // Redact attendee details for privacy
                safeReqBody.attendees = Array.isArray(safeReqBody.attendees) ? 
                    `[${safeReqBody.attendees.length} attendees]` : 'attendees present';
            }
            
            MonitoringService?.info('Processing create event request', { 
                subject: safeReqBody.subject,
                hasAttendees: !!safeReqBody.attendees,
                hasLocation: !!safeReqBody.location,
                hasStartTime: !!safeReqBody.start,
                hasEndTime: !!safeReqBody.end,
                userId: actualUserId,
                deviceId
            }, 'calendar', null, actualUserId, deviceId);
            
            const { error, value } = validateAndLog(req, eventSchema, 'Create event', { endpoint, userId: actualUserId, deviceId });
            if (error) {
                return res.status(400).json({ error: 'Invalid request', details: error.details });
            }
            
            // Try to create a real event using the calendar module
            let event;
            try {
                MonitoringService?.info('Attempting to create real calendar event using module', { 
                    subject: value.subject, 
                    userId: actualUserId, 
                    deviceId 
                }, 'calendar', null, actualUserId, deviceId);
                
                const methodName = 'createEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    // Pass req object for user-scoped token selection, but don't pass internal userId to Graph API
                    // The internal userId is only for token storage - Graph API should use 'me' (default)
                    event = await calendarModule[methodName](value, req);
                    MonitoringService?.info('Successfully created real calendar event', { 
                        eventId: event.id, 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                } else if (isModuleMethodAvailable('handleIntent', calendarModule)) {
                    // Try using the module's handleIntent method instead
                    MonitoringService?.info('Falling back to handleIntent method for event creation', { 
                        intent: 'createEvent', 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                    const result = await calendarModule.handleIntent('createEvent', value, { req });
                    event = result;
                    MonitoringService?.info('Created event via handleIntent', { 
                        eventId: event.id, 
                        userId: actualUserId, 
                        deviceId 
                    }, 'calendar', null, actualUserId, deviceId);
                } else {
                    throw new Error('No calendar module method available for event creation');
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', 'Error calling calendar module for event creation', 'error', { 
                    error: moduleError.message,
                    subject: value.subject,
                    userId: actualUserId,
                    deviceId
                }, null, actualUserId, deviceId);
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event creation', {
                    userId: actualUserId,
                    deviceId
                }, 'calendar', null, actualUserId, deviceId);
                
                // Generate a mock event for testing
                event = {
                    id: `mock-event-${Date.now()}`,
                    subject: value.subject,
                    start: value.start,
                    end: value.end,
                    attendees: value.attendees || [],
                    body: value.body || '',
                    location: value.location || '',
                    isMock: true,
                    created: new Date().toISOString()
                };
            }
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event created successfully', {
                    eventId: event.id,
                    subject: value.subject,
                    hasAttendees: !!(value.attendees && value.attendees.length > 0),
                    hasLocation: !!value.location,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event created with session', {
                    sessionId: req.session.id,
                    eventId: event.id,
                    subject: value.subject,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track performance
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.createEvent.duration', duration, {
                success: true,
                hasAttendees: !!(value.attendees && value.attendees.length > 0),
                hasLocation: !!value.location,
                isMock: !!event.isMock,
                userId: actualUserId,
                deviceId
            });
            
            res.json(event);
        } catch (err) {
            // Track error metrics
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.createEvent.error', 1, {
                errorMessage: err.message,
                duration,
                success: false,
                userId: actualUserId,
                deviceId
            });
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to create calendar event',
                'error',
                {
                    endpoint: '/api/calendar/create',
                    error: err.message,
                    stack: err.stack,
                    operation: 'createEvent',
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event creation failed', {
                    error: err.message,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event creation failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            res.status(500).json({ 
                error: 'calendar_create_error',
                error_description: 'Unable to create calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * PUT /api/calendar/events/:id
     * Updates an existing calendar event
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async updateEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event update', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    eventId: req.params.id,
                    hasBody: !!req.body,
                    bodyKeys: req.body ? Object.keys(req.body) : [],
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for update', 'warning', { 
                    endpoint 
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const attendeeSchema = Joi.object({
                emailAddress: Joi.object({
                    address: Joi.string().email().required(),
                    name: Joi.string().optional()
                }).required(),
                type: Joi.string().valid('required', 'optional', 'resource').default('required')
            });

            const updateSchema = Joi.object({
                subject: Joi.string().optional(),
                start: Joi.object({
                    dateTime: Joi.date().iso().required(),
                    timeZone: Joi.string().default('UTC')
                }).optional(),
                end: Joi.object({
                    dateTime: Joi.date().iso().required(),
                    timeZone: Joi.string().default('UTC')
                }).optional(),
                // Support both string array and object array for attendees
                attendees: Joi.array().items(attendeeSchema).optional(),
                // Location can be string or object with displayName
                location: Joi.alternatives().try(
                    Joi.string(),
                    Joi.object({
                        displayName: Joi.string().required(),
                        address: Joi.object().optional()
                    })
                ).optional(),
                // Body can be string or object with contentType and content
                body: Joi.alternatives().try(
                    Joi.string(),
                    Joi.object({
                        contentType: Joi.string().valid('text', 'html', 'HTML').optional(),
                        content: Joi.string().required()
                    })
                ).optional(),
                isAllDay: Joi.boolean().optional(),
                isOnlineMeeting: Joi.boolean().optional(),
                recurrence: Joi.object().optional()
            });
            
            // Create a sanitized request body for logging (redact sensitive data)
            const safeReqBody = { ...req.body };
            if (safeReqBody.attendees) {
                // Redact attendee details for privacy
                safeReqBody.attendees = Array.isArray(safeReqBody.attendees) ? 
                    `[${safeReqBody.attendees.length} attendees]` : 'attendees present';
            }
            
            MonitoringService?.info('Updating calendar event', { 
                eventId,
                subject: safeReqBody.subject,
                hasAttendees: !!safeReqBody.attendees,
                hasLocation: !!safeReqBody.location,
                hasStartTime: !!safeReqBody.start,
                hasEndTime: !!safeReqBody.end,
                userId: actualUserId,
                deviceId
            }, 'calendar', null, actualUserId, deviceId);
            
            const { error, value } = validateAndLog(req, updateSchema, 'Update event', { eventId, endpoint, userId: actualUserId, deviceId });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            // Try to use the module's updateEvent method if available
            let updatedEvent;
            try {
                MonitoringService?.info(`Attempting to update event ${eventId} using module`, { eventId }, 'calendar');
                
                const methodName = 'updateEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    updatedEvent = await calendarModule[methodName](eventId, value, req);
                    MonitoringService?.info(`Successfully updated event ${eventId} using module`, { eventId }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error updating event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event update', { eventId }, 'calendar');
                
                // If module method fails, create a mock updated event
                updatedEvent = {
                    id: eventId,
                    ...value,
                    organizer: {
                        name: 'Current User',
                        email: 'current.user@example.com'
                    },
                    lastModifiedDateTime: new Date().toISOString(),
                    isMock: true // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock updated event', { eventId }, 'calendar');
            }
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event updated successfully', {
                    eventId: eventId,
                    subject: updatedEvent.subject,
                    hasAttendees: !!(value.attendees && value.attendees.length > 0),
                    hasLocation: !!value.location,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event updated with session', {
                    sessionId: req.session.id,
                    eventId: eventId,
                    subject: updatedEvent.subject,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track update time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.updateEvent.duration', duration, { 
                eventId,
                subject: updatedEvent.subject,
                isMock: !!updatedEvent.isMock
            });
            
            res.json(updatedEvent);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to update calendar event',
                'error',
                {
                    endpoint: '/api/calendar/events/:id',
                    error: err.message,
                    stack: err.stack,
                    operation: 'updateEvent',
                    eventId: req.params?.id,
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event update failed', {
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event update failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.updateEvent.error', 1, { 
                errorId: mcpError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'calendar_update_error',
                error_description: 'Unable to update calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * POST /api/calendar/events/:id/accept
     * Accept a calendar event invitation
     */
    /**
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async acceptEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id/accept';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event acceptance', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    eventId: req.params.id,
                    hasBody: !!req.body,
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for accept operation', 'warning', { 
                    endpoint 
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const acceptSchema = Joi.object({
                comment: Joi.string().optional()
            });
            
            const { error, value } = validateAndLog(req, acceptSchema, 'Accept event', { eventId, endpoint });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Accepting calendar event', { 
                eventId,
                hasComment: !!value.comment 
            }, 'calendar');
            
            // Try to use the module's acceptEvent method if available
            let result;
            try {
                MonitoringService?.info(`Attempting to accept event ${eventId} using module`, { eventId }, 'calendar');
                const methodName = 'acceptEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    result = await calendarModule[methodName](eventId, value.comment, req);
                    MonitoringService?.info(`Successfully accepted event ${eventId} using module`, { eventId }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error accepting event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event acceptance', { eventId }, 'calendar');
                
                // If module method fails, create a proper event response result
                result = {
                    success: true,
                    eventId: eventId,
                    responseType: 'accept',
                    status: 'accepted',
                    timestamp: new Date().toISOString(),
                    isMock: true  // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock event acceptance response', { eventId }, 'calendar');
            }
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event accepted successfully', {
                    eventId: eventId,
                    hasComment: !!value.comment,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event accepted with session', {
                    sessionId: req.session.id,
                    eventId: eventId,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track accept time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.acceptEvent.duration', duration, { 
                eventId,
                status: result.status,
                isMock: !!result.isMock
            });
            
            res.json(result);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to accept calendar event',
                'error',
                {
                    endpoint: '/api/calendar/events/:id/accept',
                    error: err.message,
                    stack: err.stack,
                    operation: 'acceptEvent',
                    eventId: req.params?.id,
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event acceptance failed', {
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event acceptance failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.acceptEvent.error', 1, { 
                errorId: mcpError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'calendar_accept_error',
                error_description: 'Unable to accept calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * POST /api/calendar/events/:id/tentativelyAccept
     * Tentatively accept a calendar event invitation
     */
    /**
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async tentativelyAcceptEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id/tentativelyAccept';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event tentative acceptance', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    eventId: req.params.id,
                    hasBody: !!req.body,
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for tentatively accept operation', 'warning', { 
                    endpoint
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const acceptSchema = Joi.object({
                comment: Joi.string().optional()
            });
            
            const { error, value } = validateAndLog(req, acceptSchema, 'Tentative accept event', { eventId, endpoint });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Tentatively accepting calendar event', { 
                eventId,
                hasComment: !!value.comment 
            }, 'calendar');
            
            // Try to use the module's tentativelyAcceptEvent method if available
            let result;
            try {
                MonitoringService?.info(`Attempting to tentatively accept event ${eventId} using module`, { eventId }, 'calendar');
                const methodName = 'tentativelyAcceptEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    result = await calendarModule[methodName](eventId, value.comment, req);
                    MonitoringService?.info(`Successfully tentatively accepted event ${eventId} using module`, { eventId }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error tentatively accepting event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event tentative acceptance', { eventId }, 'calendar');
                
                // If module method fails, create a proper event response result
                result = {
                    success: true,
                    eventId: eventId,
                    responseType: 'tentativelyAccept',
                    status: 'tentativelyAccepted',
                    timestamp: new Date().toISOString(),
                    isMock: true // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock event tentative acceptance response', { eventId }, 'calendar');
            }
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event tentatively accepted successfully', {
                    eventId: eventId,
                    hasComment: !!value.comment,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event tentatively accepted with session', {
                    sessionId: req.session.id,
                    eventId: eventId,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track tentative accept time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.tentativelyAcceptEvent.duration', duration, { 
                eventId,
                status: result.status,
                isMock: !!result.isMock
            });
            
            res.json(result);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to tentatively accept calendar event',
                'error',
                {
                    endpoint: '/api/calendar/events/:id/tentativelyAccept',
                    error: err.message,
                    stack: err.stack,
                    operation: 'tentativelyAcceptEvent',
                    eventId: req.params?.id,
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event tentative acceptance failed', {
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event tentative acceptance failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.tentativelyAcceptEvent.error', 1, { 
                errorId: mcpError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'calendar_tentative_accept_error',
                error_description: 'Unable to tentatively accept calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * POST /api/calendar/events/:id/decline
     * Decline a calendar event invitation
     */
    /**
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async declineEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id/decline';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event decline', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    eventId: req.params.id,
                    hasBody: !!req.body,
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for decline operation', 'warning', { 
                    endpoint
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const declineSchema = Joi.object({
                comment: Joi.string().optional()
            });
            
            const { error, value } = validateAndLog(req, declineSchema, 'Decline event', { eventId, endpoint });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Declining calendar event', { 
                eventId,
                hasComment: !!value.comment 
            }, 'calendar');
            
            // Try to use the module's declineEvent method if available
            let result;
            try {
                MonitoringService?.info(`Attempting to decline event ${eventId} using module`, { eventId }, 'calendar');
                const methodName = 'declineEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    result = await calendarModule[methodName](eventId, value.comment, req);
                    MonitoringService?.info(`Successfully declined event ${eventId} using module`, { eventId }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error declining event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event decline', { eventId }, 'calendar');
                
                // If module method fails, create a proper event response result
                result = {
                    success: true,
                    eventId: eventId,
                    responseType: 'decline',
                    status: 'declined',
                    timestamp: new Date().toISOString(),
                    isMock: true // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock event decline response', { eventId }, 'calendar');
            }
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event declined successfully', {
                    eventId: eventId,
                    hasComment: !!value.comment,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event declined with session', {
                    sessionId: req.session.id,
                    eventId: eventId,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track decline time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.declineEvent.duration', duration, { 
                eventId,
                status: result.status,
                isMock: !!result.isMock
            });
            
            res.json(result);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to decline calendar event',
                'error',
                {
                    endpoint: '/api/calendar/events/:id/decline',
                    error: err.message,
                    stack: err.stack,
                    operation: 'declineEvent',
                    eventId: req.params?.id,
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event decline failed', {
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event decline failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.declineEvent.error', 1, { 
                errorId: mcpError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'calendar_decline_error',
                error_description: 'Unable to decline calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * POST /api/calendar/events/:id/cancel
     * Cancel a calendar event and send cancellation messages to attendees
     */
    /**
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async cancelEvent(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id/cancel';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar event cancellation', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    eventId: req.params.id,
                    hasBody: !!req.body,
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for cancel operation', 'warning', { 
                    endpoint
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const cancelSchema = Joi.object({
                comment: Joi.string().optional()
            });
            
            const { error, value } = validateAndLog(req, cancelSchema, 'Cancel event', { eventId, endpoint, userId: actualUserId, deviceId });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Cancelling calendar event', { 
                eventId,
                hasComment: !!value.comment,
                userId: actualUserId,
                deviceId
            }, 'calendar', null, actualUserId, deviceId);
            
            // Try to use the module's cancelEvent method if available
            let result;
            try {
                MonitoringService?.info(`Attempting to cancel event ${eventId} using module`, { eventId, userId: actualUserId, deviceId }, 'calendar', null, actualUserId, deviceId);
                const methodName = 'cancelEvent';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    result = await calendarModule[methodName](eventId, value.comment, req);
                    MonitoringService?.info(`Successfully cancelled event ${eventId} using module`, { eventId, userId: actualUserId, deviceId }, 'calendar', null, actualUserId, deviceId);
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error cancelling event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock event cancellation', { eventId, userId: actualUserId, deviceId }, 'calendar', null, actualUserId, deviceId);
                
                // If module method fails, create a mock response
                result = {
                    id: eventId,
                    status: 'cancelled',
                    timestamp: new Date().toISOString(),
                    isMock: true // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock event cancellation', { eventId, userId: actualUserId, deviceId }, 'calendar', null, actualUserId, deviceId);
            }
            
            // Track cancel time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.cancelEvent.duration', duration, { 
                eventId,
                status: result.status,
                isMock: !!result.isMock
            });
            
            // Pattern 2: User Activity Logs
            if (actualUserId) {
                MonitoringService?.info('Calendar event cancelled successfully', {
                    eventId: eventId,
                    hasComment: !!value.comment,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.info('Calendar event cancelled with session', {
                    sessionId: req.session.id,
                    eventId: eventId,
                    status: result.status,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            res.json(result);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService?.createError(
                'calendar',
                'Failed to cancel calendar event',
                'error',
                {
                    endpoint: '/api/calendar/events/:id/cancel',
                    error: err.message,
                    stack: err.stack,
                    operation: 'cancelEvent',
                    eventId: req.params?.id,
                    userId: actualUserId,
                    deviceId
                }
            );
            MonitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (actualUserId) {
                MonitoringService?.error('Calendar event cancellation failed', {
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, actualUserId, deviceId);
            } else if (req.session?.id) {
                MonitoringService?.error('Calendar event cancellation failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    eventId: req.params?.id,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.cancelEvent.error', 1, { 
                errorId: mcpError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'calendar_cancel_error',
                error_description: 'Unable to cancel calendar event',
                errorId: mcpError.id
            });
        }
    },
    
    /**
     * POST /api/calendar/availability
     * Gets availability information for users within a specified time range
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async getAvailability(req, res) {
        // Extract user context from Express session (for web-based auth) or auth middleware (for device auth)
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/availability';
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('Processing calendar availability request', {
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    method: req.method,
                    path: req.path,
                    hasBody: !!req.body,
                    bodyKeys: req.body ? Object.keys(req.body) : [],
                    userId: actualUserId,
                    deviceId
                }, 'calendar');
            }
            
            // ENHANCED LOGGING: Log the raw request body for debugging
            MonitoringService?.debug('getAvailability raw request body:', { 
                body: req.body,
                bodyType: typeof req.body,
                hasTimeSlots: req.body && req.body.timeSlots ? 'yes' : 'no',
                hasUsers: req.body && req.body.users ? 'yes' : 'no',
                hasStart: req.body && req.body.start ? 'yes' : 'no',
                hasEnd: req.body && req.body.end ? 'yes' : 'no'
            }, 'calendar');
            
            // For simpler API calls, also support direct start/end parameters
            let requestBody = { ...req.body };
            if (requestBody.start && requestBody.end && requestBody.users) {
                MonitoringService?.info('Converting simplified availability request format to standard format', {
                    originalFormat: {
                        start: requestBody.start,
                        end: requestBody.end,
                        usersType: Array.isArray(requestBody.users) ? 'array' : typeof requestBody.users,
                        userCount: Array.isArray(requestBody.users) ? requestBody.users.length : (typeof requestBody.users === 'string' ? 1 : 0)
                    }
                }, 'calendar');
                requestBody = {
                    users: Array.isArray(requestBody.users) ? requestBody.users : [requestBody.users],
                    timeSlots: [{
                        start: { dateTime: requestBody.start, timeZone: 'UTC' },
                        end: { dateTime: requestBody.end, timeZone: 'UTC' }
                    }]
                };
                
                // Log the converted format
                MonitoringService?.debug('Converted to standard format:', {
                    convertedBody: requestBody
                }, 'calendar');
            }
            
            // Validate request body with standardized dateTime validation
            const availabilitySchema = Joi.object({
                users: Joi.array().items(Joi.string().email()).min(1).required(),
                timeSlots: Joi.array().items(
                    Joi.object({
                        start: Joi.object({
                            dateTime: Joi.date().iso().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required(),
                        end: Joi.object({
                            dateTime: Joi.date().iso().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required()
                    })
                ).required()
            });
            
            // Create a modified req object with our adjusted body for validation
            const modifiedReq = { 
                ...req,
                body: requestBody
            };
            
            const { error, value } = validateAndLog(modifiedReq, availabilitySchema, 'Get availability', { endpoint });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }

            // Create safe version of request data (redact email addresses for privacy)
            const userCount = value.users.length;
            const timeSlotCount = value.timeSlots.length;
            
            MonitoringService?.info('Getting availability data', { 
                userCount, 
                timeSlotCount,
                startTime: value.timeSlots[0]?.start?.dateTime,
                endTime: value.timeSlots[value.timeSlots.length - 1]?.end?.dateTime
            }, 'calendar');

            // Try to use the module's getAvailability method if available
            let availabilityData;
            try {
                MonitoringService?.info('Attempting to get real availability data from module', { 
                    userCount, 
                    timeSlotCount 
                }, 'calendar');
                
                const methodName = 'getAvailability';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    availabilityData = await calendarModule[methodName]({
                        users: value.users,
                        timeSlots: value.timeSlots
                    }, req);
                    
                    MonitoringService?.info(`Successfully retrieved real availability data for ${value.users.length} users`, { 
                        userCount: value.users.length,
                        resultSize: JSON.stringify(availabilityData).length
                    }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError(
                    'api',
                    'Failed to retrieve availability data from module',
                    'error',
                    { 
                        method: 'getAvailability', 
                        error: moduleError.message, 
                        stack: moduleError.stack,
                        userCount,
                        timeSlotCount
                    }
                );
                MonitoringService?.logError(moduleCallError);
                
                // Track availability error metric
                MonitoringService?.trackMetric('calendar.getAvailability.error', 1, {
                    errorId: moduleCallError.id,
                    reason: moduleError.message
                });
                
                // Generate mock availability data instead of returning an error
                MonitoringService?.info('Falling back to mock availability data', { userCount, timeSlotCount }, 'calendar');
                
                const mockAvailabilityData = {
                    users: value.users.map(user => ({
                        id: user,
                        availability: value.timeSlots.map(slot => ({
                            start: slot.start,
                            end: slot.end,
                            status: ['free', 'busy', 'tentative'][Math.floor(Math.random() * 3)]
                        }))
                    })),
                    isMock: true // Flag to indicate this is mock data
                };
                
                availabilityData = mockAvailabilityData;
                MonitoringService?.info('Generated mock availability data', { 
                    userCount: value.users.length,
                    resultSize: JSON.stringify(mockAvailabilityData).length
                }, 'calendar');
            }

            // Pattern 2: User Activity Log - Successful availability retrieval
            MonitoringService?.info('User successfully retrieved calendar availability', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                userCount: value.users.length,
                timeSlotCount: value.timeSlots.length,
                isMock: !!availabilityData.isMock,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track time to retrieve availability
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.getAvailability.duration', duration, { 
                userCount,
                timeSlotCount,
                isMock: !!availabilityData.isMock
            });

            // Send the response (real or mock data)
            res.json(availabilityData);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to retrieve calendar availability',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error retrieving calendar availability', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                errorType: 'availability_retrieval_failed',
                userMessage: 'Failed to retrieve calendar availability'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.getAvailability.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'AVAILABILITY_RETRIEVAL_FAILED',
                message: 'Unable to retrieve calendar availability at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    },
    
    
    /**
     * POST /api/calendar/findMeetingTimes
     * Find suitable meeting times for attendees
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async findMeetingTimes(req, res) {
        // Extract user context for logging and tracking
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        // Define endpoint at function scope so it's accessible in catch block
        const endpoint = '/api/calendar/findMeetingTimes';
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs (conditional on NODE_ENV)
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('findMeetingTimes request received', {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    requestBody: req.body,
                    contentType: req.headers['content-type'],
                    method: req.method,
                    endpoint,
                    userAgent: req.get('User-Agent')
                }, 'calendar');
            }
            
            // Validate request body with backward compatibility for Claude's format
            const optionsSchema = Joi.object({
                // Attendees - support both simple emails (Claude format) and objects (Graph API format)
                attendees: Joi.array().items(
                    Joi.alternatives().try(
                        // Simple email string (Claude format)
                        Joi.string().email(),
                        // Full Graph API format
                        Joi.object({
                            type: Joi.string().valid('required', 'optional', 'resource').default('required'),
                            emailAddress: Joi.object({
                                address: Joi.string().email().required(),
                                name: Joi.string().optional()
                            }).required()
                        })
                    )
                ).min(1).required(),
                
                // Time constraint - support both formats
                timeConstraint: Joi.object({
                    activityDomain: Joi.string().valid('work', 'personal', 'unrestricted').default('work'),
                    // Accept either timeSlots (capital S) or timeslots (lowercase s)
                    timeSlots: Joi.array().items(Joi.object({
                        start: Joi.object({
                            dateTime: Joi.string().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required(),
                        end: Joi.object({
                            dateTime: Joi.string().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required()
                    })).min(1).optional(),
                    // Also accept lowercase 'timeslots' as used by the Graph API
                    timeslots: Joi.array().items(Joi.object({
                        start: Joi.object({
                            dateTime: Joi.string().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required(),
                        end: Joi.object({
                            dateTime: Joi.string().required(),
                            timeZone: Joi.string().default('UTC')
                        }).required()
                    })).min(1).optional()
                }).optional(),
                
                // Alternative: timeConstraints (plural) for backward compatibility with Claude
                timeConstraints: Joi.object({
                    activityDomain: Joi.string().valid('work', 'personal', 'unrestricted').default('work'),
                    timeSlots: Joi.array().items(Joi.object({
                        start: Joi.object({
                            dateTime: Joi.string().required(), // Accept any string, let Graph API validate
                            timeZone: Joi.string().default('UTC')
                        }).required(),
                        end: Joi.object({
                            dateTime: Joi.string().required(), // Accept any string, let Graph API validate
                            timeZone: Joi.string().default('UTC')
                        }).required()
                    })).min(1).required()
                }).optional(),
                
                // Meeting duration in ISO8601 format (e.g., "PT1H" for 1 hour)
                meetingDuration: Joi.string().pattern(/^PT(\d+H)?(\d+M)?(\d+S)?$/).default('PT30M'),
                
                // Location constraint
                locationConstraint: Joi.object({
                    isRequired: Joi.boolean().default(false),
                    suggestLocation: Joi.boolean().default(false),
                    locations: Joi.array().items(Joi.object({
                        displayName: Joi.string().required(),
                        locationEmailAddress: Joi.string().email().optional(),
                        resolveAvailability: Joi.boolean().default(false)
                    })).optional()
                }).optional(),
                
                // Additional Graph API parameters
                maxCandidates: Joi.number().min(1).max(100).default(20),
                minimumAttendeePercentage: Joi.number().min(0).max(100).default(50),
                returnSuggestionReasons: Joi.boolean().default(true),
                isOrganizerOptional: Joi.boolean().default(false)
            }).or('timeConstraint', 'timeConstraints'); // Require at least one time constraint
            
            const { error, value } = validateAndLog(req, optionsSchema, 'Find meeting times', { endpoint });
            if (error) {
                // Enhanced error logging for debugging
                MonitoringService?.error('findMeetingTimes validation failed', {
                    validationError: error.details,
                    requestBody: req.body,
                    endpoint,
                    timestamp: new Date().toISOString()
                }, 'calendar');
                
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details,
                    message: error.details.map(d => d.message).join('; ')
                });
            }
            
            // Log successful validation
            MonitoringService?.info('findMeetingTimes validation succeeded', {
                validatedValue: value,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            // Normalize the data to Graph API format
            const normalizedValue = {
                ...value,
                // Convert simple email strings to Graph API attendee format
                attendees: value.attendees.map(attendee => {
                    if (typeof attendee === 'string') {
                        return {
                            type: 'required',
                            emailAddress: {
                                address: attendee,
                                name: attendee.split('@')[0] // Use part before @ as name
                            }
                        };
                    }
                    return attendee; // Already in correct format
                })
            };
            
            // Handle timeConstraint/timeConstraints with special attention to the timeslots/timeSlots format
            if (value.timeConstraint || value.timeConstraints) {
                const constraint = value.timeConstraint || value.timeConstraints;
                normalizedValue.timeConstraint = {
                    ...constraint,
                    // Ensure we use lowercase 'timeslots' as expected by the Graph API
                    timeslots: constraint.timeslots || constraint.timeSlots || []
                };
                
                // Remove the capitalized version if it exists to avoid confusion
                delete normalizedValue.timeConstraint.timeSlots;
            }
            
            // Remove the old timeConstraints field if it was used
            delete normalizedValue.timeConstraints;
            
            // Log normalized data for debugging
            MonitoringService?.debug('findMeetingTimes normalized data', {
                normalizedValue,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            // Create safe request data for logging (redact attendee details)
            const attendeeCount = normalizedValue.attendees?.length || 0;
            const timeConstraint = normalizedValue.timeConstraint;
            const firstTimeSlot = timeConstraint?.timeSlots?.[0];
            const lastTimeSlot = timeConstraint?.timeSlots?.[timeConstraint.timeSlots.length - 1];
            
            MonitoringService?.info('Finding meeting times', { 
                attendeeCount,
                activityDomain: timeConstraint?.activityDomain,
                timeSlotsCount: timeConstraint?.timeSlots?.length || 0,
                startTime: firstTimeSlot?.start?.dateTime,
                endTime: lastTimeSlot?.end?.dateTime,
                meetingDuration: normalizedValue.meetingDuration,
                maxCandidates: normalizedValue.maxCandidates
            }, 'calendar');

            // Try to use the module's findMeetingTimes method if available
            let suggestions;
            let isMock = false;
            try {
                // Extract meeting duration from normalized value for logging
                const meetingDuration = normalizedValue.meetingDuration || 'PT30M';
                
                MonitoringService?.info('Attempting to find meeting times using module', {
                    attendeeCount,
                    meetingDuration
                }, 'calendar');
                
                const methodName = 'findMeetingTimes';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    // Pass req object for user-scoped token selection
                    suggestions = await calendarModule[methodName](normalizedValue, req);
                    MonitoringService?.info('Successfully found meeting times using module', {
                        suggestionCount: suggestions.meetingTimeSuggestions?.length || 0
                    }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', 'Error finding meeting times', 'error', { 
                    error: moduleError.message,
                    stack: moduleError.stack,
                    attendeeCount
                });
                MonitoringService?.logError(moduleCallError);
                
                // Temporarily disable fallback to see the actual error
                throw moduleError;
            }
            
            // Pattern 2: User Activity Log - Successful meeting times found
            MonitoringService?.info('User successfully found meeting times', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                attendeeCount: normalizedValue.attendees?.length || 0,
                suggestionCount: suggestions.meetingTimeSuggestions?.length || 0,
                meetingDuration: normalizedValue.meetingDuration,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track find meeting times duration
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.findMeetingTimes.duration', duration, { 
                suggestionCount: suggestions.meetingTimeSuggestions?.length || 0,
                attendeeCount,
                isMock
            });
            
            res.json(suggestions);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to find meeting times',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error finding meeting times', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                errorType: 'meeting_times_search_failed',
                userMessage: 'Failed to find meeting times'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.findMeetingTimes.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'MEETING_TIMES_SEARCH_FAILED',
                message: 'Unable to find meeting times at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    },
    
    /**
     * GET /api/calendar/rooms
     * Get available rooms for meetings
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async getRooms(req, res) {
        // Extract user context for logging and tracking
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/rooms';
            
            // Pattern 1: Development Debug Logs (conditional on NODE_ENV)
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('getRooms request received', {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    query: req.query,
                    method: req.method,
                    endpoint,
                    userAgent: req.get('User-Agent')
                }, 'calendar');
            }
            
            // Validate query parameters
            const querySchema = Joi.object({
                building: Joi.string().optional(),
                capacity: Joi.number().integer().min(1).optional(),
                hasAudio: Joi.boolean().optional(),
                hasVideo: Joi.boolean().optional(),
                floor: Joi.number().integer().optional(),
                limit: Joi.number().integer().min(1).max(100).default(50).optional()
            });
            
            const { error, value } = querySchema.validate(req.query);
            if (error) {
                const validationError = ErrorService?.createError('api', 'Room query validation error', 'warning', { 
                    details: error.details,
                    endpoint
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ 
                    error: 'Invalid query parameters', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Getting available meeting rooms', { query: value }, 'calendar');
            
            // Try to use the module's getRooms method if available
            let result;
            let isMock = false;
            try {
                MonitoringService?.info('Attempting to get rooms using module', {}, 'calendar');
                
                const methodName = 'getRooms';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    result = await calendarModule[methodName](value, req);
                    MonitoringService?.info('Successfully got rooms using module', { count: result.rooms?.length || 0 }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', 'Error getting meeting rooms', 'error', { 
                    error: moduleError.message,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock rooms', {}, 'calendar');
                
                // If module method fails, create mock rooms
                const mockRooms = [
                    {
                        id: 'room1',
                        displayName: 'Conference Room A',
                        emailAddress: 'room.a@example.com',
                        capacity: 10,
                        building: 'Building 1',
                        floorNumber: 2,
                        hasAudio: true,
                        hasVideo: true,
                        isMock: true
                    },
                    {
                        id: 'room2',
                        displayName: 'Conference Room B',
                        emailAddress: 'room.b@example.com',
                        capacity: 6,
                        building: 'Building 1',
                        floorNumber: 3,
                        hasAudio: true,
                        hasVideo: false,
                        isMock: true
                    },
                    {
                        id: 'room3',
                        displayName: 'Executive Boardroom',
                        emailAddress: 'boardroom@example.com',
                        capacity: 20,
                        building: 'Building 2',
                        floorNumber: 5,
                        hasAudio: true,
                        hasVideo: true,
                        isMock: true
                    }
                ];
                
                // Filter mock rooms based on query parameters
                let filteredRooms = [...mockRooms];
                
                if (value.building) {
                    filteredRooms = filteredRooms.filter(room => 
                        room.building.toLowerCase() === value.building.toLowerCase());
                }
                
                if (value.capacity) {
                    filteredRooms = filteredRooms.filter(room => room.capacity >= value.capacity);
                }
                
                if (value.hasAudio !== undefined) {
                    filteredRooms = filteredRooms.filter(room => room.hasAudio === value.hasAudio);
                }
                
                if (value.hasVideo !== undefined) {
                    filteredRooms = filteredRooms.filter(room => room.hasVideo === value.hasVideo);
                }
                
                if (value.floor) {
                    filteredRooms = filteredRooms.filter(room => room.floorNumber === value.floor);
                }
                
                // Apply limit if specified
                if (value.limit && filteredRooms.length > value.limit) {
                    filteredRooms = filteredRooms.slice(0, value.limit);
                }
                
                result = {
                    rooms: filteredRooms,
                    nextLink: null
                };
                isMock = true;
                
                MonitoringService?.info('Generated mock rooms', { count: filteredRooms.length }, 'calendar');
            }
            
            // Extract rooms array from result
            const rooms = result.rooms || [];
            
            // Pattern 2: User Activity Log - Successful rooms retrieval
            MonitoringService?.info('User successfully retrieved meeting rooms', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                roomCount: rooms.length,
                isMock,
                queryFilters: Object.keys(value).length,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track room retrieval time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.getRooms.duration', duration, { 
                count: rooms.length,
                isMock
            });
            
            res.json({ 
                rooms: rooms,
                nextLink: result.nextLink || null,
                fromCache: result.fromCache || false
            });
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to retrieve meeting rooms',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error retrieving meeting rooms', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                errorType: 'rooms_retrieval_failed',
                userMessage: 'Failed to retrieve meeting rooms'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.getRooms.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'ROOMS_RETRIEVAL_FAILED',
                message: 'Unable to retrieve meeting rooms at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    },
    
    /**
     * GET /api/calendar/calendars
     * Get user calendars
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async getCalendars(req, res) {
        // Extract user context for logging and tracking
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/calendars';
            
            // Pattern 1: Development Debug Logs (conditional on NODE_ENV)
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('getCalendars request received', {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    query: req.query,
                    method: req.method,
                    endpoint,
                    userAgent: req.get('User-Agent')
                }, 'calendar');
            }
            
            // Validate query parameters
            const querySchema = Joi.object({
                includeShared: Joi.boolean().default(true).optional()
            });
            
            const { error, value } = querySchema.validate(req.query);
            if (error) {
                const validationError = ErrorService?.createError('api', 'Calendar query validation error', 'warning', { 
                    details: error.details,
                    endpoint
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ 
                    error: 'Invalid query parameters', 
                    details: error.details 
                });
            }
            
            MonitoringService?.info('Getting user calendars', value, 'calendar');
            
            // Try to use the module's getCalendars method if available
            let calendars;
            let isMock = false;
            try {
                MonitoringService?.info('Attempting to get calendars using module', {}, 'calendar');
                
                const methodName = 'getCalendars';
                
                if (isModuleMethodAvailable(methodName, calendarModule)) {
                    calendars = await calendarModule[methodName](value, req);
                    MonitoringService?.info('Successfully got calendars using module', { count: calendars.length }, 'calendar');
                } else {
                    throw new Error(`calendarModule.${methodName} is not implemented`);
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', 'Error getting user calendars', 'error', { 
                    error: moduleError.message,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock calendars', {}, 'calendar');
                
                // If module method fails, create mock calendars
                calendars = [
                    {
                        id: 'calendar1',
                        name: 'Calendar',
                        color: 'auto',
                        isDefaultCalendar: true,
                        canShare: true,
                        canViewPrivateItems: true,
                        canEdit: true,
                        owner: {
                            name: 'Current User',
                            address: 'current.user@example.com'
                        },
                        isMock: true
                    },
                    {
                        id: 'calendar2',
                        name: 'Birthdays',
                        color: 'lightBlue',
                        isDefaultCalendar: false,
                        canShare: false,
                        canViewPrivateItems: true,
                        canEdit: false,
                        owner: {
                            name: 'Current User',
                            address: 'current.user@example.com'
                        },
                        isMock: true
                    },
                    {
                        id: 'calendar3',
                        name: 'Holidays',
                        color: 'lightGreen',
                        isDefaultCalendar: false,
                        canShare: false,
                        canViewPrivateItems: true,
                        canEdit: false,
                        owner: {
                            name: 'Current User',
                            address: 'current.user@example.com'
                        },
                        isMock: true
                    }
                ];
                isMock = true;
                
                // If includeShared is false, filter out shared calendars
                if (value.includeShared === false) {
                    calendars = calendars.filter(calendar => 
                        calendar.owner.address === 'current.user@example.com' || 
                        calendar.isDefaultCalendar === true);
                }
                
                MonitoringService?.info('Generated mock calendars', { count: calendars.length }, 'calendar');
            }
            
            // Pattern 2: User Activity Log - Successful calendars retrieval
            MonitoringService?.info('User successfully retrieved calendars', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                calendarCount: calendars.length,
                includeShared: value.includeShared,
                isMock,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track calendar retrieval time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.getCalendars.duration', duration, { 
                count: calendars.length,
                isMock
            });
            
            res.json(calendars);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to retrieve user calendars',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error retrieving calendars', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                errorType: 'calendars_retrieval_failed',
                userMessage: 'Failed to retrieve calendars'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.getCalendars.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'CALENDARS_RETRIEVAL_FAILED',
                message: 'Unable to retrieve calendars at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    },
    
    /**
     * POST /api/calendar/events/:id/attachments
     * Add an attachment to an event
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async addAttachment(req, res) {
        // Extract user context for logging and tracking
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            const endpoint = '/api/calendar/events/:id/attachments';
            
            // Pattern 1: Development Debug Logs (conditional on NODE_ENV)
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('addAttachment request received', {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    eventId: req.params.id,
                    requestBody: req.body,
                    method: req.method,
                    endpoint,
                    userAgent: req.get('User-Agent')
                }, 'calendar');
            }
            
            // Get event ID from URL parameters
            const eventId = req.params.id;
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for adding attachment', 'warning', { 
                    endpoint 
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            // Validate request body
            const attachmentSchema = Joi.object({
                name: Joi.string().required(),
                contentType: Joi.string().required(),
                contentBytes: Joi.string().required(), // Base64 encoded content
                isInline: Joi.boolean().default(false)
            });
            
            const { error, value } = validateAndLog(req, attachmentSchema, 'Add attachment', { eventId, endpoint });
            if (error) {
                return res.status(400).json({ 
                    error: 'Invalid request', 
                    details: error.details 
                });
            }
            
            // Create safe version of request body for logging
            const safeAttachmentInfo = {
                name: value.name,
                contentType: value.contentType,
                contentSize: value.contentBytes ? value.contentBytes.length : 0,
                isInline: value.isInline || false
            };
            
            MonitoringService?.info('Adding attachment to calendar event', { 
                eventId,
                attachment: safeAttachmentInfo 
            }, 'calendar');
            
            // Try to use the module's handleIntent method for addAttachment
            let attachment;
            try {
                MonitoringService?.info(`Attempting to add attachment to event ${eventId} using intent handler`, { 
                    eventId,
                    attachmentName: value.name
                }, 'calendar');
                
                if (isModuleMethodAvailable('handleIntent', calendarModule)) {
                    const entities = {
                        id: eventId,
                        name: value.name,
                        contentBytes: value.contentBytes,
                        contentType: value.contentType
                    };
                    const result = await calendarModule.handleIntent('addAttachment', entities, { req });
                    attachment = result.attachment;
                    MonitoringService?.info(`Successfully added attachment to event ${eventId} using intent handler`, { 
                        eventId,
                        attachmentId: attachment.id
                    }, 'calendar');
                } else {
                    throw new Error('calendarModule.handleIntent is not implemented');
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error adding attachment to event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                MonitoringService?.info('Falling back to mock attachment', { eventId }, 'calendar');
                
                // If module method fails, create a mock attachment
                attachment = {
                    id: 'attachment-' + Date.now(),
                    name: value.name,
                    contentType: value.contentType,
                    size: value.contentBytes.length * 0.75, // Approximate size after base64 decoding
                    isInline: value.isInline || false,
                    lastModifiedDateTime: new Date().toISOString(),
                    isMock: true // Flag to indicate this is mock data
                };
                MonitoringService?.info('Generated mock attachment', { 
                    eventId,
                    attachmentId: attachment.id
                }, 'calendar');
            }
            
            // Pattern 2: User Activity Log - Successful attachment added
            MonitoringService?.info('User successfully added attachment to calendar event', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                eventId: req.params.id,
                attachmentId: attachment.id,
                attachmentName: attachment.name,
                contentType: attachment.contentType,
                isMock: !!attachment.isMock,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track add attachment time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.addAttachment.duration', duration, { 
                eventId,
                attachmentId: attachment.id,
                contentType: attachment.contentType,
                isMock: !!attachment.isMock
            });
            
            res.json(attachment);
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to add attachment to calendar event',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    eventId: req.params?.id,
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error adding attachment to calendar event', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                eventId: req.params?.id,
                errorType: 'attachment_add_failed',
                userMessage: 'Failed to add attachment to calendar event'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.addAttachment.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'ATTACHMENT_ADD_FAILED',
                message: 'Unable to add attachment to calendar event at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    },
    
    /**
     * DELETE /api/calendar/events/:id/attachments/:attachmentId
     * Remove an attachment from an event
     * @param {import('express').Request} req
     * @param {import('express').Response} res
     */
    async removeAttachment(req, res) {
        // Extract user context for logging and tracking
        const { userId = null, deviceId = null } = req.user || {};
        const sessionUserId = req.session?.id ? `user:${req.session.id}` : null;
        const actualUserId = userId || sessionUserId;
        
        // Define endpoint at function scope so it's accessible in catch block
        const endpoint = '/api/calendar/events/:id/attachments/:attachmentId';
        
        try {
            // Start timing for performance tracking
            const startTime = Date.now();
            
            // Pattern 1: Development Debug Logs (conditional on NODE_ENV)
            if (process.env.NODE_ENV === 'development') {
                MonitoringService?.debug('removeAttachment request received', {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    eventId: req.params.id,
                    attachmentId: req.params.attachmentId,
                    method: req.method,
                    endpoint,
                    userAgent: req.get('User-Agent')
                }, 'calendar');
            }
            
            // Get event ID and attachment ID from URL parameters
            const eventId = req.params.id;
            const attachmentId = req.params.attachmentId;
            
            if (!eventId) {
                const validationError = ErrorService?.createError('api', 'Event ID is required for removing attachment', 'warning', { 
                    endpoint 
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Event ID is required' });
            }
            
            if (!attachmentId) {
                const validationError = ErrorService?.createError('api', 'Attachment ID is required for removing attachment', 'warning', { 
                    endpoint,
                    eventId
                });
                MonitoringService?.logError(validationError);
                return res.status(400).json({ error: 'Attachment ID is required' });
            }
            
            MonitoringService?.info('Removing attachment from calendar event', { 
                eventId,
                attachmentId 
            }, 'calendar');
            
            // Try to use the module's handleIntent method for removeAttachment
            let result;
            try {
                MonitoringService?.info(`Attempting to remove attachment ${attachmentId} from event ${eventId} using intent handler`, { 
                    eventId,
                    attachmentId
                }, 'calendar');
                
                if (isModuleMethodAvailable('handleIntent', calendarModule)) {
                    const entities = {
                        eventId: eventId,
                        attachmentId: attachmentId
                    };
                    const intentResult = await calendarModule.handleIntent('removeAttachment', entities, { req });
                    result = intentResult.success;
                    MonitoringService?.info(`Successfully removed attachment ${attachmentId} from event ${eventId} using intent handler`, { 
                        eventId,
                        attachmentId,
                        success: !!result
                    }, 'calendar');
                } else {
                    throw new Error('calendarModule.handleIntent is not implemented');
                }
            } catch (moduleError) {
                const moduleCallError = ErrorService?.createError('api', `Error removing attachment ${attachmentId} from event ${eventId}`, 'error', { 
                    error: moduleError.message,
                    eventId,
                    attachmentId,
                    stack: moduleError.stack
                });
                MonitoringService?.logError(moduleCallError);
                
                // Temporarily disable fallback to see the actual error
                throw moduleError;
            }
            
            // Pattern 2: User Activity Log - Successful attachment removed
            MonitoringService?.info('User successfully removed attachment from calendar event', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                eventId: req.params.id,
                attachmentId: req.params.attachmentId,
                success: typeof result === 'object' ? result.success : !!result,
                isMock: typeof result === 'object' ? !!result.isMock : false,
                duration: Date.now() - startTime
            }, 'calendar');

            // Track remove attachment time
            const duration = Date.now() - startTime;
            MonitoringService?.trackMetric('calendar.removeAttachment.duration', duration, { 
                eventId,
                attachmentId,
                success: typeof result === 'object' ? result.success : !!result,
                isMock: typeof result === 'object' ? !!result.isMock : false
            });
            
            res.json(typeof result === 'object' ? result : { success: result });
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const infrastructureError = ErrorService?.createError(
                'calendar',
                'Failed to remove attachment from calendar event',
                'error',
                {
                    userId: actualUserId,
                    deviceId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString(),
                    eventId: req.params?.id,
                    attachmentId: req.params?.attachmentId,
                    error: err.message,
                    stack: err.stack,
                    endpoint: req.originalUrl,
                    method: req.method,
                    userAgent: req.get('User-Agent')
                }
            );
            MonitoringService?.logError(infrastructureError);
            
            // Pattern 4: User Error Tracking
            MonitoringService?.warn('User encountered error removing attachment from calendar event', {
                userId: actualUserId,
                deviceId,
                sessionId: req.session?.id,
                timestamp: new Date().toISOString(),
                eventId: req.params?.id,
                attachmentId: req.params?.attachmentId,
                errorType: 'attachment_remove_failed',
                userMessage: 'Failed to remove attachment from calendar event'
            }, 'calendar');
            
            // Track error metric
            MonitoringService?.trackMetric('calendar.removeAttachment.error', 1, { 
                errorId: infrastructureError.id,
                reason: err.message
            });
            
            res.status(500).json({ 
                error: 'ATTACHMENT_REMOVE_FAILED',
                message: 'Unable to remove attachment from calendar event at this time',
                details: 'Please try again later or contact support if the issue persists'
            });
        }
    }
});
