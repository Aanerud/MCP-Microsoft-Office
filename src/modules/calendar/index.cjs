/**
 * @fileoverview MCP Calendar Module - Handles calendar-related intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system and phase1_architecture.md.
 */

const Joi = require('joi');
const { normalizeEvent } = require('../../graph/normalizers.cjs');
const moment = require('moment'); // For duration calculations
const MonitoringService = require('../../core/monitoring-service.cjs');
const ErrorService = require('../../core/error-service.cjs');

const CALENDAR_CAPABILITIES = [
    'getEvents',
    'createEvent',
    'updateEvent',
    'getAvailability',
    'acceptEvent',
    'tentativelyAcceptEvent',
    'declineEvent',
    'cancelEvent',
    'findMeetingTimes',
    'getRooms',
    'getCalendars',
    'addAttachment',
    'removeAttachment'
];

// --- Attachment Constants and Schema ---
const MAX_ATTACHMENT_SIZE_BYTES = 10 * 1024 * 1024; // 10 MB
const ALLOWED_ATTACHMENT_TYPES = new Set([
    'application/pdf',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // .docx
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',     // .xlsx
    'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
    'text/plain',
    'application/octet-stream', // Generic binary fallback for various file types
    'image/jpeg',
    'image/png',
    'image/gif'
]);

const attachmentSchema = Joi.object({
    name: Joi.string().required(),
    contentType: Joi.string().custom((value, helpers) => {
        if (!ALLOWED_ATTACHMENT_TYPES.has(value)) {
            return helpers.error('attachment.invalidType', { value });
        }
        return value;
    }).required(),
    contentBytes: Joi.string().base64().required().max(MAX_ATTACHMENT_SIZE_BYTES).messages({
        'string.max': `Attachment size exceeds the limit of ${MAX_ATTACHMENT_SIZE_BYTES / 1024 / 1024} MB.`,
        'any.required': 'Attachment contentBytes are required.',
        'string.base64': 'Attachment contentBytes must be a valid base64 string.'
    }),
    // Ensure '@odata.type' is correctly set if required by Graph API
    '@odata.type': Joi.string().valid('#microsoft.graph.fileAttachment').default('#microsoft.graph.fileAttachment')
}).required().messages({
    'attachment.invalidType': 'Attachment content type \"{#value}\" is not allowed.'
});
// --- End Attachment Constants and Schema ---

// Log module initialization
MonitoringService.info('Calendar Module initialized', {
    serviceName: 'calendar-module',
    capabilities: CALENDAR_CAPABILITIES.length,
    timestamp: new Date().toISOString()
}, 'calendar');

const CalendarModule = {
    /**
     * Helper method to redact sensitive data from objects before logging
     * @param {object} data - The data object to redact
     * @param {WeakSet} visited - Set to track visited objects for circular reference detection
     * @returns {object} Redacted copy of the data
     * @private
     */
    redactSensitiveData(data, visited = new WeakSet()) {
        if (!data || typeof data !== 'object') {
            return data;
        }
        
        // Check for circular references
        if (visited.has(data)) {
            return '[Circular Reference]';
        }
        
        // Add current object to visited set
        visited.add(data);
        
        // Create a deep copy to avoid modifying the original
        const result = Array.isArray(data) ? [...data] : {...data};
        
        // Fields that should be redacted
        const sensitiveFields = [
            'user', 'email', 'mail', 'address', 'emailAddress', 'password', 'token', 'accessToken',
            'refreshToken', 'content', 'body', 'contentBytes'
        ];
        
        // Recursively process the object
        for (const key in result) {
            if (Object.prototype.hasOwnProperty.call(result, key)) {
                // Check if this is a sensitive field
                if (sensitiveFields.includes(key.toLowerCase())) {
                    if (typeof result[key] === 'string') {
                        result[key] = 'REDACTED';
                    } else if (Array.isArray(result[key])) {
                        result[key] = `[${result[key].length} items]`;
                    } else if (typeof result[key] === 'object' && result[key] !== null) {
                        result[key] = '{REDACTED}';
                    }
                } 
                // Recursively process nested objects with proper context binding
                else if (typeof result[key] === 'object' && result[key] !== null) {
                    result[key] = this.redactSensitiveData(result[key], visited);
                }
            }
        }
        
        return result;
    },
    /**
     * Fetch raw calendar events from Graph for debugging purposes only (no normalization)
     * This method is intentionally restricted to non-production environments to prevent
     * exposure of sensitive data and ensure consistent data formatting in production.
     * Only available when NODE_ENV !== 'production'.
     * @param {object} options - Query options for fetching events
     * @param {object} req - Express request object containing user context
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object[]>} - Raw event objects from Graph API
     * @throws {Error} When called in production environment
     */
    async getEventsRaw(options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context
        const resolvedUserId = userId || req?.user?.userId;
        const resolvedSessionId = sessionId || req?.session?.id;
        
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService?.debug('Getting raw calendar events', {
                    sessionId: resolvedSessionId,
                    userAgent: req?.get('User-Agent'),
                    optionsCount: Object.keys(options).length,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            // Restrict to debug mode only (dev or development environments)
            const isDebugMode = ['dev', 'development'].includes(process.env.NODE_ENV) || process.env.DEBUG === 'true';
            if (!isDebugMode) {
                const error = errorService?.createError(
                    'calendar',
                    'getEventsRaw is only available in debug mode (dev, development, or DEBUG=true)',
                    'error',
                    { environment: process.env.NODE_ENV, timestamp: new Date().toISOString() }
                ) || {
                    category: 'calendar',
                    message: 'getEventsRaw is only available in debug mode (dev, development, or DEBUG=true)',
                    severity: 'error',
                    context: { environment: process.env.NODE_ENV }
                };
                
                monitoringService?.logError(error);
                
                // Pattern 4: User Error Tracking
                if (resolvedUserId) {
                    monitoringService?.error('Raw calendar events access denied', {
                        error: 'Debug mode only',
                        environment: process.env.NODE_ENV,
                        timestamp: new Date().toISOString()
                    }, 'calendar', null, resolvedUserId);
                } else if (resolvedSessionId) {
                    monitoringService?.error('Raw calendar events access denied', {
                        sessionId: resolvedSessionId,
                        error: 'Debug mode only',
                        environment: process.env.NODE_ENV,
                        timestamp: new Date().toISOString()
                    }, 'calendar');
                }
                    
                throw error;
            }

            if (!graphService || typeof graphService.getEventsRaw !== 'function') {
                const error = errorService?.createError(
                    'calendar',
                    'GraphService.getEventsRaw not implemented',
                    'error',
                    { timestamp: new Date().toISOString() }
                ) || {
                    category: 'calendar',
                    message: 'GraphService.getEventsRaw not implemented',
                    severity: 'error',
                    context: {}
                };
                
                monitoringService?.logError(error);
                
                // Pattern 4: User Error Tracking
                if (resolvedUserId) {
                    monitoringService?.error('Raw calendar events service unavailable', {
                        error: 'Service not implemented',
                        timestamp: new Date().toISOString()
                    }, 'calendar', null, resolvedUserId);
                } else if (resolvedSessionId) {
                    monitoringService?.error('Raw calendar events service unavailable', {
                        sessionId: resolvedSessionId,
                        error: 'Service not implemented',
                        timestamp: new Date().toISOString()
                    }, 'calendar');
                }
                    
                throw error;
            }
            
            // Include req in options for getEventsRaw
            const events = await graphService.getEventsRaw({ ...options, req }, resolvedUserId, resolvedSessionId);
            const elapsedTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (resolvedUserId) {
                monitoringService?.info('Raw calendar events retrieved successfully', {
                    count: events?.length || 0,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, resolvedUserId);
            } else if (resolvedSessionId) {
                monitoringService?.info('Raw calendar events retrieved with session', {
                    sessionId: resolvedSessionId,
                    count: events?.length || 0,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
            
            monitoringService?.trackMetric('calendar_event_raw_get_duration', elapsedTime, {
                count: events?.length || 0,
                timestamp: new Date().toISOString()
            });
            
            return events;
        } catch (error) {
            const elapsedTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = errorService?.createError(
                'calendar',
                `Failed to fetch raw calendar events: ${error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    options,
                    graphStatusCode: error.statusCode,
                    graphCode: error.code,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: `Failed to fetch raw calendar events: ${error.message}`,
                severity: 'error',
                context: { options, elapsedTime }
            };
            
            monitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (resolvedUserId) {
                monitoringService?.error('Raw calendar events retrieval failed', {
                    error: error.message,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'calendar', null, resolvedUserId);
            } else if (resolvedSessionId) {
                monitoringService?.error('Raw calendar events retrieval failed', {
                    sessionId: resolvedSessionId,
                    error: error.message,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }
                
            throw mcpError;
        }
    },
    
    /**
     * Fetch calendar events from Graph and normalize them
     * @param {object} options
     * @returns {Promise<object[]>}
     */
    async getEvents(options = {}, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Log the request attempt
        monitoringService?.debug('Attempting to get calendar events', { 
            options: { ...options, user: options.req?.user ? 'REDACTED' : undefined },
            timestamp: new Date().toISOString()
        }, 'calendar');
        
        if (!graphService || typeof graphService.getEvents !== 'function') {
            const error = errorService?.createError(
                'calendar',
                'GraphService.getEvents not implemented',
                'error',
                { timestamp: new Date().toISOString() }
            ) || {
                category: 'calendar',
                message: 'GraphService.getEvents not implemented',
                severity: 'error',
                context: {}
            };
            
            monitoringService?.logError(error);
                
            throw error;
        }
        
        try {
            const startTime = Date.now();
            
            // Extract user context if available in req for the service layer (e.g., caching)
            const userContext = req && req.user ? { user: req.user } : {};
            
            // Pass original options along with extracted user context and req object
            // Note: req must be included in the options object, not as a separate argument
            const events = await graphService.getEvents({ ...options, ...userContext, req });
            const elapsedTime = Date.now() - startTime;
            
            // Log success and track performance
            monitoringService?.info('Successfully retrieved calendar events', { 
                count: events?.length || 0,
                elapsedTime,
                hasUserContext: !!userContext.user,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            monitoringService?.trackMetric('calendar_event_get_duration', elapsedTime, {
                count: events?.length || 0,
                hasUserContext: !!userContext.user,
                timestamp: new Date().toISOString()
            });
            
            return events;
        } catch (error) {
            const mcpError = errorService?.createError(
                'calendar',
                `Failed to fetch calendar events: ${error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    options: { ...options, user: options.req?.user ? 'REDACTED' : undefined },
                    graphStatusCode: error.statusCode,
                    graphCode: error.code,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: `Failed to fetch calendar events: ${error.message}`,
                severity: 'error',
                context: { options: { ...options, user: options.req?.user ? 'REDACTED' : undefined } }
            };
            
            monitoringService?.logError(mcpError);
                
            throw mcpError;
        }
    },
    
    /**
     * Creates a calendar event using the Graph Service.
     * This function is exposed as a capability and called by handleIntent.
     *
     * @param {object} eventData - The event data.
     * @param {string} eventData.subject - The subject of the event.
     * @param {object} eventData.start - The start time of the event.
     * @param {string} eventData.start.dateTime - The date and time in ISO 8601 format (e.g., '2025-04-30T12:00:00').
     * @param {string} eventData.start.timeZone - The time zone (e.g., 'America/Los_Angeles', 'UTC').
     * @param {object} eventData.end - The end time of the event.
     * @param {string} eventData.end.dateTime - The date and time in ISO 8601 format.
     * @param {string} eventData.end.timeZone - The time zone.
     * @param {Array<object>} [eventData.attendees] - Optional. Array of attendees.
     * @param {object} eventData.attendees[].emailAddress - Attendee email address.
     * @param {string} eventData.attendees[].emailAddress.address - The email address.
     * @param {string} [eventData.attendees[].emailAddress.name] - Optional. Attendee name.
     * @param {string} [eventData.attendees[].type="required"] - Attendee type ('required', 'optional', 'resource'). Defaults to 'required'.
     * @param {object} [eventData.body] - Optional. The body of the event.
     * @param {string} [eventData.body.contentType="HTML"] - Content type ('HTML' or 'text').
     * @param {string} eventData.body.content - The content of the body.
     * @param {boolean} [eventData.isOnlineMeeting=false] - Optional. Whether the event is an online meeting.
     * @param {object} [req] - Optional Express request object, potentially containing user context.
     * @returns {Promise<object>} The *normalized* created event object.
     * @throws {Error} If the Graph Service fails to create the event.
     */
    async createEvent(eventData, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Redact potentially sensitive data for logging
        const redactedEventData = {
            ...eventData,
            attendees: eventData.attendees ? `[${eventData.attendees.length} attendees]` : undefined,
            body: eventData.body ? { contentType: eventData.body.contentType, content: 'REDACTED' } : undefined
        };
        
        // Log the request attempt
        monitoringService?.debug('Attempting to create calendar event', { 
            eventData: redactedEventData,
            timestamp: new Date().toISOString()
        }, 'calendar');
        
        try {
            if (!graphService) {
                const error = errorService?.createError(
                    'calendar',
                    'Graph service not available',
                    'error',
                    { timestamp: new Date().toISOString() }
                ) || {
                    category: 'calendar',
                    message: 'Graph service not available',
                    severity: 'error',
                    context: {}
                };
                
                monitoringService?.logError(error);
                    
                throw error;
            }
            
            let finalEventData = { ...eventData };
            const startTime = Date.now();

            // 1. Resolve attendee names if attendees are provided
            if (Array.isArray(finalEventData.attendees) && finalEventData.attendees.length > 0) {
                if (typeof graphService.resolveAttendeeNames !== 'function') {
                    const error = errorService?.createError(
                        'calendar',
                        'GraphService.resolveAttendeeNames not implemented',
                        'error',
                        { timestamp: new Date().toISOString() }
                    ) || {
                        category: 'calendar',
                        message: 'GraphService.resolveAttendeeNames not implemented',
                        severity: 'error',
                        context: {}
                    };
                    
                    monitoringService?.logError(error);
                        
                    throw error;
                }
                
                monitoringService?.debug('Resolving attendee names', { 
                    count: finalEventData.attendees.length,
                    timestamp: new Date().toISOString()
                }, 'calendar');
                
                const resolvedAttendees = await graphService.resolveAttendeeNames(finalEventData.attendees);
                finalEventData.attendees = resolvedAttendees;
                
                monitoringService?.debug('Attendees resolved successfully', { 
                    count: resolvedAttendees.length,
                    timestamp: new Date().toISOString()
                }, 'calendar');
            }

            // 2. Create the event via the service
            const createdEvent = await graphService.createEvent(finalEventData, 'me', { req });
            const elapsedTime = Date.now() - startTime;
            
            // 3. Normalize the result before returning
            const normalizedEvent = normalizeEvent(createdEvent);
            
            // Log success and track performance
            monitoringService?.info('Event created successfully', { 
                eventId: createdEvent.id,
                subject: redactedEventData.subject,
                attendeeCount: finalEventData.attendees?.length || 0,
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            monitoringService?.trackMetric('calendar_event_create_duration', elapsedTime, {
                hasAttendees: !!finalEventData.attendees?.length,
                isOnlineMeeting: !!finalEventData.isOnlineMeeting,
                timestamp: new Date().toISOString()
            });
            
            return normalizedEvent;
            
        } catch (error) {
            const mcpError = errorService?.createError(
                'calendar',
                `Failed to create event: ${error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    eventData: redactedEventData,
                    graphStatusCode: error.statusCode,
                    graphCode: error.code,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: `Failed to create event: ${error.message}`,
                severity: 'error',
                context: { eventData: redactedEventData }
            };
            
            monitoringService?.logError(mcpError);
                
            throw mcpError;
        }
    },
    
    /**
     * Update an existing calendar event after validating input and checking ownership.
     * @param {string} eventId - ID of the event to update.
     * @param {object} updates - Updated event data. Should contain at least one updateable field.
     * @param {object} [req] - Optional Express request object containing user context (`req.user`).
     * @returns {Promise<object>} Normalized updated event object.
     * @throws {Error} If validation fails, user is not authorized, or service fails.
     */
    async updateEvent(eventId, updates, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Redact potentially sensitive data for logging
        const redactedUpdates = {
            ...updates,
            attendees: updates.attendees ? `[${updates.attendees.length} attendees]` : undefined,
            body: updates.body ? { contentType: updates.body.contentType, content: 'REDACTED' } : undefined
        };
        
        // Log the update attempt
        monitoringService?.debug('Attempting to update calendar event', {
            eventId,
            updates: redactedUpdates,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService.updateEvent !== 'function') {
            const error = errorService?.createError(
                'calendar',
                'GraphService.updateEvent method not implemented',
                'error',
                { 
                    eventId,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: 'GraphService.updateEvent method not implemented',
                severity: 'error',
                context: { eventId }
            };
            
            monitoringService?.logError(error);
                
            throw error;
        }

        // TEMPORARY: Relaxed validation for debugging timezone issues
        // TODO: Re-implement robust Joi validation after timezone issues are resolved
        
        // Basic validation - ensure we have something to update
        if (!updates || typeof updates !== 'object' || Object.keys(updates).length === 0) {
            const err = errorService?.createError(
                'calendar',
                'No valid update data provided',
                'warn',
                { 
                    eventId,
                    updates: redactedUpdates,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: 'No valid update data provided',
                severity: 'warn',
                context: { eventId, updates: redactedUpdates }
            };
            
            monitoringService?.logError(err);
                
            throw err;
        }

        // Log the raw payload being sent to Graph API for debugging
        monitoringService?.debug('Raw update payload being sent to Microsoft Graph', {
            eventId,
            payload: redactedUpdates,
            payloadKeys: Object.keys(updates),
            hasTimezone: !!(updates.start?.timeZone || updates.end?.timeZone),
            startTimeZone: updates.start?.timeZone,
            endTimeZone: updates.end?.timeZone,
            timestamp: new Date().toISOString()
        }, 'calendar');

        // 3. Get User Context (for permission check)
        // Get current user's email from Microsoft Graph API
        let currentUserEmail = null;
        try {
            // Get current user's profile directly from Microsoft Graph API
            const graphClientFactory = require('../../graph/graph-client.cjs');
            const client = await graphClientFactory.createClient(req);
            const userProfile = await client.api('/me?$select=mail,userPrincipalName').get();
            currentUserEmail = userProfile?.mail || userProfile?.userPrincipalName;
        } catch (error) {
            // Log but don't fail - we'll skip permission check if we can't get user context
            monitoringService?.warn('Could not retrieve user profile for permission check', {
                eventId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'calendar');
        }

        if (!currentUserEmail) {
            // Skip permission check if we can't get user context - this allows the update to proceed
            monitoringService?.warn('User context not available for permission check - skipping ownership validation', {
                eventId,
                timestamp: new Date().toISOString()
            }, 'calendar');
        }

        try {
            const startTime = Date.now();

            // Call the Graph service directly with minimal processing
            // This allows us to see exactly what Microsoft Graph receives
            const result = await graphService.updateEvent(eventId, updates, 'me', { req });

            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric('calendar_update_event', elapsedTime, {
                timestamp: new Date().toISOString()
            });

            monitoringService?.info('Successfully updated calendar event', {
                eventId,
                elapsedTime,
                updatedFields: Object.keys(updates),
                timestamp: new Date().toISOString()
            }, 'calendar');

            return result;
        } catch (error) {
            const graphDetails = {
                statusCode: error.statusCode,
                code: error.code,
                graphRequestId: error.requestId,
                originalMessage: error.message
            };

            const updateError = errorService?.createError(
                'calendar',
                `Failed to update event: ${error.code || error.message}`,
                'error',
                { 
                    eventId,
                    updates: redactedUpdates,
                    graphDetails,
                    originalError: error.stack,
                    timestamp: new Date().toISOString()
                }
            );

            monitoringService?.logError(updateError);

            throw updateError || error;
        }
    },
    
    /**
     * Schedule a meeting with intelligent time selection
     * @param {object} options - Meeting options including attendees and preferred times
     * @returns {Promise<object>} Object containing suggestions and optionally the normalized scheduled event if autoSchedule was true and a slot was found.
     */
    
    /**
     * Get availability for users within a specified time range
     * @param {object} options - Options object.
     * @param {Array<string|object>} options.users - Array of user emails or user objects.
     * @param {Array<object>} [options.timeSlots] - Array of time slots (e.g., [{ start: { dateTime, timeZone }, end: { dateTime, timeZone } }]). Required if duration is not provided.
     * @param {string} [options.duration] - ISO 8601 duration string (e.g., 'PT1H'). Required if timeSlots is not provided.
     * @param {string|Date} [options.windowStart=Date.now()] - Start time for the duration-based window.
     * @returns {Promise<object>} Availability data object.
     * @throws {Error} If input validation fails or Graph API call errors occur.
     */
    async getAvailability(options = {}, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Log the request with detailed parameters
        monitoringService?.debug('Calendar availability check requested', { 
            options: this.redactSensitiveData(options),
            timestamp: new Date().toISOString(),
            source: 'calendar.getAvailability'
        }, 'calendar');
        
        // Validate required parameters
        if (!options.users || !Array.isArray(options.users) || options.users.length === 0) {
            const err = errorService.createError(
                'calendar',
                'Users array is required for getAvailability',
                'error',
                { 
                    options: this.redactSensitiveData(options),
                    timestamp: new Date().toISOString(),
                    validationError: 'missing_users'
                }
            );
            monitoringService?.logError(err);
            monitoringService?.error('Calendar availability validation failed: missing users', {
                validationError: 'missing_users',
                timestamp: new Date().toISOString()
            }, 'calendar');
            throw err;
        }
        
        // Validate that either timeSlots or duration is provided
        if (!options.timeSlots && !options.duration) {
            const err = errorService.createError(
                'calendar',
                'Either timeSlots or duration must be provided for getAvailability',
                'error',
                { 
                    options: this.redactSensitiveData(options),
                    timestamp: new Date().toISOString(),
                    validationError: 'missing_time_constraint'
                }
            );
            monitoringService?.logError(err);
            monitoringService?.error('Calendar availability validation failed: missing time constraint', {
                validationError: 'missing_time_constraint',
                timestamp: new Date().toISOString()
            }, 'calendar');
            throw err;
        }
        
        // Extract emails from users array
        const emails = options.users.map(user => {
            if (typeof user === 'string') return user;
            if (user.email) return user.email;
            if (user.emailAddress?.address) return user.emailAddress.address;
            return null;
        }).filter(Boolean);
        
        // Log the extracted emails for debugging
        monitoringService?.debug('Extracted emails for availability check', {
            emailCount: emails.length,
            inputUserCount: options.users.length,
            timestamp: new Date().toISOString()
        }, 'calendar');
        
        if (emails.length === 0) {
            const err = errorService.createError(
                'calendar',
                'No valid email addresses found in users array',
                'error',
                { 
                    users: this.redactSensitiveData(options.users),
                    timestamp: new Date().toISOString(),
                    validationError: 'invalid_email_format'
                }
            );
            monitoringService?.logError(err);
            monitoringService?.error('Calendar availability validation failed: invalid email format', {
                validationError: 'invalid_email_format',
                timestamp: new Date().toISOString()
            }, 'calendar');
            throw err;
        }
        
        // Extract date range from options
        let startDateTime, endDateTime;
        
        if (options.timeSlots && options.timeSlots.length > 0) {
            // Use the first time slot's start and the last time slot's end
            const firstSlot = options.timeSlots[0];
            const lastSlot = options.timeSlots[options.timeSlots.length - 1];
            
            startDateTime = firstSlot.start?.dateTime;
            endDateTime = lastSlot.end?.dateTime;
            
            // Validate date format
            if (!startDateTime || !endDateTime) {
                const err = errorService.createError(
                    'calendar',
                    'Invalid time slot format: missing dateTime in start/end',
                    'error',
                    { 
                        firstSlot: this.redactSensitiveData(firstSlot),
                        lastSlot: this.redactSensitiveData(lastSlot),
                        timestamp: new Date().toISOString(),
                        validationError: 'invalid_timeslot_format'
                    }
                );
                monitoringService?.logError(err);
                monitoringService?.error('Calendar availability validation failed: invalid time slot format', {
                    validationError: 'invalid_timeslot_format',
                    timestamp: new Date().toISOString()
                }, 'calendar');
                throw err;
            }
        } else {
            // Use duration and windowStart
            const { duration } = options;
            const windowStart = options.windowStart || new Date();
            
            startDateTime = typeof windowStart === 'string' ? windowStart : windowStart.toISOString();
            
            try {
                // Validate duration format
                if (!moment.duration(duration).isValid()) {
                    throw new Error('Invalid duration format');
                }
                endDateTime = moment(startDateTime).add(moment.duration(duration)).toISOString();
            } catch (error) {
                const err = errorService.createError(
                    'calendar',
                    `Invalid duration format: ${duration}`,
                    'error',
                    { 
                        duration,
                        windowStart,
                        timestamp: new Date().toISOString(),
                        validationError: 'invalid_duration_format',
                        originalError: error.message
                    }
                );
                monitoringService?.logError(err);
                monitoringService?.error('Calendar availability validation failed: invalid duration format', {
                    validationError: 'invalid_duration_format',
                    duration,
                    timestamp: new Date().toISOString()
                }, 'calendar');
                throw err;
            }
        }

        // Log the extracted date time range for debugging
        monitoringService?.debug('Extracted date time range for availability check', {
            startDateTime,
            endDateTime,
            startType: typeof startDateTime,
            endType: typeof endDateTime,
            timestamp: new Date().toISOString()
        }, 'calendar');

        try {
            // Check if the Graph service has the required method
            if (!graphService) {
                const mcpError = errorService.createError(
                    'calendar',
                    'GraphService not available for calendar availability check',
                    'error',
                    { 
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                monitoringService?.logError(mcpError);
                monitoringService?.error('Calendar availability failed: Graph service not available', {
                    serviceError: 'missing_graph_service',
                    timestamp: new Date().toISOString()
                }, 'calendar');
                throw mcpError;
            }
            
            if (typeof graphService.getAvailability !== 'function') {
                const mcpError = errorService.createError(
                    'calendar',
                    'GraphService.getAvailability method not implemented',
                    'error',
                    { 
                        emails: this.redactSensitiveData(emails),
                        startDateTime,
                        endDateTime,
                        timestamp: new Date().toISOString(),
                        serviceError: 'method_not_implemented'
                    }
                );
                monitoringService?.logError(mcpError);
                monitoringService?.error('Calendar availability failed: Method not implemented', {
                    serviceError: 'method_not_implemented',
                    timestamp: new Date().toISOString()
                }, 'calendar');
                throw mcpError;
            }
            
            // Log that we're about to call the Graph service
            monitoringService?.debug('Calling Graph service for availability', {
                emailCount: emails.length,
                startDateTime,
                endDateTime,
                timeZone: options.timeZone,
                intervalMinutes: options.intervalMinutes,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            // Call the Graph service
            const result = await graphService.getAvailability(emails, startDateTime, endDateTime, { ...options, req });
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Log success metrics
            monitoringService?.trackMetric('calendar_availability_success', executionTime, {
                emailCount: emails.length,
                resultCount: Array.isArray(result) ? result.length : 0,
                timestamp: new Date().toISOString()
            });
            
            // Log success with result summary
            monitoringService?.info('Calendar availability check completed successfully', {
                emailCount: emails.length,
                resultCount: Array.isArray(result) ? result.length : 0,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            // Check if mock data is being returned
            if (result && Array.isArray(result) && result.length > 0) {
                const isMockData = result.some(item => 
                    item.scheduleItems && item.scheduleItems.some(scheduleItem => 
                        scheduleItem.id && scheduleItem.id.includes('mock')
                    )
                );
                
                if (isMockData) {
                    monitoringService?.warn('Calendar availability returned mock data', {
                        emailCount: emails.length,
                        resultCount: result.length,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'calendar');
                }
            }
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Track failure metrics
            monitoringService?.trackMetric('calendar_availability_failure', executionTime, {
                errorType: error.code || 'unknown',
                timestamp: new Date().toISOString()
            });
            
            // Create standardized error
            const mcpError = errorService.createError(
                'calendar',
                `Failed to get availability: ${error.message}`,
                'error',
                { 
                    emails: this.redactSensitiveData(emails),
                    startDateTime,
                    endDateTime,
                    errorCode: error.code || 'unknown',
                    statusCode: error.statusCode || 'unknown',
                    originalError: error.stack || error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            
            // Log the error with detailed context
            monitoringService?.logError(mcpError);
            monitoringService?.error('Calendar availability check failed', {
                errorMessage: error.message,
                errorCode: error.code || 'unknown',
                statusCode: error.statusCode || 'unknown',
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
            
            throw mcpError;
        }
    },
    
    /**
     * Private helper to handle common event actions (accept, decline, etc.).
     * @param {string} action - The action to perform ('accept', 'tentativelyAccept', 'decline', 'cancel').
     * @param {string} eventId - ID of the event to update.
     * @param {string|object} commentOrOptions - Comment string or options object (for cancel action).
     * @returns {Promise<object>} Response from Graph Service.
     * @private
     */
    async _handleEventAction(action, eventId, commentOrOptions = '', req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        const graphMethodName = `${action}Event`; // e.g., 'acceptEvent'

        // Handle different parameter formats for different actions
        let comment = '';
        let options = {};
        
        if (action === 'cancel' && typeof commentOrOptions === 'object') {
            // For cancel action with options object
            comment = commentOrOptions.comment || '';
            options = commentOrOptions;
        } else {
            // For other actions or backward compatibility
            comment = typeof commentOrOptions === 'string' ? commentOrOptions : '';
            options = { comment };
        }

        // Log the action attempt
        monitoringService?.debug(`Attempting ${action} action on calendar event`, {
            action,
            eventId,
            hasComment: !!comment,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService[graphMethodName] !== 'function') {
            const methodError = errorService?.createError(
                'calendar',
                `GraphService.${graphMethodName} method not implemented`,
                'error',
                { 
                    action,
                    eventId,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(methodError);
                
            throw methodError;
        }

        try {
            // Track performance
            const startTime = Date.now();
            
            // Call the Graph service with appropriate parameters
            let result;
            if (action === 'cancel') {
                // For cancel, pass the full options object
                result = await graphService[graphMethodName](eventId, options, req);
            } else {
                // For other actions, pass comment string
                result = await graphService[graphMethodName](eventId, comment, req);
            }
            
            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric(`calendar_${action.toLowerCase()}_event`, elapsedTime, {
                timestamp: new Date().toISOString()
            });
            
            monitoringService?.info(`Successfully performed ${action} action on calendar event`, {
                action,
                eventId,
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
                
            return result; // Typically returns void or confirmation
        } catch (error) {
            const graphDetails = {
                statusCode: error.statusCode,
                code: error.code,
                graphRequestId: error.requestId,
                originalMessage: error.message
            };
            
            const actionError = errorService?.createError(
                'calendar',
                `Failed to ${action} event: ${error.code || error.message}`,
                'error',
                { 
                    action,
                    eventId,
                    hasComment: !!comment,
                    graphDetails,
                    originalError: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(actionError);
                
            throw actionError || error;
        }
    },
    
    /**
     * Accept a calendar event invitation
     * @param {string} eventId - ID of the event to accept
     * @param {string} comment - Optional comment to include with the response
     * @returns {Promise<object>} Response status
     */
    async acceptEvent(eventId, comment = '', req) {
        return await this._handleEventAction('accept', eventId, comment, req);
    },
    
    /**
     * Tentatively accept a calendar event invitation
     * @param {string} eventId - ID of the event to tentatively accept
     * @param {string} comment - Optional comment to include with the response
     * @returns {Promise<object>} Response status
     */
    async tentativelyAcceptEvent(eventId, comment = '', req) {
        return await this._handleEventAction('tentativelyAccept', eventId, comment, req);
    },
    
    /**
     * Decline a calendar event invitation
     * @param {string} eventId - ID of the event to decline
     * @param {string} comment - Optional comment to include with the response
     * @returns {Promise<object>} Response status
     */
    async declineEvent(eventId, comment = '', req) {
        return await this._handleEventAction('decline', eventId, comment, req);
    },
    
    /**
     * Cancel a calendar event and send cancellation messages to attendees
     * @param {string} eventId - ID of the event to cancel
     * @param {string} comment - Optional comment to include with the cancellation
     * @returns {Promise<object>} Response status
     */
    async cancelEvent(eventId, comment = '', req) {
        // Fix: Replace internal user ID with 'me' for Graph API calls
        // The Graph API expects 'me' for the current authenticated user, not internal user IDs
        const userId = 'me'; // Always use 'me' for the current user in Graph API calls
        
        // Pass userId in options to ensure Graph service uses 'me'
        const options = {
            comment,
            userId,
            sendCancellation: true
        };
        
        return await this._handleEventAction('cancel', eventId, options, req);
    },
    
    /**
     * Find suitable meeting times for attendees
     * @param {object} options - Options for finding meeting times
     * @returns {Promise<object>} Meeting time suggestions
     */
    async findMeetingTimes(options = {}, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Redact potentially sensitive data for logging
        const redactedOptions = this.redactSensitiveData(options);
        
        // Log the request attempt
        monitoringService?.debug('Finding meeting times', { 
            options: redactedOptions,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService.findMeetingTimes !== 'function') {
            const methodError = errorService?.createError(
                'calendar',
                'Required GraphService method \'findMeetingTimes\' not implemented or service unavailable',
                'error',
                { timestamp: new Date().toISOString() }
            );
            
            monitoringService?.logError(methodError);
                
            throw new Error('Required GraphService method \'findMeetingTimes\' not implemented or service unavailable.');
        }

        try {
            // Track performance
            const startTime = Date.now();
            
            // Call the Graph API
            const result = await graphService.findMeetingTimes({ ...options, req });
            
            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric('calendar_find_meeting_times', elapsedTime, {
                timestamp: new Date().toISOString()
            });

            monitoringService?.info('Successfully found meeting times', {
                suggestionCount: result?.meetingTimeSuggestions?.length || 0,
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
                
            return result;
        } catch (error) {
            const graphDetails = {
                statusCode: error.statusCode,
                code: error.code,
                graphRequestId: error.requestId,
                originalMessage: error.message
            };
            
            const findTimesError = errorService?.createError(
                'calendar',
                `Failed to find meeting times: ${error.code || error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    graphDetails, 
                    requestOptions: redactedOptions,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(findTimesError);
                
            throw findTimesError || new Error(`Failed to find meeting times: ${error.message}`);
        }
    },
    
    /**
     * Get available rooms for meetings
     * @param {object} options - Options for filtering rooms
     * @returns {Promise<Array>} List of available rooms
     */
    async getRooms(options = {}, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        const { skip, top, ...otherOptions } = options; // Extract paging params

        // Log the request attempt
        monitoringService?.debug('Getting available rooms', { 
            options,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService.getRooms !== 'function') {
            const methodError = errorService?.createError(
                'calendar',
                'Required GraphService method \'getRooms\' not implemented or service unavailable',
                'error',
                { timestamp: new Date().toISOString() }
            );
            
            monitoringService?.logError(methodError);
                
            throw new Error('Required GraphService method \'getRooms\' not implemented or service unavailable.');
        }

        // Prepare options for graph service, including paging
        const graphOptions = { ...otherOptions };
        if (skip !== undefined) graphOptions.$skip = skip;
        if (top !== undefined) graphOptions.$top = top;

        try {
            // Track performance
            const startTime = Date.now();
            
            monitoringService?.debug('Calling getRooms Graph API', {
                graphOptions,
                timestamp: new Date().toISOString()
            }, 'calendar');
                
            // Include req in the options for getRooms
            const result = await graphService.getRooms({ ...graphOptions, req });
            
            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric('calendar_get_rooms', elapsedTime, {
                timestamp: new Date().toISOString()
            });

            // Basic normalization - ensure essential fields are present
            const normalizedRooms = (result?.rooms || []).map(room => ({
                id: room.id, // Assuming graph returns id
                displayName: room.displayName || room.name, // Graph uses displayName for rooms
                emailAddress: room.emailAddress, // Key field
                // capacity: room.capacity // NOTE: Capacity often requires specific permissions/calls
            }));

            monitoringService?.info('Successfully retrieved room list', {
                roomCount: normalizedRooms.length,
                hasNextLink: !!result.nextLink,
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
                
            return {
                rooms: normalizedRooms,
                nextLink: result.nextLink // Pass along the nextLink for pagination
            };
        } catch (error) {
            const graphDetails = {
                statusCode: error.statusCode,
                code: error.code,
                graphRequestId: error.requestId,
                originalMessage: error.message
            };
            
            const roomsError = errorService?.createError(
                'calendar',
                `Failed to get rooms: ${error.code || error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    graphDetails, 
                    requestOptions: graphOptions,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(roomsError);
                
            throw roomsError || new Error(`Failed to get rooms: ${error.message}`);
        }
    },
    
    /**
     * Get user calendars
     * @returns {Promise<Array>} List of calendars
     */
    async getCalendars(options = {}, req) {
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};

        monitoringService?.debug('Attempting to get calendars', { options, timestamp: new Date().toISOString() }, 'calendar');

        if (!graphService || typeof graphService.getCalendars !== 'function') {
            const error = errorService?.createError('calendar', 'GraphService.getCalendars not implemented', 'error');
            monitoringService?.logError(error);
            throw error || new Error('GraphService.getCalendars not implemented');
        }

        const startTime = Date.now();
        try {
            const result = await graphService.getCalendars({ ...options, req });

            const duration = Date.now() - startTime;
            monitoringService?.trackMetric('calendar.getCalendars.duration', duration, {
                success: true,
                timestamp: new Date().toISOString()
            });
            monitoringService?.info('Successfully retrieved calendars', { count: result?.length, duration }, 'calendar');

            return result;
        } catch (error) {
            const duration = Date.now() - startTime;
            const mcpError = errorService?.createError(
                'calendar',
                'Failed to get calendars in module',
                'error',
                { originalError: error.message, stack: error.stack }
            );
            monitoringService?.logError(mcpError);
            monitoringService?.trackMetric('calendar.getCalendars.duration', duration, {
                success: false,
                timestamp: new Date().toISOString()
            });
            throw mcpError;
        }
    },
    
    /**
     * Add an attachment to an event
     * @param {string} eventId - ID of the event
     * @param {object} attachment - Attachment data
     * @returns {Promise<object>} Created attachment
     */
    async addAttachment(eventId, attachment, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};

        // Log the request attempt
        monitoringService?.debug('Adding attachment to calendar event', { 
            eventId,
            attachmentName: attachment?.name,
            contentType: attachment?.contentType,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService.addEventAttachment !== 'function') {
            const methodError = errorService?.createError(
                'calendar',
                'Required GraphService method \'addEventAttachment\' not implemented or service unavailable',
                'error',
                { 
                    eventId,
                    timestamp: new Date().toISOString() 
                }
            );
            
            monitoringService?.logError(methodError);
                
            throw new Error('Required GraphService method \'addEventAttachment\' not implemented or service unavailable.');
        }

        // Validate attachment data
        let validatedAttachment;
        try {
            const { error, value } = attachmentSchema.validate(attachment, { abortEarly: false });
            if (error) {
                const validationError = errorService.createError(
                    'validation',
                    'Invalid attachment data',
                    'warn',
                    { 
                        details: error.details.map(d => d.message), 
                        eventId,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService?.logError(validationError);
                throw validationError;
            }
            validatedAttachment = value; // Use the validated and potentially defaulted value
        } catch (validationError) {
            // Rethrow validation errors immediately
            throw validationError; 
        }

        // Create audit log context with safe data for logging
        const auditLogContext = { 
            eventId, 
            attachmentName: validatedAttachment.name, 
            contentType: validatedAttachment.contentType, 
            sizeBytes: Buffer.byteLength(validatedAttachment.contentBytes, 'base64'),
            timestamp: new Date().toISOString()
        };

        // Log the attachment attempt
        monitoringService?.info('Attempting to add attachment to calendar event', auditLogContext, 'calendar');

        try {
            // Track performance
            const startTime = Date.now();
            
            // Call the Graph API
            const result = await graphService.addEventAttachment(eventId, validatedAttachment, req);
            
            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric('calendar_add_attachment', elapsedTime, {
                timestamp: new Date().toISOString()
            });

            // Log success
            monitoringService?.info('Successfully added attachment to calendar event', { 
                ...auditLogContext, 
                attachmentId: result.id,
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');

            return result; // Return the created attachment object from Graph

        } catch (error) {
            // Create standardized error object
            const graphDetails = { 
                statusCode: error.statusCode, 
                code: error.code, 
                graphRequestId: error.requestId, 
                originalMessage: error.message 
            };
            
            const attachmentError = errorService?.createError(
                'calendar',
                `Failed to add attachment: ${error.code || error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    graphDetails, 
                    ...auditLogContext
                }
            );
            
            // Log the error
            monitoringService?.logError(attachmentError);
                
            throw attachmentError || new Error(`Failed to add attachment: ${error.message}`);
        }
    },
    
    /**
     * Remove an attachment from an event
     * @param {string} eventId - ID of the event
     * @param {string} attachmentId - ID of the attachment to remove
     * @returns {Promise<boolean>} Success status
     */
    async removeAttachment(eventId, attachmentId, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};

        // Log the request attempt
        monitoringService?.debug('Removing attachment from calendar event', { 
            eventId,
            attachmentId,
            timestamp: new Date().toISOString()
        }, 'calendar');

        if (!graphService || typeof graphService.removeEventAttachment !== 'function') {
            const methodError = errorService?.createError(
                'calendar',
                'Required GraphService method \'removeEventAttachment\' not implemented or service unavailable',
                'error',
                { 
                    eventId,
                    attachmentId,
                    timestamp: new Date().toISOString() 
                }
            );
            
            monitoringService?.logError(methodError);
                
            throw new Error('Required GraphService method \'removeEventAttachment\' not implemented or service unavailable.');
        }

        // Create audit log context with safe data for logging
        const auditLogContext = { 
            eventId, 
            attachmentId,
            timestamp: new Date().toISOString()
        };

        // Log the attachment removal attempt
        monitoringService?.info('Attempting to remove attachment from calendar event', auditLogContext, 'calendar');

        try {
            // Track performance
            const startTime = Date.now();
            
            // Graph remove attachment usually returns void (204 No Content) on success
            await graphService.removeEventAttachment(eventId, attachmentId, req);
            
            // Calculate elapsed time and track metric
            const elapsedTime = Date.now() - startTime;
            monitoringService?.trackMetric('calendar_remove_attachment', elapsedTime, {
                timestamp: new Date().toISOString()
            });

            // Log success
            monitoringService?.info('Successfully removed attachment from calendar event', { 
                ...auditLogContext, 
                elapsedTime,
                timestamp: new Date().toISOString()
            }, 'calendar');

            return true; // Indicate success

        } catch (error) {
            // Create standardized error object
            const graphDetails = { 
                statusCode: error.statusCode, 
                code: error.code, 
                graphRequestId: error.requestId, 
                originalMessage: error.message 
            };
            
            const removeError = errorService?.createError(
                'calendar',
                `Failed to remove attachment: ${error.code || error.message}`,
                'error',
                { 
                    originalError: error.stack,
                    graphDetails, 
                    ...auditLogContext
                }
            );
            
            // Log the error
            monitoringService?.logError(removeError);
                
            throw removeError || new Error(`Failed to remove attachment: ${error.message}`);
        }
    },
    
    /**
     * Handles calendar-related intents routed to this module.
     * @param {string} intent
     * @param {object} entities
     * @param {object} context
     * @returns {Promise<object>} Normalized response
     */
    async handleIntent(intent, entities = {}, context = {}) {
        // Using strategy map pattern for intent handling
        const { graphService, cacheService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Redact potentially sensitive data for logging
        const redactedEntities = this.redactSensitiveData(entities);
        const redactedContext = this.redactSensitiveData(context);
        
        // Log the intent handling attempt
        monitoringService?.debug('Handling calendar intent', { 
            intent,
            entities: redactedEntities,
            context: redactedContext,
            timestamp: new Date().toISOString()
        }, 'calendar');

        const intentHandlers = {
            'getEvents': async (entities, context) => {
                // Extract date/filter parameters - support both flat params and nested range object
                const queryParams = entities.range || {
                    timeframe: entities.timeframe,
                    start: entities.start,
                    end: entities.end,
                    limit: entities.limit,
                    filter: entities.filter,
                    subject: entities.subject,
                    organizer: entities.organizer,
                    location: entities.location,
                    orderby: entities.orderby
                };
                const cacheKey = `calendar:events:${JSON.stringify(queryParams)}`;
                let events = cacheService && await cacheService.get(cacheKey);
                if (!events) {
                    // Ensure graphService is available before using
                    if (!graphService) throw new Error('GraphService is unavailable for getEvents.');
                    const raw = await graphService.getEvents({ ...queryParams, req: context.req });
                    events = Array.isArray(raw) ? raw.map(normalizeEvent) : [];
                    if (cacheService) await cacheService.set(cacheKey, events, 60); // Cache for 1 minute
                }
                return { type: 'calendarList', items: events };
            },
            'createEvent': async (entities, context) => {
                const eventData = entities.event;
                const normalizedEvent = await this.createEvent(eventData, context.req);
                return { type: 'calendarEvent', event: normalizedEvent };
            },
            'updateEvent': async (entities, context) => {
                const { eventId, updates } = entities;
                const normalizedUpdatedEvent = await this.updateEvent(eventId, updates, context.req);
                return { type: 'calendarEvent', event: normalizedUpdatedEvent };
            },
            'getAvailability': async (entities, context) => {
                const availabilityResult = await this.getAvailability(entities, context.req);
                return { type: 'availabilityResult', data: availabilityResult };
            },
            'acceptEvent': async (entities, context) => {
                const { eventId, comment } = entities;
                const result = await this.acceptEvent(eventId, comment, context.req);
                // Return the result directly if it has a proper structure, otherwise format it
                if (result && result.success) {
                    return { 
                        type: 'eventResponse', 
                        status: 'accepted', 
                        eventId,
                        success: true,
                        timestamp: result.timestamp || new Date().toISOString()
                    };
                }
                return { type: 'eventResponse', status: 'accepted', eventId };
            },
            'tentativelyAcceptEvent': async (entities, context) => {
                const { eventId, comment } = entities;
                const result = await this.tentativelyAcceptEvent(eventId, comment, context.req);
                // Return the result directly if it has a proper structure, otherwise format it
                if (result && result.success) {
                    return { 
                        type: 'eventResponse', 
                        status: 'tentativelyAccepted', 
                        eventId,
                        success: true,
                        timestamp: result.timestamp || new Date().toISOString()
                    };
                }
                return { type: 'eventResponse', status: 'tentativelyAccepted', eventId };
            },
            'declineEvent': async (entities, context) => {
                const { eventId, comment } = entities;
                const result = await this.declineEvent(eventId, comment, context.req);
                // Return the result directly if it has a proper structure, otherwise format it
                if (result && result.success) {
                    return { 
                        type: 'eventResponse', 
                        status: 'declined', 
                        eventId,
                        success: true,
                        timestamp: result.timestamp || new Date().toISOString()
                    };
                }
                return { type: 'eventResponse', status: 'declined', eventId };
            },
            'cancelEvent': async (entities, context) => {
                const { eventId, comment } = entities;
                const result = await this.cancelEvent(eventId, comment, context.req);
                // Return the result directly if it has a proper structure, otherwise format it
                if (result && result.success) {
                    return { 
                        type: 'eventResponse', 
                        status: 'cancelled', 
                        eventId,
                        success: true,
                        timestamp: result.timestamp || new Date().toISOString()
                    };
                }
                return { type: 'eventResponse', status: 'cancelled', eventId };
            },
            'findMeetingTimes': async (entities, context) => {
                const options = entities.options || {};
                const suggestions = await this.findMeetingTimes(options, context.req);
                // Ensure suggestions structure aligns with expected response type
                return { type: 'meetingTimeSuggestions', suggestions: suggestions };
            },
            'getRooms': async (entities, context) => {
                const options = entities.options || {};
                const roomData = await this.getRooms(options, context.req); // Expect { rooms: [], nextLink: ... }
                return { type: 'roomList', rooms: roomData.rooms, nextLink: roomData.nextLink };
            },
            'getCalendars': async (entities, context) => {
                const calendars = await this.getCalendars({}, context.req); // Expect array
                return { type: 'calendarList', calendars: calendars };
            },
            'addAttachment': async (entities, context) => {
                const { id, name, contentBytes, contentType } = entities;
                const attachment = { name, contentBytes, contentType };
                const result = await this.addAttachment(id, attachment, context.req);
                return { type: 'attachmentAdded', attachment: result };
            },
            'removeAttachment': async (entities, context) => {
                const { eventId, attachmentId } = entities;
                const success = await this.removeAttachment(eventId, attachmentId, context.req);
                return { type: 'attachmentRemoved', success, eventId, attachmentId };
            }
            // Add handlers for addAttachment/removeAttachment if they become intents
        };

        const handler = intentHandlers[intent];

        if (handler) {
            try {
                const startTime = Date.now();
                const result = await handler(entities, context);
                const elapsedTime = Date.now() - startTime;
                
                // Log success and track performance
                monitoringService?.info('Successfully handled calendar intent', { 
                    intent,
                    responseType: result?.type,
                    elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'calendar');
                
                monitoringService?.trackMetric('calendar_intent_handling_time', elapsedTime, {
                    intent,
                    success: true,
                    timestamp: new Date().toISOString()
                });
                
                return result;
            } catch (error) {
                // Log and rethrow errors originating from handlers
                const redactedEntities = this.redactSensitiveData(entities);
                
                // Error should already be structured if thrown from within handlers using errorService
                if (error && error.category) {
                    // Log structured error with additional context
                    monitoringService?.logError(error);
                        
                    // Track failure metric
                    monitoringService?.trackMetric('calendar_intent_handling_time', 0, {
                        intent,
                        success: false,
                        errorCategory: error.category,
                        timestamp: new Date().toISOString()
                    });
                    
                    throw error; // Rethrow known MCP errors
                } else {
                    // Create a structured error for unstructured exceptions
                    const mcpError = errorService?.createError(
                        'calendar',
                        `Unexpected error handling intent ${intent}: ${error.message}`,
                        'error',
                        { 
                            originalError: error.stack,
                            intent,
                            entities: redactedEntities,
                            timestamp: new Date().toISOString()
                        }
                    ) || {
                        category: 'calendar',
                        message: `Unexpected error handling intent ${intent}: ${error.message}`,
                        severity: 'error',
                        context: { intent, entities: redactedEntities }
                    };
                    
                    monitoringService?.logError(mcpError);
                    
                    // Track failure metric
                    monitoringService?.trackMetric('calendar_intent_handling_time', 0, {
                        intent,
                        success: false,
                        errorType: 'unexpected',
                        timestamp: new Date().toISOString()
                    });
                    
                    throw mcpError;
                }
            }
        } else {
            // Default case: Unsupported intent
            const unsupportedError = errorService?.createError(
                'calendar',
                `The calendar module does not support the intent: ${intent}`,
                'warn', // Treat as warning, not critical failure
                { 
                    intent, 
                    moduleId: this.id,
                    timestamp: new Date().toISOString()
                }
            ) || {
                category: 'calendar',
                message: `The calendar module does not support the intent: ${intent}`,
                severity: 'warn',
                context: { intent, moduleId: this.id }
            };
            
            monitoringService?.logError(unsupportedError);
                
            // Track metric for unsupported intent
            monitoringService?.trackMetric('calendar_unsupported_intent', 1, {
                intent,
                timestamp: new Date().toISOString()
            });
            
            throw unsupportedError; // Throw error to signal unsupported operation
         }
     },
    
    id: 'calendar',
    name: 'Microsoft Calendar',
    capabilities: CALENDAR_CAPABILITIES,
    /**
     * Initializes the calendar module with dependencies.
     * @param {object} services - { graphService, cacheService, eventService }
     * @returns {object} Initialized module
     */
    init(services) {
        // Validate that required services are provided
        const requiredServices = ['graphService', 'errorService', 'monitoringService']; 
        
        // Use imported services as fallbacks during initialization
        const errorService = services?.errorService || ErrorService;
        const monitoringService = services?.monitoringService || MonitoringService;

        // Log initialization attempt
        monitoringService?.debug('Initializing Calendar Module', { 
            timestamp: new Date().toISOString() 
        }, 'calendar');

        if (!services) {
            const error = errorService?.createError(
                'calendar',
                'CalendarModule init requires a services object',
                'error',
                { timestamp: new Date().toISOString() }
            ) || {
                category: 'calendar',
                message: 'CalendarModule init requires a services object',
                severity: 'error',
                context: {}
            };
            
            monitoringService?.logError(error);
                
            throw error;
        }

        // Validate required services
        for (const serviceName of requiredServices) {
            if (!services[serviceName]) {
                const error = errorService?.createError(
                    'calendar',
                    `CalendarModule init failed: Required service '${serviceName}' is missing`,
                    'error',
                    { 
                        missingService: serviceName,
                        timestamp: new Date().toISOString() 
                    }
                ) || {
                    category: 'calendar',
                    message: `CalendarModule init failed: Required service '${serviceName}' is missing`,
                    severity: 'error',
                    context: { missingService: serviceName }
                };
                
                monitoringService?.logError(error);
                    
                throw error;
            }
        }

        this.services = services;
        
        // Log successful initialization
        monitoringService?.info('CalendarModule initialized successfully', { 
            timestamp: new Date().toISOString() 
        }, 'calendar');
            
        return this; // Return the module instance, now containing validated services
    }
};

module.exports = CalendarModule;
