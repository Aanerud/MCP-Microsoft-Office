/**
 * @fileoverview Normalization functions for Microsoft Graph API responses.
 * Follows MCP modular, testable, and consistent data contract rules.
 */

const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

/**
 * Normalizes a Microsoft Graph email object to the MCP standard format.
 * @param {object} graphEmail - Raw email object from Graph API
 * @param {string} [userId] - User ID for context tracking
 * @param {string} [sessionId] - Session ID for context tracking
 * @returns {object} Normalized email object
 */
function normalizeEmail(graphEmail, userId, sessionId) {
    const MAX_SUBJECT_LENGTH = 150;
    const startTime = Date.now();

    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting email normalization', {
            emailId: graphEmail?.id?.substring(0, 20) + '...' || 'unknown',
            hasAttachments: !!graphEmail?.hasAttachments,
            userId: userId || 'anonymous',
            sessionId: sessionId || 'no-session',
            timestamp: new Date().toISOString()
        }, 'graph');
    }

    if (!graphEmail || typeof graphEmail !== 'object') {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'graph',
            'Invalid email object for normalization',
            'error',
            {
                emailType: typeof graphEmail,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Email normalization failed - invalid input', {
                error: 'Invalid email object provided',
                emailType: typeof graphEmail,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Email normalization failed - invalid input', {
                sessionId: sessionId,
                error: 'Invalid email object provided',
                emailType: typeof graphEmail,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw error;
    }
    try {
        const normalizedEmail = {
            id: graphEmail.id,
            type: 'email',
            subject: graphEmail.subject ? (graphEmail.subject.length > MAX_SUBJECT_LENGTH ? graphEmail.subject.substring(0, MAX_SUBJECT_LENGTH) + '...' : graphEmail.subject) : '',
            from: graphEmail.from && graphEmail.from.emailAddress ? {
                name: graphEmail.from.emailAddress.name,
                email: graphEmail.from.emailAddress.address
            } : undefined,
            to: Array.isArray(graphEmail.toRecipients) ? graphEmail.toRecipients.map(r => ({
                name: r.emailAddress.name,
                email: r.emailAddress.address
            })) : [],
            received: graphEmail.receivedDateTime || null,
            sent: graphEmail.sentDateTime || null,
            preview: graphEmail.bodyPreview ? graphEmail.bodyPreview.substring(0, 150) : '',
            isRead: !!graphEmail.isRead,
            importance: graphEmail.importance,
            hasAttachments: !!graphEmail.hasAttachments,
            hasInlineImages: !!(graphEmail.attachments && graphEmail.attachments.some(att => att.isInline))
        };
        
        // Track performance metric
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('email_normalization_time', executionTime, {
            emailId: graphEmail.id || 'unknown',
            userId: userId || 'anonymous',
            timestamp: new Date().toISOString()
        }, userId, null, false, sessionId);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Email normalized successfully', {
                emailId: graphEmail.id?.substring(0, 20) + '...' || 'unknown',
                hasAttachments: !!graphEmail.hasAttachments,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Email normalized successfully', {
                sessionId: sessionId,
                emailId: graphEmail.id?.substring(0, 20) + '...' || 'unknown',
                hasAttachments: !!graphEmail.hasAttachments,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return normalizedEmail;
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Failed to normalize email: ${error.message || 'Unknown error'}`,
            'error',
            {
                emailId: graphEmail.id || 'unknown',
                originalError: error.message,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Email normalization failed', {
                error: error.message || 'Unknown error',
                emailId: graphEmail.id?.substring(0, 20) + '...' || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Email normalization failed', {
                sessionId: sessionId,
                error: error.message || 'Unknown error',
                emailId: graphEmail.id?.substring(0, 20) + '...' || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
}

/**
 * Normalizes a Microsoft Graph driveItem (file/folder) to MCP format.
 * @param {object} item - Raw driveItem from Graph API
 * @param {string} [userId] - User ID for context tracking
 * @param {string} [sessionId] - Session ID for context tracking
 * @returns {object} Normalized file object
 */
function normalizeFile(item, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting file normalization', {
            fileId: item?.id?.substring(0, 20) + '...' || 'unknown',
            fileName: item?.name || 'unknown',
            isFolder: !!item?.folder,
            userId: userId || 'anonymous',
            sessionId: sessionId || 'no-session',
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    if (!item || typeof item !== 'object') {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'graph',
            'Invalid file object for normalization',
            'error',
            {
                itemType: typeof item,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('File normalization failed - invalid input', {
                error: 'Invalid file object provided',
                itemType: typeof item,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('File normalization failed - invalid input', {
                sessionId: sessionId,
                error: 'Invalid file object provided',
                itemType: typeof item,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw error;
    }
    try {
        const normalizedFile = {
            id: item.id,
            type: 'file',
            name: item.name,
            description: typeof item.description === 'string' ? item.description : undefined,
            size: typeof item.size === 'number' ? item.size : undefined,
            isFolder: !!item.folder,
            isFile: !!item.file,
            webUrl: typeof item.webUrl === 'string' ? item.webUrl : undefined,
            parentId: item.parentReference && item.parentReference.id ? item.parentReference.id : undefined,
            lastModified: item.lastModifiedDateTime || null,
            created: item.createdDateTime || null,
            mimeType: item.file && item.file.mimeType ? item.file.mimeType : undefined,
            shared: item.shared ? {
              scope: item.shared.scope,
              owner: item.shared.owner ? {
                user: item.shared.owner.user ? {
                  id: item.shared.owner.user.id,
                  displayName: item.shared.owner.user.displayName
                } : undefined,
                // Add other owner types if needed (e.g., group, application)
              } : undefined,
              sharedDateTime: item.shared.sharedDateTime,
              shareId: item.shared.shareId,
              // Normalize link details if available (may vary based on link type)
              link: item.shared.link ? {
                type: item.shared.link.type,
                scope: item.shared.link.scope,
                webUrl: item.shared.link.webUrl
                // Add other link properties as needed
              } : undefined
            } : undefined,
            sharepointIds: item.sharepointIds ? {
              listId: item.sharepointIds.listId,
              listItemId: item.sharepointIds.listItemId,
              listItemUniqueId: item.sharepointIds.listItemUniqueId,
              siteId: item.sharepointIds.siteId,
              siteUrl: item.sharepointIds.siteUrl,
              tenantId: item.sharepointIds.tenantId,
              webId: item.sharepointIds.webId
            } : undefined
        };
        
        // Track performance metric
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('file_normalization_time', executionTime, {
            fileId: item.id || 'unknown',
            isFolder: !!item.folder,
            userId: userId || 'anonymous',
            timestamp: new Date().toISOString()
        }, userId, null, false, sessionId);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('File normalized successfully', {
                fileId: item.id?.substring(0, 20) + '...' || 'unknown',
                fileName: item.name || 'unknown',
                isFolder: !!item.folder,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.info('File normalized successfully', {
                sessionId: sessionId,
                fileId: item.id?.substring(0, 20) + '...' || 'unknown',
                fileName: item.name || 'unknown',
                isFolder: !!item.folder,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return normalizedFile;
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Failed to normalize file: ${error.message || 'Unknown error'}`,
            'error',
            {
                fileId: item.id || 'unknown',
                fileName: item.name || 'unknown',
                originalError: error.message,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('File normalization failed', {
                error: error.message || 'Unknown error',
                fileId: item.id?.substring(0, 20) + '...' || 'unknown',
                fileName: item.name || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('File normalization failed', {
                sessionId: sessionId,
                error: error.message || 'Unknown error',
                fileId: item.id?.substring(0, 20) + '...' || 'unknown',
                fileName: item.name || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
}

/**
 * Normalizes a Microsoft Graph calendar event object to MCP format.
 * @param {object} event - Raw event object from Graph API
 * @param {string} [userId] - User ID for context tracking
 * @param {string} [sessionId] - Session ID for context tracking
 * @returns {object} Normalized event object
 */
function normalizeEvent(event, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting event normalization', {
            eventId: event?.id?.substring(0, 20) + '...' || 'unknown',
            subject: event?.subject?.substring(0, 50) + '...' || 'untitled',
            userId: userId || 'anonymous',
            sessionId: sessionId || 'no-session',
            timestamp: new Date().toISOString()
        }, 'calendar');
    }
    
    // Enhanced validation with detailed error message
    if (!event) {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'calendar',
            'Event object is null or undefined',
            'error',
            {
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Event normalization failed - null event', {
                error: 'Event object is null or undefined',
                timestamp: new Date().toISOString()
            }, 'calendar', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Event normalization failed - null event', {
                sessionId: sessionId,
                error: 'Event object is null or undefined',
                timestamp: new Date().toISOString()
            }, 'calendar');
        }
        
        throw error;
    }
    
    if (typeof event !== 'object') {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'calendar',
            `Invalid event object type: ${typeof event}`,
            'error',
            {
                eventType: typeof event,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Event normalization failed - invalid type', {
                error: `Invalid event object type: ${typeof event}`,
                eventType: typeof event,
                timestamp: new Date().toISOString()
            }, 'calendar', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Event normalization failed - invalid type', {
                sessionId: sessionId,
                error: `Invalid event object type: ${typeof event}`,
                eventType: typeof event,
                timestamp: new Date().toISOString()
            }, 'calendar');
        }
        
        throw error;
    }
    
    try {
        // Handle the case where the event is a response from Graph API or a client-side event
        const normalizedEvent = {
            id: event.id || `temp-${Date.now()}`,
            type: 'event',
            subject: event.subject || 'Untitled Event',
            
            // Normalize start/end with robust fallbacks
            start: {
                dateTime: event.start?.dateTime || (event.start || null),
                timeZone: event.start?.timeZone || 'UTC'
            },
            
            end: {
                dateTime: event.end?.dateTime || (event.end || null),
                timeZone: event.end?.timeZone || 'UTC'
            },
            
            // Backward compatibility fields with robust fallbacks
            startTime: event.start?.dateTime || (event.start || null),
            startTimeZone: event.start?.timeZone || 'UTC',
            endTime: event.end?.dateTime || (event.end || null),
            endTimeZone: event.end?.timeZone || 'UTC',
            
            // Handle location with fallbacks
            location: event.location ? {
                displayName: event.location.displayName || '',
                address: event.location.address ? 
                    (typeof event.location.address === 'string' ? 
                        event.location.address : 
                        JSON.stringify(event.location.address)) : 
                    undefined,
                coordinates: event.location.coordinates || undefined
            } : undefined,
            
            // Handle organizer with fallbacks
            organizer: event.organizer ? {
                name: event.organizer.emailAddress?.name || event.organizer.name || '',
                email: event.organizer.emailAddress?.address || event.organizer.email || ''
            } : undefined,
            
            // Handle attendees with robust parsing
            attendees: Array.isArray(event.attendees) ? 
                event.attendees.map(att => {
                    // Handle different attendee formats
                    if (!att) return null;
                    
                    // Handle direct email string format
                    if (typeof att === 'string') {
                        return {
                            email: att,
                            name: att.split('@')[0] || '',
                            type: 'required',
                            status: 'notResponded',
                            responseTime: null
                        };
                    }
                    
                    // Handle object with email property
                    if (att.email) {
                        return {
                            email: att.email,
                            name: att.name || att.email.split('@')[0] || '',
                            type: att.type || 'required',
                            status: att.status || 'notResponded',
                            responseTime: null
                        };
                    }
                    
                    // Handle Graph API format
                    if (att.emailAddress) {
                        return {
                            email: att.emailAddress.address || '',
                            name: att.emailAddress.name || '',
                            type: att.type || 'required',
                            status: (att.status?.response) || 'notResponded',
                            responseTime: (att.status?.time) || null
                        };
                    }
                    
                    return null;
                }).filter(Boolean) : 
                [],
            
            // Boolean properties with fallbacks
            isAllDay: !!event.isAllDay,
            isCancelled: !!event.isCancelled,
            isOnlineMeeting: !!event.isOnlineMeeting,
            
            // Optional properties
            onlineMeetingUrl: event.onlineMeeting?.joinUrl || event.onlineMeetingUrl,
            recurrence: event.recurrence,
            importance: event.importance || 'normal',
            webLink: event.webLink,
            
            // Handle body content
            body: event.body?.content || event.body || '',
            bodyType: event.body?.contentType || 'text',
            
            // Preview with fallbacks
            preview: event.bodyPreview ? 
                event.bodyPreview.substring(0, 150) : 
                (typeof event.body === 'string' ? 
                    event.body.substring(0, 150) : 
                    (event.body?.content ? 
                        event.body.content.substring(0, 150) : 
                        '')),
            
            // Timestamps
            created: event.createdDateTime || null,
            lastModified: event.lastModifiedDateTime || null
        };
        
        // Track performance metric
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('event_normalization_time', executionTime, {
            eventId: event.id || 'unknown',
            userId: userId || 'anonymous',
            timestamp: new Date().toISOString()
        }, userId, null, false, sessionId);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Event normalized successfully', {
                eventId: event.id?.substring(0, 20) + '...' || 'unknown',
                subject: event.subject?.substring(0, 50) + '...' || 'Untitled Event',
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'calendar', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Event normalized successfully', {
                sessionId: sessionId,
                eventId: event.id?.substring(0, 20) + '...' || 'unknown',
                subject: event.subject?.substring(0, 50) + '...' || 'Untitled Event',
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'calendar');
        }
        
        return normalizedEvent;
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'calendar',
            `Failed to normalize event: ${error.message || 'Unknown error'}`,
            'error',
            {
                eventId: event.id || 'unknown',
                subject: event.subject || 'Untitled Event',
                originalError: error.stack || error.toString(),
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Event normalization failed', {
                error: error.message || 'Unknown error',
                eventId: event.id?.substring(0, 20) + '...' || 'unknown',
                subject: event.subject?.substring(0, 50) + '...' || 'Untitled Event',
                timestamp: new Date().toISOString()
            }, 'calendar', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Event normalization failed', {
                sessionId: sessionId,
                error: error.message || 'Unknown error',
                eventId: event.id?.substring(0, 20) + '...' || 'unknown',
                subject: event.subject?.substring(0, 50) + '...' || 'Untitled Event',
                timestamp: new Date().toISOString()
            }, 'calendar');
        }
        
        throw mcpError;
    }
}

/**
 * Normalizes a Microsoft Graph user object to MCP user profile format.
 * @param {object} user - Raw user object from Graph API
 * @param {object} [options] - Normalization options
 * @param {boolean} [options.strictPrivacy=false] - If true, removes PII fields (phones, address).
 * @param {string} [userId] - User ID for context tracking
 * @param {string} [sessionId] - Session ID for context tracking
 * @returns {object} Normalized user profile
 */
function normalizeUser(user, options = {}, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting user normalization', {
            targetUserId: user?.id?.substring(0, 20) + '...' || 'unknown',
            userPrincipalName: user?.userPrincipalName || 'unknown',
            strictPrivacy: options.strictPrivacy || false,
            userId: userId || 'anonymous',
            sessionId: sessionId || 'no-session',
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    if (!user || typeof user !== 'object') {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'graph',
            'Invalid user object for normalization',
            'error',
            {
                userType: typeof user,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('User normalization failed - invalid input', {
                error: 'Invalid user object provided',
                userType: typeof user,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('User normalization failed - invalid input', {
                sessionId: sessionId,
                error: 'Invalid user object provided',
                userType: typeof user,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw error;
    }
    
    try {
        const { strictPrivacy = false } = options;

        const normalized = {
            id: user.id,
            type: 'user',
            displayName: user.displayName || '',
            givenName: user.givenName || '',
            surname: user.surname || '',
            email: user.mail || user.userPrincipalName || '',
            userPrincipalName: user.userPrincipalName || '',
            jobTitle: user.jobTitle || '',
            department: user.department || '',
            mobilePhone: user.mobilePhone || '',
            businessPhones: Array.isArray(user.businessPhones) ? user.businessPhones : [],
            streetAddress: typeof user.streetAddress === 'string' ? user.streetAddress : undefined,
            city: typeof user.city === 'string' ? user.city : undefined,
            state: typeof user.state === 'string' ? user.state : undefined,
            postalCode: typeof user.postalCode === 'string' ? user.postalCode : undefined,
            country: typeof user.country === 'string' ? user.country : undefined,
            officeLocation: user.officeLocation,
            preferredLanguage: user.preferredLanguage,
            // TODO: Consider normalizing photo differently (e.g., base64 or direct URL if possible)
            photo: user.photo && user.photo['@odata.mediaEditLink'] ? user.photo['@odata.mediaEditLink'] : undefined
            // Deprecated/removed: companyName, country, city are now conditionally removed
        };

        // Remove PII (address/phones) when `strictPrivacy` flag on.
        if (strictPrivacy) {
            delete normalized.mobilePhone;
            delete normalized.businessPhones;
            delete normalized.streetAddress;
            delete normalized.city;
            delete normalized.state;
            delete normalized.postalCode;
            delete normalized.country;
            
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Applied privacy restrictions to user data', {
                    targetUserId: user.id?.substring(0, 20) + '...' || 'unknown',
                    userId: userId || 'anonymous',
                    sessionId: sessionId || 'no-session',
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
        }
        
        // Track performance metric
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('user_normalization_time', executionTime, {
            targetUserId: user.id || 'unknown',
            userId: userId || 'anonymous',
            timestamp: new Date().toISOString()
        }, userId, null, false, sessionId);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('User normalized successfully', {
                targetUserId: user.id?.substring(0, 20) + '...' || 'unknown',
                userPrincipalName: user.userPrincipalName || 'unknown',
                strictPrivacy: strictPrivacy,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.info('User normalized successfully', {
                sessionId: sessionId,
                targetUserId: user.id?.substring(0, 20) + '...' || 'unknown',
                userPrincipalName: user.userPrincipalName || 'unknown',
                strictPrivacy: strictPrivacy,
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return normalized;
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Failed to normalize user: ${error.message || 'Unknown error'}`,
            'error',
            {
                targetUserId: user.id || 'unknown',
                userPrincipalName: user.userPrincipalName || 'unknown',
                originalError: error.message,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('User normalization failed', {
                error: error.message || 'Unknown error',
                targetUserId: user.id?.substring(0, 20) + '...' || 'unknown',
                userPrincipalName: user.userPrincipalName || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('User normalization failed', {
                sessionId: sessionId,
                error: error.message || 'Unknown error',
                targetUserId: user.id?.substring(0, 20) + '...' || 'unknown',
                userPrincipalName: user.userPrincipalName || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
}

/**
 * Normalizes a Microsoft Graph person object to MCP format.
 * @param {object} person - Raw person object from Graph API
 * @param {string} [userId] - User ID for context tracking
 * @param {string} [sessionId] - Session ID for context tracking
 * @returns {object} Normalized person object
 */
function normalizePerson(person, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting person normalization', {
            personId: person?.id?.substring(0, 20) + '...' || 'unknown',
            displayName: person?.displayName || 'unknown',
            userId: userId || 'anonymous',
            sessionId: sessionId || 'no-session',
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    if (!person || typeof person !== 'object') {
        // Pattern 3: Infrastructure Error Logging
        const error = ErrorService.createError(
            'graph',
            'Invalid person object for normalization',
            'error',
            {
                personType: typeof person,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(error);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Person normalization failed - invalid input', {
                error: 'Invalid person object provided',
                personType: typeof person,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Person normalization failed - invalid input', {
                sessionId: sessionId,
                error: 'Invalid person object provided',
                personType: typeof person,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw error;
    }
    
    try {
        // Extract primary email from scoredEmailAddresses if available
        let primaryEmail = '';
        if (Array.isArray(person.scoredEmailAddresses) && person.scoredEmailAddresses.length > 0) {
            // Sort by relevance score (highest first) and take the first one
            const sortedEmails = [...person.scoredEmailAddresses].sort((a, b) => {
                const scoreA = typeof a.relevanceScore === 'number' ? a.relevanceScore : 0;
                const scoreB = typeof b.relevanceScore === 'number' ? b.relevanceScore : 0;
                return scoreB - scoreA;
            });
            primaryEmail = sortedEmails[0].address;
        }
        
        // Extract primary phone if available
        let primaryPhone = '';
        if (Array.isArray(person.phones) && person.phones.length > 0) {
            // Prefer business phones if available
            const businessPhone = person.phones.find(p => p.type === 'business');
            primaryPhone = businessPhone ? businessPhone.number : person.phones[0].number;
        }
        
        const normalizedPerson = {
            id: person.id,
            type: 'person',
            displayName: person.displayName || '',
            givenName: person.givenName || '',
            surname: person.surname || '',
            email: primaryEmail,
            phone: primaryPhone,
            jobTitle: person.jobTitle || '',
            companyName: person.companyName || '',
            department: person.department || '',
            officeLocation: person.officeLocation || '',
            userPrincipalName: person.userPrincipalName || '',
            imAddress: person.imAddress || '',
            scoredEmailAddresses: (Array.isArray(person.scoredEmailAddresses) && person.scoredEmailAddresses.length > 0) ? person.scoredEmailAddresses.map(email => ({
                address: email.address,
                relevanceScore: typeof email.relevanceScore === 'number' ? email.relevanceScore : 0
            })) : [],
            phones: (Array.isArray(person.phones) && person.phones.length > 0) ? person.phones.map(phone => ({
                type: phone.type,
                number: phone.number
            })) : [],
            personType: person.personType ? {
                class: person.personType.class,
                subclass: person.personType.subclass
            } : { class: 'Person', subclass: 'Unknown' },
            relevanceScore: typeof person.relevanceScore === 'number' ? person.relevanceScore : 0,
            isFavorite: !!person.isFavorite
        };
        
        // Track performance metric
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('person_normalization_time', executionTime, {
            personId: person.id || 'unknown',
            userId: userId || 'anonymous',
            timestamp: new Date().toISOString()
        }, userId, null, false, sessionId);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Person normalized successfully', {
                personId: person.id?.substring(0, 20) + '...' || 'unknown',
                displayName: person.displayName || 'unknown',
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Person normalized successfully', {
                sessionId: sessionId,
                personId: person.id?.substring(0, 20) + '...' || 'unknown',
                displayName: person.displayName || 'unknown',
                executionTime: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return normalizedPerson;
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Failed to normalize person: ${error.message || 'Unknown error'}`,
            'error',
            {
                personId: person.id || 'unknown',
                displayName: person.displayName || 'unknown',
                originalError: error.message,
                userId: userId || 'anonymous',
                sessionId: sessionId || 'no-session',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Person normalization failed', {
                error: error.message || 'Unknown error',
                personId: person.id?.substring(0, 20) + '...' || 'unknown',
                displayName: person.displayName || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Person normalization failed', {
                sessionId: sessionId,
                error: error.message || 'Unknown error',
                personId: person.id?.substring(0, 20) + '...' || 'unknown',
                displayName: person.displayName || 'unknown',
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
}

// ============================================================================
// DEEP LINK URL BUILDERS
// ============================================================================

/**
 * Builds a deep link URL to an Outlook email item.
 * @param {string} itemId - The email item ID
 * @returns {string|null} Outlook web URL or null if itemId is missing
 */
function buildOutlookItemUrl(itemId) {
    if (!itemId) return null;
    const encodedId = encodeURIComponent(itemId);
    return `https://outlook.office365.com/owa/?ItemID=${encodedId}&exvsurl=1&viewmodel=ReadMessageItem`;
}

/**
 * Builds a deep link URL to a Teams channel message.
 * @param {string} channelId - The channel ID (format: 19:xxx@thread.tacv2)
 * @param {string} messageId - The message ID
 * @param {string} [tenantId] - The tenant ID (optional)
 * @param {string} [groupId] - The group/team ID (optional)
 * @returns {string|null} Teams deep link URL or null if required params are missing
 */
function buildTeamsMessageUrl(channelId, messageId, tenantId, groupId) {
    if (!channelId || !messageId) return null;

    let url = `https://teams.microsoft.com/l/message/${encodeURIComponent(channelId)}/${encodeURIComponent(messageId)}`;
    const params = [];

    if (tenantId) params.push(`tenantId=${encodeURIComponent(tenantId)}`);
    if (groupId) params.push(`groupId=${encodeURIComponent(groupId)}`);

    if (params.length > 0) {
        url += '?' + params.join('&');
    }

    return url;
}

/**
 * Builds a deep link URL to a Teams chat.
 * @param {string} chatId - The chat ID
 * @returns {string|null} Teams chat URL or null if chatId is missing
 */
function buildTeamsChatUrl(chatId) {
    if (!chatId) return null;
    return `https://teams.microsoft.com/l/chat/${encodeURIComponent(chatId)}`;
}

/**
 * Builds a deep link URL to a SharePoint/OneDrive file.
 * @param {string} webUrl - The file's web URL from Graph API
 * @returns {string|null} The web URL or null if missing
 */
function buildSharePointFileUrl(webUrl) {
    // Graph API already provides webUrl, just validate and return
    if (!webUrl) return null;
    return webUrl;
}

/**
 * Builds a deep link URL to a calendar event.
 * @param {string} eventId - The calendar event ID
 * @returns {string|null} Outlook calendar URL or null if eventId is missing
 */
function buildCalendarEventUrl(eventId) {
    if (!eventId) return null;
    const encodedId = encodeURIComponent(eventId);
    return `https://outlook.office365.com/owa/?ItemID=${encodedId}&exvsurl=1&viewmodel=ReadMessageItem`;
}

/**
 * Builds a Microsoft Search deep link for a query.
 * @param {string} query - The search query
 * @param {string} [entityType] - Optional entity type filter (person, message, event, file)
 * @returns {string} Microsoft Search URL
 */
function buildSearchUrl(query, entityType) {
    if (!query) return null;
    let url = `https://www.office.com/search?q=${encodeURIComponent(query)}`;
    if (entityType) {
        // Map entity types to Office search entity filters
        const entityMap = {
            'person': 'People',
            'message': 'Mail',
            'event': 'Events',
            'driveItem': 'Files',
            'file': 'Files'
        };
        const filter = entityMap[entityType] || entityType;
        url += `&filters=EntityType:${filter}`;
    }
    return url;
}

module.exports = {
    normalizeEmail,
    normalizeFile,
    normalizeEvent,
    normalizeUser,
    normalizePerson,
    // URL builders
    buildOutlookItemUrl,
    buildTeamsMessageUrl,
    buildTeamsChatUrl,
    buildSharePointFileUrl,
    buildCalendarEventUrl,
    buildSearchUrl
};
