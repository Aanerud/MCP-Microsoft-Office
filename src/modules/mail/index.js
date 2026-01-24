/**
 * @fileoverview MCP Mail Module - Handles mail-related intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system and phase1_architecture.md.
 */

const { normalizeEmail } = require('../../graph/normalizers.cjs');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const MAIL_CAPABILITIES = [
    'readMail',
    'searchMail',
    'sendMail',
    'flagMail',
    'getMailAttachments',
    'readMailDetails',
    'markEmailRead',
    'addMailAttachment',
    'removeMailAttachment'
];

// Log module initialization
MonitoringService.info('Mail Module initialized', {
    serviceName: 'mail-module',
    capabilities: MAIL_CAPABILITIES.length,
    timestamp: new Date().toISOString()
}, 'mail');

const MailModule = {
    /**
     * Helper method to redact sensitive email data from objects before logging
     * @param {object} data - The data object to redact
     * @returns {object} Redacted copy of the data
     * @private
     */
    redactSensitiveEmailData(data) {
        if (!data || typeof data !== 'object') {
            return data;
        }
        
        // Create a deep copy to avoid modifying the original
        const result = Array.isArray(data) ? [...data] : {...data};
        
        // Fields that should be redacted for email data
        const sensitiveFields = [
            'body', 'content', 'subject', 'to', 'from', 'cc', 'bcc', 
            'emailAddress', 'address', 'email', 'recipients', 'sender',
            'attachment', 'attachments', 'contentBytes'
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
                // Recursively process nested objects
                else if (typeof result[key] === 'object' && result[key] !== null) {
                    result[key] = this.redactSensitiveEmailData(result[key]);
                }
            }
        }
        
        return result;
    },

    /**
     * Fetch raw inbox data from Graph for debugging (no normalization)
     * @param {object} options
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<object[]>}
     */
    async getInboxRaw(options = {}, userId, sessionId) {
        const startTime = Date.now();
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting raw inbox data', {
                    options: this.redactSensitiveEmailData(options),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.getInboxRaw !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.getInboxRaw not implemented',
                    'error',
                    {
                        method: 'getInboxRaw',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get raw inbox data', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get raw inbox data', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.getInboxRaw(options, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Raw inbox data retrieved successfully', {
                    itemCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Raw inbox data retrieved with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    itemCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get raw inbox data', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get raw inbox data', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error getting raw inbox data: ${error.message}`,
                'error',
                {
                    method: 'getInboxRaw',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to get raw inbox data', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to get raw inbox data', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Get inbox emails
     * @param {object} options - Options including top, filter
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Array<object>>} List of emails
     */
    async getInbox(options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting inbox emails', {
                    options: this.redactSensitiveEmailData(options),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.getInbox !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.getInbox not implemented',
                    'error',
                    {
                        method: 'getInbox',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get inbox emails', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get inbox emails', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.getInbox(options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Inbox emails retrieved successfully', {
                    emailCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Inbox emails retrieved with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    emailCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get inbox emails', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get inbox emails', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error getting inbox emails: ${error.message}`,
                'error',
                {
                    method: 'getInbox',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to get inbox emails', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to get inbox emails', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Search emails by query
     * @param {string} query - Search query
     * @param {object} options - Search options
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Array<object>>} List of matching emails
     */
    async searchEmails(query, options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Searching emails', {
                    query: query ? query.substring(0, 50) + '...' : 'empty',
                    options: this.redactSensitiveEmailData(options),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.searchEmails !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.searchEmails not implemented',
                    'error',
                    {
                        method: 'searchEmails',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to search emails', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to search emails', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.searchEmails(query, options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email search completed successfully', {
                    resultCount: Array.isArray(result) ? result.length : 0,
                    queryLength: query ? query.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email search completed with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    resultCount: Array.isArray(result) ? result.length : 0,
                    queryLength: query ? query.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to search emails', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to search emails', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error searching emails: ${error.message}`,
                'error',
                {
                    method: 'searchEmails',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to search emails', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to search emails', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Send an email
     * @param {object} emailData - Email data with to, subject, body, cc, bcc
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<boolean>} Success indicator
     */
    async sendEmail(emailData, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Sending email', {
                    emailData: this.redactSensitiveEmailData(emailData),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.sendEmail !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.sendEmail not implemented',
                    'error',
                    {
                        method: 'sendEmail',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to send email', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to send email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            // Validate required fields
            if (!emailData.to) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'Recipient (to) is required',
                    'warn',
                    {
                        method: 'sendEmail',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to send email', {
                        error: 'Recipient is required',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to send email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'Recipient is required',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            if (!emailData.subject) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'Subject is required',
                    'warn',
                    {
                        method: 'sendEmail',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to send email', {
                        error: 'Subject is required',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to send email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'Subject is required',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            if (!emailData.body) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'Body is required',
                    'warn',
                    {
                        method: 'sendEmail',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to send email', {
                        error: 'Body is required',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to send email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'Body is required',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            // Send email via graph service
            const result = await graphService.sendEmail(emailData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email sent successfully', {
                    hasAttachments: emailData.attachments && emailData.attachments.length > 0,
                    recipientCount: Array.isArray(emailData.to) ? emailData.to.length : 1,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email sent with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    hasAttachments: emailData.attachments && emailData.attachments.length > 0,
                    recipientCount: Array.isArray(emailData.to) ? emailData.to.length : 1,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to send email', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to send email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error sending email: ${error.message}`,
                'error',
                {
                    method: 'sendEmail',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to send email', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to send email', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Flag or unflag an email
     * @param {string} id - Email ID
     * @param {boolean} flag - Flag state
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<boolean>} Success indicator
     */
    async flagEmail(id, flag = true, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Flagging email', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    flag: flag,
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.flagEmail !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.flagEmail not implemented',
                    'error',
                    {
                        method: 'flagEmail',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to flag email', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to flag email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.flagEmail(id, flag, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email flagged successfully', {
                    flagState: flag,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email flagged with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    flagState: flag,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to flag email', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to flag email', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error flagging email: ${error.message}`,
                'error',
                {
                    method: 'flagEmail',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to flag email', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to flag email', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Get email attachments
     * @param {string} id - Email ID
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Array<object>>} List of attachments
     */
    async getAttachments(id, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting email attachments', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.getAttachments !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.getAttachments not implemented',
                    'error',
                    {
                        method: 'getAttachments',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get email attachments', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get email attachments', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.getAttachments(id, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email attachments retrieved successfully', {
                    attachmentCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email attachments retrieved with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    attachmentCount: Array.isArray(result) ? result.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get email attachments', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get email attachments', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error getting email attachments: ${error.message}`,
                'error',
                {
                    method: 'getAttachments',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to get email attachments', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to get email attachments', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Get detailed information for a specific email
     * @param {string} id - Email ID
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<object>} Email details
     */
    async getEmailDetails(id, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting email details', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.getEmailDetails !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.getEmailDetails not implemented',
                    'error',
                    {
                        method: 'getEmailDetails',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get email details', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get email details', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.getEmailDetails(id, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email details retrieved successfully', {
                    hasAttachments: result && result.hasAttachments,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email details retrieved with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    hasAttachments: result && result.hasAttachments,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to get email details', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to get email details', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error getting email details: ${error.message}`,
                'error',
                {
                    method: 'getEmailDetails',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to get email details', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to get email details', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Mark an email as read or unread
     * @param {string} id - Email ID
     * @param {boolean} isRead - Read status to set
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<boolean>} Success indicator
     */
    async markAsRead(id, isRead = true, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Marking email as read/unread', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    isRead: isRead,
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.markAsRead !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.markAsRead not implemented',
                    'error',
                    {
                        method: 'markAsRead',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to mark email as read', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to mark email as read', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.markAsRead(id, isRead, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email read status updated successfully', {
                    readStatus: isRead,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email read status updated with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    readStatus: isRead,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to mark email as read', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to mark email as read', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error marking email as read: ${error.message}`,
                'error',
                {
                    method: 'markAsRead',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to mark email as read', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to mark email as read', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Add an attachment to an email
     * @param {string} id - Email ID
     * @param {object} attachment - Attachment data
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<boolean>} Success indicator
     */
    async addMailAttachment(id, attachment, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Adding email attachment', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    attachmentName: attachment?.name || 'unnamed',
                    attachmentSize: attachment?.size || 0,
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.addMailAttachment !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.addMailAttachment not implemented',
                    'error',
                    {
                        method: 'addMailAttachment',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to add email attachment', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to add email attachment', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.addMailAttachment(id, attachment, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email attachment added successfully', {
                    attachmentName: attachment?.name || 'unnamed',
                    attachmentSize: attachment?.size || 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email attachment added with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    attachmentName: attachment?.name || 'unnamed',
                    attachmentSize: attachment?.size || 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to add email attachment', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to add email attachment', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error adding email attachment: ${error.message}`,
                'error',
                {
                    method: 'addMailAttachment',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to add email attachment', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to add email attachment', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Remove an attachment from an email
     * @param {string} id - Email ID
     * @param {string} attachmentId - Attachment ID
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<boolean>} Success indicator
     */
    async removeMailAttachment(id, attachmentId, req, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from req if not provided
        if (!userId && req?.user?.userId) {
            userId = req.user.userId;
        }
        if (!sessionId && req?.session?.id) {
            sessionId = req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Removing email attachment', {
                    emailId: id ? id.substring(0, 20) + '...' : 'none',
                    attachmentId: attachmentId ? attachmentId.substring(0, 20) + '...' : 'none',
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService } = this.services || {};
            if (!graphService || typeof graphService.removeMailAttachment !== 'function') {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'GraphService.removeMailAttachment not implemented',
                    'error',
                    {
                        method: 'removeMailAttachment',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to remove email attachment', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to remove email attachment', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            const result = await graphService.removeMailAttachment(id, attachmentId, req, userId, sessionId);
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Email attachment removed successfully', {
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Email attachment removed with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to remove email attachment', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to remove email attachment', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error removing email attachment: ${error.message}`,
                'error',
                {
                    method: 'removeMailAttachment',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to remove email attachment', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to remove email attachment', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    
    id: 'mail',
    name: 'Outlook Mail',
    capabilities: MAIL_CAPABILITIES,
    /**
     * Initializes the mail module with dependencies.
     * @param {object} services - { graphService, cacheService, eventService }
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {object} Initialized module
     */
    init(services, userId, sessionId) {
        const startTime = Date.now();
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Initializing mail module', {
                    services: Object.keys(services || {}),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            if (!services) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'mail',
                    'Services object is required for mail module initialization',
                    'error',
                    {
                        method: 'init',
                        moduleId: 'mail',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to initialize mail module', {
                        error: 'Services object is required',
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to initialize mail module', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'Services object is required',
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                
                throw mcpError;
            }
            
            this.services = services;
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Mail module initialized successfully', {
                    serviceCount: Object.keys(services).length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Mail module initialized with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    serviceCount: Object.keys(services).length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return this;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to initialize mail module', {
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to initialize mail module', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error initializing mail module: ${error.message}`,
                'error',
                {
                    method: 'init',
                    moduleId: 'mail',
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to initialize mail module', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to initialize mail module', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    },
    /**
     * Handles mail-related intents routed to this module.
     * @param {string} intent
     * @param {object} entities
     * @param {object} context
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<object>} Normalized response
     */
    async handleIntent(intent, entities = {}, context = {}, userId, sessionId) {
        const startTime = Date.now();
        
        // Extract user context from context.req if not provided
        if (!userId && context.req?.user?.userId) {
            userId = context.req.user.userId;
        }
        if (!sessionId && context.req?.session?.id) {
            sessionId = context.req.session.id;
        }
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Handling mail intent', {
                    intent: intent,
                    entities: this.redactSensitiveEmailData(entities),
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            const { graphService, cacheService } = this.services || {};
            let result;
            
            switch (intent) {
                case 'readMail': {
                    const count = entities.count || 10;
                    // Try cache first
                    const cacheKey = `mail:inbox:${count}`;
                    let mailList = cacheService && await cacheService.get(cacheKey);
                    if (!mailList) {
                        const raw = await graphService.getInbox({ top: count }, context.req, userId, sessionId);
                        mailList = Array.isArray(raw) ? raw.map(normalizeEmail) : [];

                        if (cacheService) await cacheService.set(cacheKey, mailList, 60);
                    }
                    result = { type: 'mailList', items: mailList };
                    break;
                }
                case 'searchMail': {
                    const query = entities.query || '';
                    const cacheKey = `mail:search:${query}`;
                    let results = cacheService && await cacheService.get(cacheKey);
                    if (!results) {
                        const raw = await graphService.searchEmails(query, {}, context.req, userId, sessionId);
                        results = Array.isArray(raw) ? raw.map(normalizeEmail) : [];
                        if (cacheService) await cacheService.set(cacheKey, results, 60);
                    }
                    result = { type: 'mailList', items: results };
                    break;
                }
                case 'sendMail': {
                    const { to, subject, body, cc, bcc } = entities;
                    const sent = await graphService.sendEmail({ to, subject, body, cc, bcc }, context.req, userId, sessionId);
                    result = { type: 'mailSendResult', success: !!sent, sent };
                    break;
                }
                case 'flagMail': {
                    const { mailId, flag } = entities;
                    const flagged = await graphService.flagEmail(mailId, flag, context.req, userId, sessionId);
                    result = { type: 'mailFlagResult', flagged };
                    break;
                }
                case 'getMailAttachments': {
                    const { mailId } = entities;
                    const attachments = await graphService.getAttachments(mailId, context.req, userId, sessionId);
                    result = { type: 'mailAttachments', attachments };
                    break;
                }
                case 'readMailDetails': {
                    const { id } = entities;
                    const cacheKey = `mail:details:${id}`;
                    let details = cacheService && await cacheService.get(cacheKey);
                    if (!details) {
                        details = await graphService.getEmailDetails(id, context.req, userId, sessionId);
                        if (cacheService) await cacheService.set(cacheKey, details, 60);
                    }
                    result = { type: 'mailDetails', email: details };
                    break;
                }
                case 'markEmailRead': {
                    const { id, isRead = true } = entities;
                    const success = await graphService.markAsRead(id, isRead, context.req, userId, sessionId);
                    result = { type: 'mailMarkReadResult', success, isRead };
                    break;
                }
                case 'addMailAttachment': {
                    const { id, attachment } = entities;
                    const added = await graphService.addMailAttachment(id, attachment, context.req, userId, sessionId);
                    result = { type: 'mailAttachmentAddResult', added };
                    break;
                }
                case 'removeMailAttachment': {
                    const { id, attachmentId } = entities;
                    const removed = await graphService.removeMailAttachment(id, attachmentId, context.req, userId, sessionId);
                    result = { type: 'mailAttachmentRemoveResult', removed };
                    break;
                }
                default: {
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'mail',
                        `MailModule cannot handle intent: ${intent}`,
                        'error',
                        {
                            method: 'handleIntent',
                            moduleId: 'mail',
                            intent: intent,
                            timestamp: new Date().toISOString()
                        }
                    );
                    MonitoringService.logError(mcpError);
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('Failed to handle mail intent', {
                            intent: intent,
                            error: 'Unsupported intent',
                            timestamp: new Date().toISOString()
                        }, 'mail', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('Failed to handle mail intent', {
                            sessionId: sessionId.substring(0, 8) + '...',
                            intent: intent,
                            error: 'Unsupported intent',
                            timestamp: new Date().toISOString()
                        }, 'mail');
                    }
                    
                    throw mcpError;
                }
            }
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Mail intent handled successfully', {
                    intent: intent,
                    resultType: result?.type || 'unknown',
                    itemCount: result?.items?.length || (result?.attachments?.length || 0),
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Mail intent handled with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    intent: intent,
                    resultType: result?.type || 'unknown',
                    itemCount: result?.items?.length || (result?.attachments?.length || 0),
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // If it's already an MCP error, just track user error and rethrow
            if (error.category) {
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Failed to handle mail intent', {
                        intent: intent,
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Failed to handle mail intent', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        intent: intent,
                        error: error.message,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'mail');
                }
                throw error;
            }
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'mail',
                `Error handling mail intent: ${error.message}`,
                'error',
                {
                    method: 'handleIntent',
                    moduleId: 'mail',
                    intent: intent,
                    stack: error.stack,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Failed to handle mail intent', {
                    intent: intent,
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Failed to handle mail intent', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    intent: intent,
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'mail');
            }
            
            throw mcpError;
        }
    }
};

module.exports = MailModule;
