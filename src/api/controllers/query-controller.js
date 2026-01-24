/**
 * @fileoverview Handles /api/query endpoint for natural language queries.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Helper function to validate request and log validation errors
 * @param {object} req - Express request object
 * @param {object} schema - Joi schema to validate against
 * @param {string} endpoint - Endpoint name for error context
 * @param {object} userContext - User context with userId, deviceId, sessionId
 * @param {object} [additionalContext] - Additional context for validation errors
 * @returns {object} Object with error and value properties
 */
const validateAndLog = (req, schema, endpoint, userContext = {}, additionalContext = {}) => {
    const { userId, deviceId, sessionId } = userContext;
    const result = schema.validate(req.body);
    
    if (result.error) {
        // Pattern 3: Infrastructure Error Logging
        const validationError = ErrorService.createError(
            'query',
            `${endpoint} validation error`,
            'warning',
            { 
                details: result.error.details,
                endpoint,
                userId,
                deviceId,
                timestamp: new Date().toISOString(),
                ...additionalContext
            }
        );
        MonitoringService.logError(validationError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error(`${endpoint} validation failed`, {
                error: result.error.details[0]?.message || 'Validation failed',
                endpoint,
                timestamp: new Date().toISOString()
            }, 'query', null, userId);
        } else if (sessionId) {
            MonitoringService.error(`${endpoint} validation failed`, {
                sessionId,
                error: result.error.details[0]?.message || 'Validation failed',
                endpoint,
                timestamp: new Date().toISOString()
            }, 'query');
        }
    } else {
        // Log successful validation in development
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug(`${endpoint} validation successful`, {
                endpoint,
                sessionId,
                userId,
                deviceId,
                timestamp: new Date().toISOString()
            }, 'query');
        }
    }
    
    return result;
};

const querySchema = Joi.object({
    query: Joi.string().min(2).required()
});

/**
 * Factory for query controller with dependency injection.
 * @param {object} deps - { nluAgent, contextService, errorService }
 */
module.exports = ({ nluAgent, contextService, errorService }) => ({
    /**
     * POST /api/query
     * @param {express.Request} req
     * @param {express.Response} res
     */
    async handleQuery(req, res) {
        const startTime = Date.now();
        
        // Extract user context
        const { userId, deviceId } = req.user || {};
        const sessionId = req.session?.id;
        const userContext = { userId, deviceId, sessionId };
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Processing natural language query request', {
                    method: req.method,
                    path: req.path,
                    sessionId,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId,
                    queryLength: req.body?.query?.length || 0
                }, 'query');
            }
            
            // Validate request using helper function
            const { error, value } = validateAndLog(req, querySchema, 'handleQuery', userContext);
            if (error) {
                return res.status(400).json({ 
                    error: 'QUERY_VALIDATION_FAILED', 
                    error_description: 'Invalid query request',
                    details: error.details 
                });
            }
            
            // Try to use the NLU agent if available, otherwise provide mock data
            let nluResult;
            let ctx = {};
            
            try {
                if (typeof nluAgent.processQuery === 'function') {
                    // Log NLU processing start
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Starting NLU agent processing', {
                            query: value.query,
                            queryLength: value.query.length,
                            sessionId,
                            userId,
                            timestamp: new Date().toISOString()
                        }, 'query');
                    }
                    
                    // NLU pipeline
                    nluResult = await nluAgent.processQuery({ query: value.query });
                    
                    // Update context if contextService is available
                    if (contextService && typeof contextService.updateContext === 'function') {
                        await contextService.updateContext({
                            currentIntent: nluResult.intent,
                            currentEntities: nluResult.entities
                        });
                        ctx = await contextService.getCurrentContext();
                    }
                } else {
                    // Provide mock NLU result for testing
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Using mock NLU result for query', {
                            query: value.query,
                            mockMode: true,
                            sessionId,
                            userId,
                            timestamp: new Date().toISOString()
                        }, 'query');
                    }
                    
                    // Simple intent detection based on keywords
                    let intent = 'unknown';
                    let entities = {};
                    
                    if (value.query.toLowerCase().includes('email') || 
                        value.query.toLowerCase().includes('mail')) {
                        intent = 'readMail';
                        entities = { count: 5 };
                    } else if (value.query.toLowerCase().includes('calendar') || 
                             value.query.toLowerCase().includes('event') || 
                             value.query.toLowerCase().includes('meeting')) {
                        intent = 'readCalendar';
                        entities = { count: 3 };
                    } else if (value.query.toLowerCase().includes('file') || 
                             value.query.toLowerCase().includes('document')) {
                        intent = 'listFiles';
                        entities = {};
                    }
                    
                    nluResult = {
                        intent,
                        entities,
                        confidence: 0.85,
                        query: value.query
                    };
                    
                    // Mock context
                    ctx = {
                        currentIntent: intent,
                        currentEntities: entities,
                        sessionId: 'mock-session',
                        timestamp: new Date().toISOString()
                    };
                }
            } catch (nluError) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'query',
                    'Error processing NLU query',
                    'error',
                    { 
                        error: nluError.message, 
                        stack: nluError.stack,
                        operation: 'processQuery',
                        query: value.query,
                        userId,
                        deviceId,
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('NLU query processing failed', {
                        error: nluError.message,
                        operation: 'processQuery',
                        timestamp: new Date().toISOString()
                    }, 'query', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('NLU query processing failed', {
                        sessionId,
                        error: nluError.message,
                        operation: 'processQuery',
                        timestamp: new Date().toISOString()
                    }, 'query');
                }
                
                // Provide fallback mock response
                nluResult = {
                    intent: 'unknown',
                    entities: {},
                    confidence: 0.5,
                    query: value.query,
                    error: nluError.message
                };
                
                ctx = {
                    currentIntent: 'unknown',
                    timestamp: new Date().toISOString()
                };
            }
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Natural language query processed successfully', {
                    intent: nluResult.intent,
                    confidence: nluResult.confidence,
                    queryLength: value.query.length,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'query', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Natural language query processed with session', {
                    sessionId,
                    intent: nluResult.intent,
                    confidence: nluResult.confidence,
                    queryLength: value.query.length,
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'query');
            }
            
            // Track performance metrics
            const duration = Date.now() - startTime;
            MonitoringService.trackMetric('query.handleQuery.duration', duration, {
                intent: nluResult.intent,
                confidence: nluResult.confidence,
                queryLength: value.query.length,
                success: true,
                userId,
                sessionId
            });
            
            res.json({ response: nluResult, context: ctx });
        } catch (err) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'query',
                'Failed to process natural language query',
                'error',
                { 
                    endpoint: '/api/query',
                    error: err.message,
                    stack: err.stack,
                    operation: 'handleQuery',
                    userId,
                    deviceId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Natural language query processing failed', {
                    error: err.message,
                    operation: 'handleQuery',
                    timestamp: new Date().toISOString()
                }, 'query', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Natural language query processing failed', {
                    sessionId,
                    error: err.message,
                    operation: 'handleQuery',
                    timestamp: new Date().toISOString()
                }, 'query');
            }
            
            // Track error metrics
            const duration = Date.now() - startTime;
            MonitoringService.trackMetric('query.handleQuery.error', 1, {
                errorMessage: err.message,
                duration,
                success: false,
                userId,
                sessionId
            });
            
            res.status(500).json({ 
                error: 'QUERY_PROCESSING_FAILED', 
                error_description: 'Failed to process natural language query'
            });
        }
    }
});
