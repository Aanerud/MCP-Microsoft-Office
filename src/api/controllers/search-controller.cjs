/**
 * @fileoverview Search Controller - Handles unified search API requests.
 * Provides cross-entity search across Microsoft 365 (emails, events, files, people).
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const { validateAndLog } = require('../middleware/validation-utils.cjs');

/**
 * Joi validation schemas for search endpoints
 */
const schemas = {
    search: Joi.object({
        query: Joi.string().min(1).max(500).required(),
        entityTypes: Joi.array()
            .items(Joi.string().valid('message', 'chatMessage', 'event', 'driveItem', 'person', 'site', 'list', 'listItem'))
            .min(1)
            .max(8)
            .optional()
            .default(['message', 'event', 'driveItem', 'person']),
        limit: Joi.number().integer().min(1).max(25).optional().default(10),
        from: Joi.number().integer().min(0).optional().default(0)
    })
};

/**
 * Creates a search controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.searchModule - Initialized search module
 * @returns {object} Controller methods
 */
function createSearchController({ searchModule }) {
    if (!searchModule) {
        throw new Error('Search module is required for SearchController');
    }

    return {
        /**
         * Unified search across Microsoft 365 content.
         * Supports searching emails, calendar events, files, and people.
         * @param {object} req - Express request
         * @param {object} res - Express response
         */
        async search(req, res) {
            // Extract user context from auth middleware
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;

            const startTime = Date.now();
            try {
                // Pattern 1: Development Debug Logs
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Processing unified search request', {
                        method: req.method,
                        path: req.path,
                        sessionId,
                        userAgent: req.get('User-Agent'),
                        timestamp: new Date().toISOString(),
                        userId,
                        deviceId
                    }, 'search');
                }

                // Validate request (supports both GET and POST)
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.search,
                    'search',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message,
                        details: validationError.details
                    });
                }

                const options = {
                    query: validatedData.query,
                    entityTypes: validatedData.entityTypes,
                    limit: validatedData.limit,
                    from: validatedData.from
                };

                // Execute search through module
                const results = await searchModule.search(options, req, userId, sessionId);

                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Unified search completed successfully', {
                        query: options.query?.substring(0, 30),
                        entityTypes: options.entityTypes,
                        resultCount: results.results?.length || 0,
                        totalHits: results.pagination?.total || 0,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'search', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Unified search completed with session', {
                        sessionId,
                        query: options.query?.substring(0, 30),
                        entityTypes: options.entityTypes,
                        resultCount: results.results?.length || 0,
                        totalHits: results.pagination?.total || 0,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'search');
                }

                // Track performance with user context
                const duration = Date.now() - startTime;
                MonitoringService.trackMetric('search.unified.duration', duration, {
                    resultCount: results.results?.length || 0,
                    entityTypes: options.entityTypes?.join(','),
                    success: true,
                    userId,
                    deviceId
                });

                res.json(results);
            } catch (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'search',
                    'Unified search failed',
                    'error',
                    {
                        endpoint: '/api/v1/search',
                        error: error.message,
                        stack: error.stack,
                        operation: 'search',
                        query: req.body?.query?.substring(0, 50) || req.query?.query?.substring(0, 50),
                        userId,
                        deviceId,
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);

                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Unified search failed', {
                        error: error.message,
                        operation: 'search',
                        timestamp: new Date().toISOString()
                    }, 'search', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Unified search failed', {
                        sessionId,
                        error: error.message,
                        operation: 'search',
                        timestamp: new Date().toISOString()
                    }, 'search');
                }

                // Track error metrics with user context
                const duration = Date.now() - startTime;
                MonitoringService.trackMetric('search.unified.error', 1, {
                    errorMessage: error.message,
                    duration,
                    success: false,
                    userId,
                    deviceId
                });

                res.status(500).json({
                    error: 'SEARCH_FAILED',
                    error_description: 'Failed to execute unified search'
                });
            }
        }
    };
}

module.exports = createSearchController;
