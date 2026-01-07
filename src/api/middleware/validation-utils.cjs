/**
 * @fileoverview Shared validation utilities for API controllers.
 * Provides common validation functions using Joi and logging integration.
 */

const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Validate request data and log validation results.
 * @param {object} req - Express request
 * @param {object} schema - Joi validation schema
 * @param {string} operation - Operation name for logging
 * @param {string|object} categoryOrUserContext - Logging category string OR user context object (backwards compatible)
 * @param {object} [userContext] - User context containing userId and deviceId (if category provided)
 * @returns {object} Validation result with error and value properties
 */
function validateAndLog(req, schema, operation, categoryOrUserContext, userContext) {
    // Backwards compatibility: if 4th param is an object, it's userContext and category defaults to 'api'
    let category = 'api';
    let context = {};

    if (typeof categoryOrUserContext === 'string') {
        category = categoryOrUserContext;
        context = userContext || {};
    } else if (typeof categoryOrUserContext === 'object') {
        context = categoryOrUserContext || {};
    }

    const { userId = null, deviceId = null } = context;

    // For POST/PUT/PATCH requests, validate body; for GET/DELETE requests, validate query
    const dataToValidate = ['POST', 'PUT', 'PATCH'].includes(req.method)
        ? req.body
        : req.query;

    const { error, value } = schema.validate(dataToValidate);

    if (error) {
        MonitoringService.warn(`Validation failed for ${operation}`, {
            operation,
            error: error.details[0].message,
            userId,
            deviceId
        }, category, null, userId, deviceId);
    } else {
        MonitoringService.debug(`Validation passed for ${operation}`, {
            operation,
            userId,
            deviceId
        }, category, null, userId, deviceId);
    }

    return { error, value };
}

/**
 * Create a standard validation error response.
 * Uses OAuth-style error format for consistency.
 * @param {object} validationError - Joi validation error
 * @returns {object} Standardized error response
 */
function createValidationErrorResponse(validationError) {
    return {
        error: 'INVALID_REQUEST',
        error_description: validationError.details[0].message,
        details: validationError.details
    };
}

/**
 * Validate request and return early response if validation fails.
 * Combines validateAndLog and response handling.
 * @param {object} req - Express request
 * @param {object} res - Express response
 * @param {object} schema - Joi validation schema
 * @param {string} operation - Operation name for logging
 * @param {string} category - Logging category
 * @param {object} userContext - User context
 * @returns {object|null} Validated data or null if validation failed (response already sent)
 */
function validateRequest(req, res, schema, operation, category, userContext = {}) {
    const { error, value } = validateAndLog(req, schema, operation, category, userContext);

    if (error) {
        res.status(400).json(createValidationErrorResponse(error));
        return null;
    }

    return value;
}

module.exports = {
    validateAndLog,
    createValidationErrorResponse,
    validateRequest
};
