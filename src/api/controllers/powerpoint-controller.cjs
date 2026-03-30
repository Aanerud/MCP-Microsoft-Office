/**
 * @fileoverview PowerPoint Controller - Handles API requests for PowerPoint presentation operations.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for PowerPoint endpoints
 */
const schemas = {
  fileId: Joi.object({
    fileId: Joi.string().required()
  }),

  createPresentation: Joi.object({
    fileName: Joi.string().required(),
    slides: Joi.array().optional(),
    folderId: Joi.string().optional()
  })
};

/**
 * Creates a PowerPoint controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.powerpointModule - Initialized PowerPoint module
 * @returns {object} Controller methods
 */
function createPowerPointController({ powerpointModule }) {
  if (!powerpointModule) {
    throw new Error('PowerPoint module is required for PowerPointController');
  }

  return {
    /**
     * POST /api/powerpoint/presentations
     */
    async createPresentation(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createPresentation.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await powerpointModule.handleIntent('createPresentation', value, { req });

        MonitoringService.trackMetric('powerpoint.createPresentation.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('powerpoint', 'Failed to create presentation', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'POWERPOINT_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/powerpoint/presentations/read
     */
    async readPresentation(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await powerpointModule.handleIntent('readPresentation', value, { req });

        MonitoringService.trackMetric('powerpoint.readPresentation.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('powerpoint', 'Failed to read presentation', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'POWERPOINT_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/powerpoint/presentations/pdf
     */
    async convertToPdf(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await powerpointModule.handleIntent('convertPresentationToPdf', value, { req });

        MonitoringService.trackMetric('powerpoint.convertToPdf.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('powerpoint', 'Failed to convert presentation to PDF', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'POWERPOINT_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/powerpoint/presentations/metadata
     */
    async getMetadata(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await powerpointModule.handleIntent('getPresentationMetadata', value, { req });

        MonitoringService.trackMetric('powerpoint.getMetadata.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('powerpoint', 'Failed to get presentation metadata', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'POWERPOINT_OPERATION_FAILED', error_description: err.message });
      }
    },

    // ========== Consolidated compound tool handler ==========

    /** POST /api/powerpoint/action */
    async powerpointPresentation(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await powerpointModule.handleIntent('powerpointPresentation', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'POWERPOINT_OPERATION_FAILED', error_description: err.message });
      }
    }
  };
}

module.exports = createPowerPointController;
