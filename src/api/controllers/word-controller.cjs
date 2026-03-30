/**
 * @fileoverview Word Controller - Handles API requests for Word document operations.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for Word endpoints
 */
const schemas = {
  fileId: Joi.object({
    fileId: Joi.string().required()
  }),

  createDocument: Joi.object({
    fileName: Joi.string().required(),
    content: Joi.object().required(),
    folderId: Joi.string().optional()
  })
};

/**
 * Creates a Word controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.wordModule - Initialized Word module
 * @returns {object} Controller methods
 */
function createWordController({ wordModule }) {
  if (!wordModule) {
    throw new Error('Word module is required for WordController');
  }

  return {
    /**
     * POST /api/word/documents
     */
    async createDocument(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createDocument.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await wordModule.handleIntent('createWordDocument', value, { req });

        MonitoringService.trackMetric('word.createDocument.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('word', 'Failed to create Word document', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/word/documents/read
     */
    async readDocument(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await wordModule.handleIntent('readWordDocument', value, { req });

        MonitoringService.trackMetric('word.readDocument.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('word', 'Failed to read Word document', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/word/documents/pdf
     */
    async convertToPdf(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await wordModule.handleIntent('convertDocumentToPdf', value, { req });

        MonitoringService.trackMetric('word.convertToPdf.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('word', 'Failed to convert document to PDF', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/word/documents/metadata
     */
    async getMetadata(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await wordModule.handleIntent('getWordDocumentMetadata', value, { req });

        MonitoringService.trackMetric('word.getMetadata.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('word', 'Failed to get document metadata', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/word/documents/html
     */
    async getAsHtml(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await wordModule.handleIntent('getWordDocumentAsHtml', value, { req });

        MonitoringService.trackMetric('word.getAsHtml.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('word', 'Failed to get document as HTML', 'error', {
          fileId: req.query.fileId, error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    },

    // ========== Consolidated compound tool handler ==========

    /** POST /api/word/action */
    async wordDocument(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await wordModule.handleIntent('wordDocument', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'WORD_OPERATION_FAILED', error_description: err.message });
      }
    }
  };
}

module.exports = createWordController;
