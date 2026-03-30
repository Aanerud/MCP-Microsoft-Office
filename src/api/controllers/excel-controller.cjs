/**
 * @fileoverview Excel Controller - Handles API requests for Excel workbook operations.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for Excel endpoints
 */
const schemas = {
  fileId: Joi.object({
    fileId: Joi.string().required()
  }),

  createSession: Joi.object({
    fileId: Joi.string().required(),
    persistent: Joi.boolean().optional()
  }),

  worksheetAction: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required()
  }),

  addWorksheet: Joi.object({
    fileId: Joi.string().required(),
    name: Joi.string().optional()
  }),

  updateWorksheet: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    properties: Joi.object().required()
  }),

  rangeAction: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required()
  }),

  updateRange: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required(),
    values: Joi.array().items(Joi.array()).required()
  }),

  updateRangeFormat: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required(),
    format: Joi.object().required()
  }),

  sortRange: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required(),
    fields: Joi.array().items(Joi.object()).required()
  }),

  mergeRange: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required(),
    across: Joi.boolean().optional()
  }),

  createTable: Joi.object({
    fileId: Joi.string().required(),
    sheetIdOrName: Joi.string().required(),
    address: Joi.string().required(),
    hasHeaders: Joi.boolean().optional()
  }),

  tableAction: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required()
  }),

  updateTable: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    properties: Joi.object().required()
  }),

  addTableRow: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    values: Joi.array().required(),
    index: Joi.number().integer().optional()
  }),

  deleteTableRow: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    index: Joi.number().integer().required()
  }),

  addTableColumn: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    values: Joi.array().optional(),
    index: Joi.number().integer().optional()
  }),

  deleteTableColumn: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    columnIdOrName: Joi.string().required()
  }),

  sortTable: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    fields: Joi.array().items(Joi.object()).required()
  }),

  filterTable: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    columnId: Joi.number().integer().required(),
    criteria: Joi.object().required()
  }),

  clearTableFilter: Joi.object({
    fileId: Joi.string().required(),
    tableIdOrName: Joi.string().required(),
    columnId: Joi.number().integer().required()
  }),

  callWorkbookFunction: Joi.object({
    fileId: Joi.string().required(),
    functionName: Joi.string().required(),
    args: Joi.object().required()
  }),

  calculateWorkbook: Joi.object({
    fileId: Joi.string().required(),
    calculationType: Joi.string().valid('recalculate', 'full', 'fullRebuild').optional()
  })
};

/**
 * Creates an Excel controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.excelModule - Initialized Excel module
 * @returns {object} Controller methods
 */
function createExcelController({ excelModule }) {
  if (!excelModule) {
    throw new Error('Excel module is required for ExcelController');
  }

  return {
    /**
     * POST /api/excel/session
     */
    async createSession(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createSession.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('createWorkbookSession', value, { req });

        MonitoringService.trackMetric('excel.createSession.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to create workbook session', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/session
     */
    async closeSession(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId || req.body.fileId;
        if (!fileId) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId is required' });
        }

        const result = await excelModule.handleIntent('closeWorkbookSession', { fileId }, { req });

        MonitoringService.trackMetric('excel.closeSession.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to close workbook session', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/worksheets
     */
    async listWorksheets(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.fileId.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('listWorksheets', value, { req });

        MonitoringService.trackMetric('excel.listWorksheets.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to list worksheets', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/worksheets
     */
    async addWorksheet(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.addWorksheet.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('addWorksheet', value, { req });

        MonitoringService.trackMetric('excel.addWorksheet.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to add worksheet', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/worksheets/:sheetIdOrName
     */
    async getWorksheet(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId;
        const sheetIdOrName = req.query.sheetIdOrName || req.body.sheetIdOrName;
        if (!fileId || !sheetIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and sheetIdOrName are required' });
        }

        const result = await excelModule.handleIntent('getWorksheet', { fileId, sheetIdOrName }, { req });

        MonitoringService.trackMetric('excel.getWorksheet.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to get worksheet', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * PATCH /api/excel/worksheets/:sheetIdOrName
     */
    async updateWorksheet(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, fileId: req.body.fileId || req.query.fileId, sheetIdOrName: req.query.sheetIdOrName || req.body.sheetIdOrName };
        const { error, value } = schemas.updateWorksheet.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('updateWorksheet', value, { req });

        MonitoringService.trackMetric('excel.updateWorksheet.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to update worksheet', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/worksheets/:sheetIdOrName
     */
    async deleteWorksheet(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId || req.body.fileId;
        const sheetIdOrName = req.query.sheetIdOrName || req.body.sheetIdOrName;
        if (!fileId || !sheetIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and sheetIdOrName are required' });
        }

        const result = await excelModule.handleIntent('deleteWorksheet', { fileId, sheetIdOrName }, { req });

        MonitoringService.trackMetric('excel.deleteWorksheet.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to delete worksheet', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/range
     */
    async getRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.rangeAction.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('getRange', value, { req });

        MonitoringService.trackMetric('excel.getRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to get range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * PATCH /api/excel/range
     */
    async updateRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.updateRange.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('updateRange', value, { req });

        MonitoringService.trackMetric('excel.updateRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to update range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/range/format
     */
    async getRangeFormat(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.rangeAction.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('getRangeFormat', value, { req });

        MonitoringService.trackMetric('excel.getRangeFormat.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to get range format', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * PATCH /api/excel/range/format
     */
    async updateRangeFormat(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.updateRangeFormat.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('updateRangeFormat', value, { req });

        MonitoringService.trackMetric('excel.updateRangeFormat.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to update range format', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/range/sort
     */
    async sortRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.sortRange.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('sortRange', value, { req });

        MonitoringService.trackMetric('excel.sortRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to sort range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/range/merge
     */
    async mergeRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.mergeRange.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('mergeRange', value, { req });

        MonitoringService.trackMetric('excel.mergeRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to merge range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/range/unmerge
     */
    async unmergeRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.rangeAction.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('unmergeRange', value, { req });

        MonitoringService.trackMetric('excel.unmergeRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to unmerge range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/tables
     */
    async listTables(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId;
        const sheetIdOrName = req.query.sheetIdOrName;
        if (!fileId) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId is required' });
        }

        const result = await excelModule.handleIntent('listTables', { fileId, sheetIdOrName }, { req });

        MonitoringService.trackMetric('excel.listTables.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to list tables', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables
     */
    async createTable(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createTable.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('createTable', value, { req });

        MonitoringService.trackMetric('excel.createTable.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to create table', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * PATCH /api/excel/tables/:tableIdOrName
     */
    async updateTable(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, fileId: req.body.fileId || req.query.fileId, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName };
        const { error, value } = schemas.updateTable.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('updateTable', value, { req });

        MonitoringService.trackMetric('excel.updateTable.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to update table', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/tables/:tableIdOrName
     */
    async deleteTable(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId || req.body.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        if (!fileId || !tableIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and tableIdOrName are required' });
        }

        const result = await excelModule.handleIntent('deleteTable', { fileId, tableIdOrName }, { req });

        MonitoringService.trackMetric('excel.deleteTable.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to delete table', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/tables/:tableIdOrName/rows
     */
    async listTableRows(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        if (!fileId || !tableIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and tableIdOrName are required' });
        }

        const result = await excelModule.handleIntent('listTableRows', { fileId, tableIdOrName }, { req });

        MonitoringService.trackMetric('excel.listTableRows.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to list table rows', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables/:tableIdOrName/rows
     */
    async addTableRow(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName };
        const { error, value } = schemas.addTableRow.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('addTableRow', value, { req });

        MonitoringService.trackMetric('excel.addTableRow.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to add table row', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/tables/:tableIdOrName/rows/:index
     */
    async deleteTableRow(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId || req.body.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        const index = parseInt(req.query.index || req.body.index, 10);
        if (!fileId || !tableIdOrName || isNaN(index)) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId, tableIdOrName, and index are required' });
        }

        const result = await excelModule.handleIntent('deleteTableRow', { fileId, tableIdOrName, index }, { req });

        MonitoringService.trackMetric('excel.deleteTableRow.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to delete table row', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * GET /api/excel/tables/:tableIdOrName/columns
     */
    async listTableColumns(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        if (!fileId || !tableIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and tableIdOrName are required' });
        }

        const result = await excelModule.handleIntent('listTableColumns', { fileId, tableIdOrName }, { req });

        MonitoringService.trackMetric('excel.listTableColumns.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to list table columns', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables/:tableIdOrName/columns
     */
    async addTableColumn(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName };
        const { error, value } = schemas.addTableColumn.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('addTableColumn', value, { req });

        MonitoringService.trackMetric('excel.addTableColumn.duration', Date.now() - startTime, { success: true, userId });
        res.status(201).json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to add table column', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/tables/:tableIdOrName/columns/:columnIdOrName
     */
    async deleteTableColumn(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.query.fileId || req.body.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        const columnIdOrName = req.query.columnIdOrName || req.body.columnIdOrName;
        if (!fileId || !tableIdOrName || !columnIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId, tableIdOrName, and columnIdOrName are required' });
        }

        const result = await excelModule.handleIntent('deleteTableColumn', { fileId, tableIdOrName, columnIdOrName }, { req });

        MonitoringService.trackMetric('excel.deleteTableColumn.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to delete table column', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables/:tableIdOrName/sort
     */
    async sortTable(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName };
        const { error, value } = schemas.sortTable.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('sortTable', value, { req });

        MonitoringService.trackMetric('excel.sortTable.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to sort table', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables/:tableIdOrName/filter
     */
    async filterTable(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { ...req.body, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName };
        const { error, value } = schemas.filterTable.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('filterTable', value, { req });

        MonitoringService.trackMetric('excel.filterTable.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to filter table', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * DELETE /api/excel/tables/:tableIdOrName/filter
     */
    async clearTableFilter(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const entities = { fileId: req.query.fileId || req.body.fileId, tableIdOrName: req.query.tableIdOrName || req.body.tableIdOrName, columnId: parseInt(req.query.columnId || req.body.columnId, 10) };
        const { error, value } = schemas.clearTableFilter.validate(entities);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('clearTableFilter', value, { req });

        MonitoringService.trackMetric('excel.clearTableFilter.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to clear table filter', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/tables/:tableIdOrName/convertToRange
     */
    async convertTableToRange(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const fileId = req.body.fileId || req.query.fileId;
        const tableIdOrName = req.query.tableIdOrName || req.body.tableIdOrName;
        if (!fileId || !tableIdOrName) {
          return res.status(400).json({ error: 'Invalid request', details: 'fileId and tableIdOrName are required' });
        }

        const result = await excelModule.handleIntent('convertTableToRange', { fileId, tableIdOrName }, { req });

        MonitoringService.trackMetric('excel.convertTableToRange.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to convert table to range', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/functions
     */
    async callWorkbookFunction(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.callWorkbookFunction.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('callWorkbookFunction', value, { req });

        MonitoringService.trackMetric('excel.callWorkbookFunction.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to call workbook function', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /**
     * POST /api/excel/calculate
     */
    async calculateWorkbook(req, res) {
      const { userId = null } = req.user || {};
      const startTime = Date.now();

      try {
        const { error, value } = schemas.calculateWorkbook.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const result = await excelModule.handleIntent('calculateWorkbook', value, { req });

        MonitoringService.trackMetric('excel.calculateWorkbook.duration', Date.now() - startTime, { success: true, userId });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(ErrorService.createError('excel', 'Failed to calculate workbook', 'error', {
          error: err.message, userId, timestamp: new Date().toISOString()
        }));
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    // ========== Consolidated compound tool handlers ==========

    /** POST /api/excel/session/action */
    async excelSession(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await excelModule.handleIntent('excelSession', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /** POST /api/excel/worksheet/action */
    async excelWorksheet(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await excelModule.handleIntent('excelWorksheet', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /** POST /api/excel/range/action */
    async excelRange(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await excelModule.handleIntent('excelRange', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /** POST /api/excel/table/action */
    async excelTable(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await excelModule.handleIntent('excelTable', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    },

    /** POST /api/excel/function/action */
    async excelFunction(req, res) {
      const { userId = null } = req.user || {};
      try {
        const result = await excelModule.handleIntent('excelFunction', req.body, { req });
        res.json(result);
      } catch (err) {
        MonitoringService.logError(err);
        res.status(err.statusCode || 500).json({ error: 'EXCEL_OPERATION_FAILED', error_description: err.message });
      }
    }
  };
}

module.exports = createExcelController;
