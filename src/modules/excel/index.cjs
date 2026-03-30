/**
 * @fileoverview MCP Excel Module - Handles Excel workbook operations for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const EXCEL_CAPABILITIES = [
  'excelSession', 'excelWorksheet', 'excelRange', 'excelTable', 'excelFunction',
  // Legacy granular capabilities kept for backward compatibility (REST API)
  'createWorkbookSession', 'closeWorkbookSession',
  'listWorksheets', 'addWorksheet', 'getWorksheet', 'updateWorksheet', 'deleteWorksheet',
  'getRange', 'updateRange', 'getRangeFormat', 'updateRangeFormat', 'sortRange', 'mergeRange', 'unmergeRange',
  'listTables', 'createTable', 'updateTable', 'deleteTable',
  'listTableRows', 'addTableRow', 'deleteTableRow',
  'listTableColumns', 'addTableColumn', 'deleteTableColumn',
  'sortTable', 'filterTable', 'clearTableFilter', 'convertTableToRange',
  'callWorkbookFunction', 'calculateWorkbook'
];

const ExcelModule = {
  id: 'excel',
  name: 'Excel Workbook Operations',
  capabilities: EXCEL_CAPABILITIES,

  init(services) {
    if (!services) throw ErrorService.createError('excel', 'Services object required', 'critical', {});
    if (!services.excelService) throw ErrorService.createError('excel', 'excelService required', 'critical', {});
    this.services = services;
    MonitoringService.info('Excel Module initialized', { serviceName: 'excel-module', capabilities: EXCEL_CAPABILITIES.length }, 'excel');
    return this;
  },

  async handleIntent(intent, entities = {}, context = {}) {
    const { excelService } = this.services;
    const req = context.req;
    const userId = context.userId || req?.user?.userId;
    const sessionId = context.sessionId || req?.session?.id;
    const { fileId } = entities;

    try {
      switch (intent) {
        // Session
        case 'createWorkbookSession':
          return { type: 'workbookSession', ...(await excelService.createWorkbookSession(fileId, entities.persistent !== false, req, userId, sessionId)) };
        case 'closeWorkbookSession':
          await excelService.closeWorkbookSession(fileId, req, userId, sessionId);
          return { type: 'success', message: 'Session closed' };

        // Worksheets
        case 'listWorksheets':
          return { type: 'worksheetList', worksheets: await excelService.listWorksheets(fileId, req, userId, sessionId) };
        case 'addWorksheet':
          return { type: 'worksheet', worksheet: await excelService.addWorksheet(fileId, entities.name, req, userId, sessionId) };
        case 'getWorksheet':
          return { type: 'worksheet', worksheet: await excelService.getWorksheet(fileId, entities.sheetIdOrName, req, userId, sessionId) };
        case 'updateWorksheet':
          return { type: 'worksheet', worksheet: await excelService.updateWorksheet(fileId, entities.sheetIdOrName, entities.properties, req, userId, sessionId) };
        case 'deleteWorksheet':
          await excelService.deleteWorksheet(fileId, entities.sheetIdOrName, req, userId, sessionId);
          return { type: 'success', message: 'Worksheet deleted' };

        // Ranges
        case 'getRange':
          return { type: 'range', range: await excelService.getRange(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId) };
        case 'updateRange':
          return { type: 'range', range: await excelService.updateRange(fileId, entities.sheetIdOrName, entities.address, entities.values, req, userId, sessionId) };
        case 'getRangeFormat':
          return { type: 'rangeFormat', format: await excelService.getRangeFormat(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId) };
        case 'updateRangeFormat':
          return { type: 'rangeFormat', format: await excelService.updateRangeFormat(fileId, entities.sheetIdOrName, entities.address, entities.format, req, userId, sessionId) };
        case 'sortRange':
          await excelService.sortRange(fileId, entities.sheetIdOrName, entities.address, entities.fields, req, userId, sessionId);
          return { type: 'success', message: 'Range sorted' };
        case 'mergeRange':
          await excelService.mergeRange(fileId, entities.sheetIdOrName, entities.address, entities.across, req, userId, sessionId);
          return { type: 'success', message: 'Range merged' };
        case 'unmergeRange':
          await excelService.unmergeRange(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId);
          return { type: 'success', message: 'Range unmerged' };

        // Tables
        case 'listTables':
          return { type: 'tableList', tables: await excelService.listTables(fileId, entities.sheetIdOrName, req, userId, sessionId) };
        case 'createTable':
          return { type: 'table', table: await excelService.createTable(fileId, entities.sheetIdOrName, entities.address, entities.hasHeaders, req, userId, sessionId) };
        case 'updateTable':
          return { type: 'table', table: await excelService.updateTable(fileId, entities.tableIdOrName, entities.properties, req, userId, sessionId) };
        case 'deleteTable':
          await excelService.deleteTable(fileId, entities.tableIdOrName, req, userId, sessionId);
          return { type: 'success', message: 'Table deleted' };
        case 'listTableRows':
          return { type: 'tableRows', rows: await excelService.listTableRows(fileId, entities.tableIdOrName, req, userId, sessionId) };
        case 'addTableRow':
          return { type: 'tableRow', row: await excelService.addTableRow(fileId, entities.tableIdOrName, entities.values, entities.index, req, userId, sessionId) };
        case 'deleteTableRow':
          await excelService.deleteTableRow(fileId, entities.tableIdOrName, entities.index, req, userId, sessionId);
          return { type: 'success', message: 'Table row deleted' };
        case 'listTableColumns':
          return { type: 'tableColumns', columns: await excelService.listTableColumns(fileId, entities.tableIdOrName, req, userId, sessionId) };
        case 'addTableColumn':
          return { type: 'tableColumn', column: await excelService.addTableColumn(fileId, entities.tableIdOrName, entities.values, entities.index, req, userId, sessionId) };
        case 'deleteTableColumn':
          await excelService.deleteTableColumn(fileId, entities.tableIdOrName, entities.columnIdOrName, req, userId, sessionId);
          return { type: 'success', message: 'Table column deleted' };
        case 'sortTable':
          await excelService.sortTable(fileId, entities.tableIdOrName, entities.fields, req, userId, sessionId);
          return { type: 'success', message: 'Table sorted' };
        case 'filterTable':
          await excelService.filterTable(fileId, entities.tableIdOrName, entities.columnId, entities.criteria, req, userId, sessionId);
          return { type: 'success', message: 'Table filter applied' };
        case 'clearTableFilter':
          await excelService.clearTableFilter(fileId, entities.tableIdOrName, entities.columnId, req, userId, sessionId);
          return { type: 'success', message: 'Table filter cleared' };
        case 'convertTableToRange':
          return { type: 'range', range: await excelService.convertTableToRange(fileId, entities.tableIdOrName, req, userId, sessionId) };

        // Functions
        case 'callWorkbookFunction':
          return { type: 'functionResult', result: await excelService.callWorkbookFunction(fileId, entities.functionName, entities.args, req, userId, sessionId) };
        case 'calculateWorkbook':
          await excelService.calculateWorkbook(fileId, entities.calculationType, req, userId, sessionId);
          return { type: 'success', message: 'Workbook calculated' };

        // ========== Consolidated compound tools ==========
        case 'excelSession': {
          const ALLOWED = ['create', 'close'];
          const { action } = entities;
          if (!ALLOWED.includes(action)) throw ErrorService.createError('excel', `Invalid action: ${String(action).substring(0, 30)}`, 'error', { action: String(action).substring(0, 30) });
          switch (action) {
            case 'create':
              return { type: 'workbookSession', ...(await excelService.createWorkbookSession(fileId, entities.persistent !== false, req, userId, sessionId)) };
            case 'close':
              await excelService.closeWorkbookSession(fileId, req, userId, sessionId);
              return { type: 'success', message: 'Session closed' };
            default:
              throw ErrorService.createError('excel', `Unknown excelSession action: ${action}`, 'error', { action });
          }
        }
        case 'excelWorksheet': {
          const ALLOWED = ['list', 'add', 'get', 'update', 'delete'];
          const { action } = entities;
          if (!ALLOWED.includes(action)) throw ErrorService.createError('excel', `Invalid action: ${String(action).substring(0, 30)}`, 'error', { action: String(action).substring(0, 30) });
          switch (action) {
            case 'list':
              return { type: 'worksheetList', worksheets: await excelService.listWorksheets(fileId, req, userId, sessionId) };
            case 'add':
              return { type: 'worksheet', worksheet: await excelService.addWorksheet(fileId, entities.name, req, userId, sessionId) };
            case 'get':
              return { type: 'worksheet', worksheet: await excelService.getWorksheet(fileId, entities.sheetIdOrName, req, userId, sessionId) };
            case 'update':
              return { type: 'worksheet', worksheet: await excelService.updateWorksheet(fileId, entities.sheetIdOrName, entities.properties, req, userId, sessionId) };
            case 'delete':
              await excelService.deleteWorksheet(fileId, entities.sheetIdOrName, req, userId, sessionId);
              return { type: 'success', message: 'Worksheet deleted' };
            default:
              throw ErrorService.createError('excel', `Unknown excelWorksheet action: ${action}`, 'error', { action });
          }
        }
        case 'excelRange': {
          const ALLOWED = ['get', 'update', 'getFormat', 'updateFormat', 'sort', 'merge', 'unmerge'];
          const { action } = entities;
          if (!ALLOWED.includes(action)) throw ErrorService.createError('excel', `Invalid action: ${String(action).substring(0, 30)}`, 'error', { action: String(action).substring(0, 30) });
          switch (action) {
            case 'get':
              return { type: 'range', range: await excelService.getRange(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId) };
            case 'update':
              return { type: 'range', range: await excelService.updateRange(fileId, entities.sheetIdOrName, entities.address, entities.values, req, userId, sessionId) };
            case 'getFormat':
              return { type: 'rangeFormat', format: await excelService.getRangeFormat(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId) };
            case 'updateFormat':
              return { type: 'rangeFormat', format: await excelService.updateRangeFormat(fileId, entities.sheetIdOrName, entities.address, entities.format, req, userId, sessionId) };
            case 'sort':
              await excelService.sortRange(fileId, entities.sheetIdOrName, entities.address, entities.fields, req, userId, sessionId);
              return { type: 'success', message: 'Range sorted' };
            case 'merge':
              await excelService.mergeRange(fileId, entities.sheetIdOrName, entities.address, entities.across, req, userId, sessionId);
              return { type: 'success', message: 'Range merged' };
            case 'unmerge':
              await excelService.unmergeRange(fileId, entities.sheetIdOrName, entities.address, req, userId, sessionId);
              return { type: 'success', message: 'Range unmerged' };
            default:
              throw ErrorService.createError('excel', `Unknown excelRange action: ${action}`, 'error', { action });
          }
        }
        case 'excelTable': {
          const ALLOWED = ['list', 'create', 'update', 'delete', 'listRows', 'addRow', 'deleteRow', 'listColumns', 'addColumn', 'deleteColumn', 'sort', 'filter', 'clearFilter', 'convertToRange'];
          const { action } = entities;
          if (!ALLOWED.includes(action)) throw ErrorService.createError('excel', `Invalid action: ${String(action).substring(0, 30)}`, 'error', { action: String(action).substring(0, 30) });
          switch (action) {
            case 'list':
              return { type: 'tableList', tables: await excelService.listTables(fileId, entities.sheetIdOrName, req, userId, sessionId) };
            case 'create':
              return { type: 'table', table: await excelService.createTable(fileId, entities.sheetIdOrName, entities.address, entities.hasHeaders, req, userId, sessionId) };
            case 'update':
              return { type: 'table', table: await excelService.updateTable(fileId, entities.tableIdOrName, entities.properties, req, userId, sessionId) };
            case 'delete':
              await excelService.deleteTable(fileId, entities.tableIdOrName, req, userId, sessionId);
              return { type: 'success', message: 'Table deleted' };
            case 'listRows':
              return { type: 'tableRows', rows: await excelService.listTableRows(fileId, entities.tableIdOrName, req, userId, sessionId) };
            case 'addRow':
              return { type: 'tableRow', row: await excelService.addTableRow(fileId, entities.tableIdOrName, entities.values, entities.index, req, userId, sessionId) };
            case 'deleteRow':
              await excelService.deleteTableRow(fileId, entities.tableIdOrName, entities.index, req, userId, sessionId);
              return { type: 'success', message: 'Table row deleted' };
            case 'listColumns':
              return { type: 'tableColumns', columns: await excelService.listTableColumns(fileId, entities.tableIdOrName, req, userId, sessionId) };
            case 'addColumn':
              return { type: 'tableColumn', column: await excelService.addTableColumn(fileId, entities.tableIdOrName, entities.values, entities.index, req, userId, sessionId) };
            case 'deleteColumn':
              await excelService.deleteTableColumn(fileId, entities.tableIdOrName, entities.columnIdOrName, req, userId, sessionId);
              return { type: 'success', message: 'Table column deleted' };
            case 'sort':
              await excelService.sortTable(fileId, entities.tableIdOrName, entities.fields, req, userId, sessionId);
              return { type: 'success', message: 'Table sorted' };
            case 'filter':
              await excelService.filterTable(fileId, entities.tableIdOrName, entities.columnId, entities.criteria, req, userId, sessionId);
              return { type: 'success', message: 'Table filter applied' };
            case 'clearFilter':
              await excelService.clearTableFilter(fileId, entities.tableIdOrName, entities.columnId, req, userId, sessionId);
              return { type: 'success', message: 'Table filter cleared' };
            case 'convertToRange':
              return { type: 'range', range: await excelService.convertTableToRange(fileId, entities.tableIdOrName, req, userId, sessionId) };
            default:
              throw ErrorService.createError('excel', `Unknown excelTable action: ${action}`, 'error', { action });
          }
        }
        case 'excelFunction': {
          const ALLOWED = ['call', 'calculate'];
          const { action } = entities;
          if (!ALLOWED.includes(action)) throw ErrorService.createError('excel', `Invalid action: ${String(action).substring(0, 30)}`, 'error', { action: String(action).substring(0, 30) });
          switch (action) {
            case 'call':
              return { type: 'functionResult', result: await excelService.callWorkbookFunction(fileId, entities.functionName, entities.args, req, userId, sessionId) };
            case 'calculate':
              await excelService.calculateWorkbook(fileId, entities.calculationType, req, userId, sessionId);
              return { type: 'success', message: 'Workbook calculated' };
            default:
              throw ErrorService.createError('excel', `Unknown excelFunction action: ${action}`, 'error', { action });
          }
        }

        default:
          throw ErrorService.createError('excel', `Unknown intent: ${intent}`, 'error', { intent });
      }
    } catch (error) {
      const mcpError = error.id ? error : ErrorService.createError('excel', `Excel operation failed: ${error.message}`, 'error', { intent, error: error.message, stack: error.stack });
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  }
};

module.exports = ExcelModule;
