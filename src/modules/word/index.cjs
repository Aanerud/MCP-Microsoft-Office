/**
 * @fileoverview MCP Word Module - Handles Word document operations for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const WORD_CAPABILITIES = [
  'wordDocument',
  // Legacy granular capabilities kept for backward compatibility (REST API)
  'createWordDocument', 'readWordDocument', 'convertDocumentToPdf',
  'getWordDocumentMetadata', 'getWordDocumentAsHtml'
];

const WordModule = {
  id: 'word',
  name: 'Word Document Operations',
  capabilities: WORD_CAPABILITIES,

  init(services) {
    if (!services) throw ErrorService.createError('word', 'Services object required', 'critical', {});
    if (!services.wordService) throw ErrorService.createError('word', 'wordService required', 'critical', {});
    this.services = services;
    MonitoringService.info('Word Module initialized', { serviceName: 'word-module', capabilities: WORD_CAPABILITIES.length }, 'word');
    return this;
  },

  async handleIntent(intent, entities = {}, context = {}) {
    const { wordService } = this.services;
    const req = context.req;
    const userId = context.userId || req?.user?.userId;
    const sessionId = context.sessionId || req?.session?.id;

    try {
      switch (intent) {
        case 'createWordDocument':
          return { type: 'fileCreated', file: await wordService.createWordDocument(entities.fileName, entities.content, entities.folderId, req, userId, sessionId) };
        case 'readWordDocument':
          return { type: 'documentContent', ...(await wordService.readWordDocument(entities.fileId, req, userId, sessionId)) };
        case 'convertDocumentToPdf':
          return { type: 'pdfContent', content: await wordService.convertDocumentToPdf(entities.fileId, req, userId, sessionId) };
        case 'getWordDocumentMetadata':
          return { type: 'documentMetadata', metadata: await wordService.getWordDocumentMetadata(entities.fileId, req, userId, sessionId) };
        case 'getWordDocumentAsHtml':
          return { type: 'documentHtml', ...(await wordService.getWordDocumentAsHtml(entities.fileId, req, userId, sessionId)) };

        // ========== Consolidated compound tool ==========
        case 'wordDocument': {
          const { action } = entities;
          switch (action) {
            case 'create':
              return { type: 'fileCreated', file: await wordService.createWordDocument(entities.fileName, entities.content, entities.folderId, req, userId, sessionId) };
            case 'read':
              return { type: 'documentContent', ...(await wordService.readWordDocument(entities.fileId, req, userId, sessionId)) };
            case 'metadata':
              return { type: 'documentMetadata', metadata: await wordService.getWordDocumentMetadata(entities.fileId, req, userId, sessionId) };
            case 'html':
              return { type: 'documentHtml', ...(await wordService.getWordDocumentAsHtml(entities.fileId, req, userId, sessionId)) };
            case 'pdf':
              return { type: 'pdfContent', content: await wordService.convertDocumentToPdf(entities.fileId, req, userId, sessionId) };
            default:
              throw ErrorService.createError('word', `Unknown wordDocument action: ${action}`, 'error', { action });
          }
        }

        default:
          throw ErrorService.createError('word', `Unknown intent: ${intent}`, 'error', { intent });
      }
    } catch (error) {
      const mcpError = error.id ? error : ErrorService.createError('word', `Word operation failed: ${error.message}`, 'error', { intent, error: error.message, stack: error.stack });
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  }
};

module.exports = WordModule;
