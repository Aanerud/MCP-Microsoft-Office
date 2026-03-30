/**
 * @fileoverview MCP PowerPoint Module - Handles PowerPoint presentation operations for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const POWERPOINT_CAPABILITIES = [
  'powerpointPresentation',
  // Legacy granular capabilities kept for backward compatibility (REST API)
  'createPresentation', 'readPresentation', 'convertPresentationToPdf',
  'getPresentationMetadata'
];

const PowerPointModule = {
  id: 'powerpoint',
  name: 'PowerPoint Presentation Operations',
  capabilities: POWERPOINT_CAPABILITIES,

  init(services) {
    if (!services) throw ErrorService.createError('powerpoint', 'Services object required', 'critical', {});
    if (!services.powerpointService) throw ErrorService.createError('powerpoint', 'powerpointService required', 'critical', {});
    this.services = services;
    MonitoringService.info('PowerPoint Module initialized', { serviceName: 'powerpoint-module', capabilities: POWERPOINT_CAPABILITIES.length }, 'powerpoint');
    return this;
  },

  async handleIntent(intent, entities = {}, context = {}) {
    const { powerpointService } = this.services;
    const req = context.req;
    const userId = context.userId || req?.user?.userId;
    const sessionId = context.sessionId || req?.session?.id;

    try {
      switch (intent) {
        case 'createPresentation':
          return { type: 'fileCreated', file: await powerpointService.createPresentation(entities.fileName, entities.slides, entities.folderId, req, userId, sessionId) };
        case 'readPresentation':
          return { type: 'presentationContent', ...(await powerpointService.readPresentation(entities.fileId, req, userId, sessionId)) };
        case 'convertPresentationToPdf':
          return { type: 'pdfContent', content: await powerpointService.convertPresentationToPdf(entities.fileId, req, userId, sessionId) };
        case 'getPresentationMetadata':
          return { type: 'presentationMetadata', metadata: await powerpointService.getPresentationMetadata(entities.fileId, req, userId, sessionId) };

        // ========== Consolidated compound tool ==========
        case 'powerpointPresentation': {
          const { action } = entities;
          switch (action) {
            case 'create':
              return { type: 'fileCreated', file: await powerpointService.createPresentation(entities.fileName, entities.slides, entities.folderId, req, userId, sessionId) };
            case 'read':
              return { type: 'presentationContent', ...(await powerpointService.readPresentation(entities.fileId, req, userId, sessionId)) };
            case 'metadata':
              return { type: 'presentationMetadata', metadata: await powerpointService.getPresentationMetadata(entities.fileId, req, userId, sessionId) };
            case 'pdf':
              return { type: 'pdfContent', content: await powerpointService.convertPresentationToPdf(entities.fileId, req, userId, sessionId) };
            default:
              throw ErrorService.createError('powerpoint', `Unknown powerpointPresentation action: ${action}`, 'error', { action });
          }
        }

        default:
          throw ErrorService.createError('powerpoint', `Unknown intent: ${intent}`, 'error', { intent });
      }
    } catch (error) {
      const mcpError = error.id ? error : ErrorService.createError('powerpoint', `PowerPoint operation failed: ${error.message}`, 'error', { intent, error: error.message, stack: error.stack });
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  }
};

module.exports = PowerPointModule;
