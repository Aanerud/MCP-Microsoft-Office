/**
 * @fileoverview MCP Search Module - Unified search across Microsoft 365.
 * Provides cross-entity search for emails, calendar events, files, and people.
 * Uses Microsoft Graph Search API (/search/query) with KQL support.
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const SEARCH_CAPABILITIES = [
  'search'
];

// Log module initialization
MonitoringService.info('Search Module initialized', {
  serviceName: 'search-module',
  capabilities: SEARCH_CAPABILITIES.length,
  timestamp: new Date().toISOString()
}, 'search');

const SearchModule = {
  /**
   * Module ID
   */
  id: 'search',

  /**
   * Module name
   */
  name: 'Microsoft 365 Unified Search',

  /**
   * Module capabilities
   */
  capabilities: SEARCH_CAPABILITIES,

  /**
   * Service dependencies (injected during init)
   */
  services: null,

  /**
   * Helper method to redact sensitive data from objects before logging
   * @param {object} data - The data object to redact
   * @returns {object} Redacted copy of the data
   * @private
   */
  redactSensitiveData(data) {
    if (!data || typeof data !== 'object') {
      return data;
    }

    const result = Array.isArray(data) ? [...data] : { ...data };

    const sensitiveFields = [
      'body', 'content', 'email', 'emailAddress', 'address',
      'from', 'to', 'subject', 'bodyPreview'
    ];

    for (const key in result) {
      if (Object.prototype.hasOwnProperty.call(result, key)) {
        if (sensitiveFields.includes(key.toLowerCase())) {
          if (typeof result[key] === 'string') {
            result[key] = result[key].substring(0, 20) + '...';
          } else if (Array.isArray(result[key])) {
            result[key] = `[${result[key].length} items]`;
          } else if (typeof result[key] === 'object' && result[key] !== null) {
            result[key] = '{...}';
          }
        } else if (typeof result[key] === 'object' && result[key] !== null) {
          result[key] = this.redactSensitiveData(result[key]);
        }
      }
    }

    return result;
  },

  /**
   * Initialize the search module with dependencies
   * @param {object} services - Service dependencies
   */
  init(services) {
    this.services = services;

    MonitoringService.info('Search Module services initialized', {
      hasSearchService: !!services?.searchService,
      timestamp: new Date().toISOString()
    }, 'search');

    return this;
  },

  /**
   * Unified search across Microsoft 365
   * @param {object} options - Search options
   * @param {string} options.query - Search query (KQL supported)
   * @param {Array<string>} [options.entityTypes] - Entity types to search
   * @param {number} [options.limit] - Max results
   * @param {number} [options.from] - Pagination offset
   * @param {object} req - Express request
   * @param {string} userId - User ID
   * @param {string} sessionId - Session ID
   * @returns {Promise<object>} Search results
   */
  async search(options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const { searchService } = this.services || {};

    // Development debug logging
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Search module: search operation started', {
        query: options.query?.substring(0, 50),
        entityTypes: options.entityTypes,
        limit: options.limit,
        timestamp: new Date().toISOString()
      }, 'search');
    }

    if (!searchService || typeof searchService.search !== 'function') {
      const mcpError = ErrorService.createError(
        'search',
        'SearchService not available',
        'error',
        { method: 'search', moduleId: 'search' }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }

    try {
      const results = await searchService.search(
        {
          query: options.query,
          entityTypes: options.entityTypes,
          size: options.limit || 25,
          from: options.from || 0
        },
        req,
        userId,
        sessionId
      );

      const executionTime = Date.now() - startTime;

      // User activity logging
      if (userId) {
        MonitoringService.info('Search completed via module', {
          query: options.query?.substring(0, 30),
          resultCount: results.results?.length || 0,
          executionTimeMs: executionTime,
          timestamp: new Date().toISOString()
        }, 'search', null, userId);
      }

      return results;

    } catch (error) {
      const executionTime = Date.now() - startTime;

      MonitoringService.error('Search module: search failed', {
        query: options.query?.substring(0, 30),
        error: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'search', null, userId);

      throw error;
    }
  },

  /**
   * Handle search intents from MCP
   * @param {string} intent - The intent to handle
   * @param {object} params - Intent parameters
   * @param {object} context - Execution context
   * @returns {Promise<object>} Intent result
   */
  async handleIntent(intent, params = {}, context = {}) {
    const { req, userId, sessionId } = context;

    MonitoringService.debug('Search module handling intent', {
      intent,
      hasParams: Object.keys(params).length > 0,
      timestamp: new Date().toISOString()
    }, 'search');

    switch (intent) {
      case 'search':
        return this.search(params, req, userId, sessionId);

      case 'searchMessages':
        return this.search({ ...params, entityTypes: ['message'] }, req, userId, sessionId);

      case 'searchEvents':
        return this.search({ ...params, entityTypes: ['event'] }, req, userId, sessionId);

      case 'searchFiles':
        return this.search({ ...params, entityTypes: ['driveItem'] }, req, userId, sessionId);

      case 'searchPeople':
        return this.search({ ...params, entityTypes: ['person'] }, req, userId, sessionId);

      default:
        const error = ErrorService.createError(
          'search',
          `Unknown search intent: ${intent}`,
          'warning',
          { intent, availableIntents: SEARCH_CAPABILITIES }
        );
        MonitoringService.logError(error);
        throw error;
    }
  }
};

module.exports = SearchModule;
