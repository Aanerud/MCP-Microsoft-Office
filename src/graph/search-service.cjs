/**
 * @fileoverview SearchService - Microsoft Graph Search API operations.
 * Provides unified search across emails, calendar events, files, and people.
 * Uses the /search/query endpoint for cross-entity searching with KQL support.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

// Log service initialization
MonitoringService.info('Graph Search Service initialized', {
  serviceName: 'graph-search-service',
  supportedEntityTypes: ['message', 'event', 'driveItem', 'person'],
  timestamp: new Date().toISOString()
}, 'graph');

/**
 * Valid entity types for Microsoft Graph Search API
 */
const VALID_ENTITY_TYPES = ['message', 'event', 'driveItem', 'person', 'chatMessage', 'site', 'list', 'listItem'];

/**
 * Entity type combinations that can be searched together
 * Microsoft Graph has restrictions on which types can be combined
 */
const ENTITY_TYPE_GROUPS = {
  // These can be searched together
  sharepoint: ['driveItem', 'site', 'list', 'listItem'],
  // These must be searched separately
  separate: ['message', 'event', 'chatMessage', 'person']
};

/**
 * Normalizes a search hit to a consistent format
 * @param {object} hit - Search hit from Graph API
 * @param {string} entityType - The entity type of the hit
 * @returns {object} Normalized search result
 */
function normalizeSearchHit(hit, entityType) {
  const resource = hit.resource || {};

  const base = {
    id: resource.id || hit.hitId,
    entityType,
    rank: hit.rank,
    summary: hit.summary || null
  };

  switch (entityType) {
    case 'message':
      return {
        ...base,
        subject: resource.subject,
        from: resource.from?.emailAddress ? {
          name: resource.from.emailAddress.name,
          email: resource.from.emailAddress.address
        } : null,
        receivedDateTime: resource.receivedDateTime,
        bodyPreview: resource.bodyPreview?.substring(0, 200),
        hasAttachments: resource.hasAttachments,
        importance: resource.importance,
        webLink: resource.webLink
      };

    case 'event':
      return {
        ...base,
        subject: resource.subject,
        start: resource.start,
        end: resource.end,
        location: resource.location?.displayName,
        organizer: resource.organizer?.emailAddress ? {
          name: resource.organizer.emailAddress.name,
          email: resource.organizer.emailAddress.address
        } : null,
        isAllDay: resource.isAllDay,
        webLink: resource.webLink
      };

    case 'driveItem':
      return {
        ...base,
        name: resource.name,
        webUrl: resource.webUrl,
        size: resource.size,
        createdDateTime: resource.createdDateTime,
        lastModifiedDateTime: resource.lastModifiedDateTime,
        createdBy: resource.createdBy?.user?.displayName,
        lastModifiedBy: resource.lastModifiedBy?.user?.displayName,
        mimeType: resource.file?.mimeType,
        parentPath: resource.parentReference?.path
      };

    case 'person':
      return {
        ...base,
        displayName: resource.displayName,
        givenName: resource.givenName,
        surname: resource.surname,
        emailAddresses: resource.emailAddresses || resource.scoredEmailAddresses?.map(e => e.address) || [],
        jobTitle: resource.jobTitle,
        department: resource.department,
        officeLocation: resource.officeLocation,
        companyName: resource.companyName
      };

    default:
      return {
        ...base,
        ...resource
      };
  }
}

/**
 * Performs a unified search across Microsoft 365 content
 * @param {object} options - Search options
 * @param {string} options.query - Search query string (KQL supported)
 * @param {Array<string>} [options.entityTypes] - Entity types to search
 * @param {number} [options.from=0] - Pagination offset
 * @param {number} [options.size=25] - Results per page (max 25)
 * @param {Array<string>} [options.fields] - Specific fields to return
 * @param {object} req - Express request object
 * @param {string} userId - User ID for context
 * @param {string} sessionId - Session ID for context
 * @returns {Promise<object>} Search results with normalized hits
 */
async function search(options = {}, req, userId, sessionId) {
  const startTime = Date.now();

  // Extract user context
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;

  // Validate and set defaults
  const {
    query,
    entityTypes = ['message', 'event', 'driveItem', 'person'],
    from = 0,
    size = 25,
    fields
  } = options;

  if (!query || typeof query !== 'string' || query.trim().length === 0) {
    const error = ErrorService.createError(
      'search',
      'Search query is required and must be a non-empty string',
      'warning',
      { providedQuery: query }
    );
    MonitoringService.logError(error);
    throw error;
  }

  // Validate entity types
  const validatedTypes = entityTypes.filter(type => VALID_ENTITY_TYPES.includes(type));
  if (validatedTypes.length === 0) {
    const error = ErrorService.createError(
      'search',
      `Invalid entity types provided. Valid types: ${VALID_ENTITY_TYPES.join(', ')}`,
      'warning',
      { providedTypes: entityTypes }
    );
    MonitoringService.logError(error);
    throw error;
  }

  // Development debug logging
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Search operation started', {
      query: query.substring(0, 100),
      entityTypes: validatedTypes,
      from,
      size,
      sessionId: contextSessionId,
      timestamp: new Date().toISOString()
    }, 'search');
  }

  try {
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

    // Microsoft Graph Search requires certain entity types to be searched separately
    // We need to split requests based on entity type restrictions
    const separateTypes = validatedTypes.filter(t => ENTITY_TYPE_GROUPS.separate.includes(t));
    const sharepointTypes = validatedTypes.filter(t => ENTITY_TYPE_GROUPS.sharepoint.includes(t));

    // Build search requests - one per separate type, plus one for SharePoint types
    const requests = [];

    // Add separate type requests (message, event, person each need their own request)
    for (const entityType of separateTypes) {
      requests.push({
        entityTypes: [entityType],
        query: { queryString: query },
        from,
        size: Math.min(size, 25) // Max 25 per request
      });
    }

    // Add combined SharePoint request if any SharePoint types requested
    if (sharepointTypes.length > 0) {
      requests.push({
        entityTypes: sharepointTypes,
        query: { queryString: query },
        from,
        size: Math.min(size, 25)
      });
    }

    // If no requests, something went wrong
    if (requests.length === 0) {
      throw new Error('No valid search requests could be constructed');
    }

    MonitoringService.debug('Executing search requests', {
      requestCount: requests.length,
      entityTypesPerRequest: requests.map(r => r.entityTypes),
      timestamp: new Date().toISOString()
    }, 'search');

    // Execute SEPARATE API calls for each request (Microsoft Graph doesn't allow
    // combining incompatible entity types even in separate request objects)
    // Use Promise.all for parallel execution
    const apiCalls = requests.map(request =>
      client.api('/search/query').version('beta').post({ requests: [request] })
    );

    const responses = await Promise.all(apiCalls);

    // Process results from all responses
    const allResults = [];
    let totalHits = 0;
    let moreResultsAvailable = false;

    for (const response of responses) {
      if (response.value && Array.isArray(response.value)) {
        for (const searchResponse of response.value) {
          if (searchResponse.hitsContainers && Array.isArray(searchResponse.hitsContainers)) {
            for (const container of searchResponse.hitsContainers) {
              totalHits += container.total || 0;
              moreResultsAvailable = moreResultsAvailable || container.moreResultsAvailable;

              if (container.hits && Array.isArray(container.hits)) {
                for (const hit of container.hits) {
                  // Determine entity type from the resource
                  const resourceType = hit.resource?.['@odata.type']?.replace('#microsoft.graph.', '') || 'unknown';
                  const normalizedHit = normalizeSearchHit(hit, resourceType);
                  allResults.push(normalizedHit);
                }
              }
            }
          }
        }
      }
    }

    // Sort results by rank
    allResults.sort((a, b) => (a.rank || 999) - (b.rank || 999));

    const executionTime = Date.now() - startTime;

    // User activity logging
    if (contextUserId) {
      MonitoringService.info('Search completed successfully', {
        query: query.substring(0, 50) + (query.length > 50 ? '...' : ''),
        entityTypes: validatedTypes,
        resultCount: allResults.length,
        totalHits,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'search', null, contextUserId);
    }

    // Track performance
    MonitoringService.trackMetric('search_query_time', executionTime, {
      entityTypes: validatedTypes.join(','),
      resultCount: allResults.length
    });

    return {
      query,
      entityTypes: validatedTypes,
      results: allResults,
      pagination: {
        from,
        size,
        total: totalHits,
        moreResultsAvailable
      },
      executionTimeMs: executionTime
    };

  } catch (error) {
    const executionTime = Date.now() - startTime;

    // Create standardized error
    const mcpError = ErrorService.createError(
      'search',
      `Search failed: ${error.message}`,
      'error',
      {
        query: query.substring(0, 50),
        entityTypes: validatedTypes,
        statusCode: error.statusCode || error.code,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );

    MonitoringService.logError(mcpError);

    if (contextUserId) {
      MonitoringService.error('Search operation failed', {
        query: query.substring(0, 50),
        error: error.message,
        timestamp: new Date().toISOString()
      }, 'search', null, contextUserId);
    }

    throw mcpError;
  }
}

/**
 * Searches only emails
 * @param {string} query - Search query
 * @param {object} options - Additional options
 * @param {object} req - Express request
 * @returns {Promise<object>} Search results
 */
async function searchMessages(query, options = {}, req, userId, sessionId) {
  return search({ ...options, query, entityTypes: ['message'] }, req, userId, sessionId);
}

/**
 * Searches only calendar events
 * @param {string} query - Search query
 * @param {object} options - Additional options
 * @param {object} req - Express request
 * @returns {Promise<object>} Search results
 */
async function searchEvents(query, options = {}, req, userId, sessionId) {
  return search({ ...options, query, entityTypes: ['event'] }, req, userId, sessionId);
}

/**
 * Searches only files
 * @param {string} query - Search query
 * @param {object} options - Additional options
 * @param {object} req - Express request
 * @returns {Promise<object>} Search results
 */
async function searchFiles(query, options = {}, req, userId, sessionId) {
  return search({ ...options, query, entityTypes: ['driveItem'] }, req, userId, sessionId);
}

/**
 * Searches only people
 * @param {string} query - Search query
 * @param {object} options - Additional options
 * @param {object} req - Express request
 * @returns {Promise<object>} Search results
 */
async function searchPeople(query, options = {}, req, userId, sessionId) {
  return search({ ...options, query, entityTypes: ['person'] }, req, userId, sessionId);
}

module.exports = {
  search,
  searchMessages,
  searchEvents,
  searchFiles,
  searchPeople,
  VALID_ENTITY_TYPES,
  normalizeSearchHit
};
