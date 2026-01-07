/**
 * @fileoverview SearchService - Microsoft Graph Search API operations.
 * Provides unified search across emails, calendar events, files, and people.
 * Uses the /search/query endpoint for cross-entity searching with KQL support.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const { buildOutlookItemUrl, buildCalendarEventUrl, buildSharePointFileUrl, buildSearchUrl } = require('./normalizers.cjs');

// Log service initialization
MonitoringService.info('Graph Search Service initialized', {
  serviceName: 'graph-search-service',
  supportedEntityTypes: ['message', 'event', 'driveItem', 'person'],
  timestamp: new Date().toISOString()
}, 'graph');

/**
 * Valid entity types for Microsoft Graph Search API
 * See: https://learn.microsoft.com/en-us/graph/search-concept-interleaving
 */
const VALID_ENTITY_TYPES = ['message', 'event', 'driveItem', 'person', 'chatMessage', 'site', 'list', 'listItem'];

/**
 * Answer entity types for enterprise knowledge
 * See: https://learn.microsoft.com/en-us/graph/search-concept-answers
 */
const ANSWER_ENTITY_TYPES = ['acronym', 'bookmark', 'qna'];

/**
 * Entity type combinations that can be interleaved together
 * Based on Microsoft Graph interleaving rules:
 * - messages group: message, chatMessage (can interleave)
 * - sharepoint group: driveItem, site, list, listItem (can interleave)
 * - standalone: event, person (each must be searched alone)
 */
const ENTITY_INTERLEAVE_GROUPS = {
  messages: ['message', 'chatMessage'],
  sharepoint: ['driveItem', 'site', 'list', 'listItem'],
  standalone: ['event', 'person']  // Each searched separately
};

/**
 * Get the interleave group for an entity type
 */
function getInterleavGroup(entityType) {
  for (const [group, types] of Object.entries(ENTITY_INTERLEAVE_GROUPS)) {
    if (types.includes(entityType)) return group;
  }
  return 'standalone';
}

/**
 * Group entity types by their interleaving compatibility
 */
function groupEntityTypes(entityTypes) {
  const groups = {
    messages: [],
    sharepoint: [],
    standalone: []
  };

  for (const type of entityTypes) {
    const group = getInterleavGroup(type);
    if (group === 'standalone') {
      // Standalone types each get their own array
      groups.standalone.push(type);
    } else {
      groups[group].push(type);
    }
  }

  return groups;
}

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
        webLink: resource.webLink || buildOutlookItemUrl(resource.id)
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
        webLink: resource.webLink || buildCalendarEventUrl(resource.id)
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
 * Normalizes an answer hit (acronym, bookmark, qna) to a consistent format
 * See: https://learn.microsoft.com/en-us/graph/search-concept-answers
 * @param {object} hit - Search hit from Graph API
 * @param {string} entityType - The answer entity type
 * @returns {object} Normalized answer result
 */
function normalizeAnswerHit(hit, entityType) {
  const resource = hit.resource || {};

  const base = {
    id: resource.id || hit.hitId,
    entityType,
    rank: hit.rank
  };

  switch (entityType) {
    case 'acronym':
      return {
        ...base,
        displayName: resource.displayName,
        standsFor: resource.standsFor,
        description: resource.description,
        webUrl: resource.webUrl
      };

    case 'bookmark':
      return {
        ...base,
        displayName: resource.displayName,
        description: resource.description,
        webUrl: resource.webUrl,
        keywords: resource.keywords
      };

    case 'qna':
      return {
        ...base,
        displayName: resource.displayName,
        question: resource.displayName,
        answer: resource.description,
        webUrl: resource.webUrl
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
 * @param {boolean} [options.enableSpellingSuggestion=true] - Return spelling suggestions
 * @param {boolean} [options.enableSpellingModification=false] - Auto-correct query typos
 * @param {boolean} [options.enableTopResults=true] - Relevance-ranked results for messages
 * @param {boolean} [options.includeAnswers=false] - Include acronym, bookmark, qna results
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
    fields,
    enableSpellingSuggestion = true,
    enableSpellingModification = false,
    enableTopResults = true,
    includeAnswers = false
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

    // Group entity types by interleaving compatibility
    // See: https://learn.microsoft.com/en-us/graph/search-concept-interleaving
    const groups = groupEntityTypes(validatedTypes);
    const requests = [];

    // Build query object with speller options
    // See: https://learn.microsoft.com/en-us/graph/search-concept-speller
    const queryConfig = {
      queryString: query
    };

    // Add speller configuration (only one mode can be active)
    if (enableSpellingModification) {
      queryConfig.queryTemplate = '{searchTerms}';  // Required for speller
    }

    // Messages group (message, chatMessage can interleave)
    if (groups.messages.length > 0) {
      const messageRequest = {
        entityTypes: groups.messages,
        query: queryConfig,
        from,
        size: Math.min(size, 25)
      };
      // enableTopResults improves message relevance ranking
      if (enableTopResults) {
        messageRequest.enableTopResults = true;
      }
      requests.push(messageRequest);
    }

    // SharePoint group (driveItem, site, list, listItem can interleave)
    if (groups.sharepoint.length > 0) {
      requests.push({
        entityTypes: groups.sharepoint,
        query: queryConfig,
        from,
        size: Math.min(size, 25)
      });
    }

    // Standalone types (event, person - each must be searched separately)
    for (const entityType of groups.standalone) {
      requests.push({
        entityTypes: [entityType],
        query: queryConfig,
        from,
        size: Math.min(size, 25)
      });
    }

    // Add answer types if requested (acronym, bookmark, qna)
    // See: https://learn.microsoft.com/en-us/graph/search-concept-answers
    if (includeAnswers) {
      requests.push({
        entityTypes: ANSWER_ENTITY_TYPES,
        query: queryConfig,
        from: 0,
        size: 10  // Answers typically have fewer results
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
    let spellingAlteration = null;  // Track spelling suggestions/modifications
    const answers = [];  // Track answer entities (acronym, bookmark, qna)

    for (const response of responses) {
      if (response.value && Array.isArray(response.value)) {
        for (const searchResponse of response.value) {
          // Capture spelling alteration if present
          // See: https://learn.microsoft.com/en-us/graph/search-concept-speller
          if (searchResponse.queryAlterationResponse) {
            spellingAlteration = {
              originalQuery: searchResponse.queryAlterationResponse.originalQueryString,
              alteredQuery: searchResponse.queryAlterationResponse.queryAlteration?.alteredQueryString,
              alterationType: searchResponse.queryAlterationResponse.queryAlterationType
            };
          }

          if (searchResponse.hitsContainers && Array.isArray(searchResponse.hitsContainers)) {
            for (const container of searchResponse.hitsContainers) {
              totalHits += container.total || 0;
              moreResultsAvailable = moreResultsAvailable || container.moreResultsAvailable;

              if (container.hits && Array.isArray(container.hits)) {
                for (const hit of container.hits) {
                  // Determine entity type from the resource
                  const resourceType = hit.resource?.['@odata.type']?.replace('#microsoft.graph.', '') || 'unknown';

                  // Separate answer types from regular results
                  if (ANSWER_ENTITY_TYPES.includes(resourceType)) {
                    answers.push(normalizeAnswerHit(hit, resourceType));
                  } else {
                    const normalizedHit = normalizeSearchHit(hit, resourceType);
                    allResults.push(normalizedHit);
                  }
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

    // Build response object
    const response = {
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

    // Add spelling alteration if present (typo correction/suggestion)
    if (spellingAlteration && spellingAlteration.alteredQuery) {
      response.spelling = spellingAlteration;
    }

    // Add answers if any were found (acronym, bookmark, qna)
    if (answers.length > 0) {
      response.answers = answers;
    }

    return response;

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
  ANSWER_ENTITY_TYPES,
  normalizeSearchHit,
  normalizeAnswerHit
};
