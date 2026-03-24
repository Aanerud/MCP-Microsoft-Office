/**
 * @fileoverview ExcelService - Microsoft Graph Excel (Workbook) API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

// --- Session Management ---

const SESSION_TTL = 240000; // 4 minutes
const sessionCache = new Map();
const pendingCreations = new Map();

// Clean expired sessions every 2 minutes
setInterval(() => {
  const now = Date.now();
  for (const [key, entry] of sessionCache) {
    if (now - entry.createdAt > SESSION_TTL) {
      sessionCache.delete(key);
    }
  }
}, 120000).unref();

/**
 * @param {string} fileId
 * @param {boolean} persistent
 * @param {object} req
 * @param {string} userId
 * @param {string} sessionId
 * @returns {Promise<string>} Graph workbook session ID
 */
async function createSession(fileId, persistent, req, userId, sessionId) {
  const client = await graphClientFactory.createClient(req, userId, sessionId);
  const path = `/me/drive/items/${fileId}/workbook/createSession`;
  const res = await client.api(path, userId, sessionId).post({ persistChanges: persistent });
  return res.id;
}

/**
 * @param {string} fileId
 * @param {string} graphSessionId
 * @param {object} req
 * @param {string} userId
 * @param {string} sessionId
 */
async function closeSession(fileId, graphSessionId, req, userId, sessionId) {
  const client = await graphClientFactory.createClient(req, userId, sessionId);
  const path = `/me/drive/items/${fileId}/workbook/closeSession`;
  await client.api(path, userId, sessionId).header('workbook-session-id', graphSessionId).post({});
}

function invalidateSession(userId, fileId) {
  sessionCache.delete(`${userId}:${fileId}`);
}

/**
 * @param {string} userId
 * @param {string} fileId
 * @param {object} req
 * @returns {Promise<{sessionId: string, createdAt: number, persistent: boolean}>}
 */
async function getOrCreateSession(userId, fileId, req) {
  const cacheKey = `${userId}:${fileId}`;

  // Return cached if still valid
  const cached = sessionCache.get(cacheKey);
  if (cached && (Date.now() - cached.createdAt < SESSION_TTL)) {
    return cached;
  }

  // Await any pending creation to avoid races
  const pending = pendingCreations.get(cacheKey);
  if (pending) {
    return pending;
  }

  // Remove expired entry
  if (cached) {
    sessionCache.delete(cacheKey);
  }

  const creationPromise = (async () => {
    try {
      const graphSessionId = await createSession(fileId, true, req, userId, null);
      const entry = { sessionId: graphSessionId, createdAt: Date.now(), persistent: true };
      sessionCache.set(cacheKey, entry);
      return entry;
    } finally {
      pendingCreations.delete(cacheKey);
    }
  })();

  pendingCreations.set(cacheKey, creationPromise);
  return creationPromise;
}

/**
 * Executes an operation within a workbook session, retrying once on invalid session.
 * @param {string} fileId
 * @param {object} req
 * @param {string} userId
 * @param {string} sessionId
 * @param {function} operation - async (graphSessionId) => result
 * @returns {Promise<*>}
 */
async function withSession(fileId, req, userId, sessionId, operation) {
  const resolvedUserId = userId || req?.user?.userId;
  const session = await getOrCreateSession(resolvedUserId, fileId, req);
  try {
    return await operation(session.sessionId);
  } catch (err) {
    if (err.statusCode === 404 || (err.code && err.code.includes('InvalidSession'))) {
      invalidateSession(resolvedUserId, fileId);
      const newSession = await getOrCreateSession(resolvedUserId, fileId, req);
      return await operation(newSession.sessionId);
    }
    throw err;
  }
}

// --- Helpers ---

const WB_PREFIX = (fileId) => `/me/drive/items/${fileId}/workbook`;

function enc(name) {
  return encodeURIComponent(name);
}

function resolveContext(req, userId, sessionId) {
  return {
    userId: userId || req?.user?.userId,
    sessionId: sessionId || req?.session?.id,
  };
}

/**
 * Standard error handler following the 4-pattern approach.
 */
function handleError(domain, message, error, context, resolvedUserId, resolvedSessionId, executionTime) {
  const mcpError = error.id
    ? error
    : ErrorService.createError(domain, `${message}: ${error.message}`, 'error', {
        ...context,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString(),
      });
  MonitoringService.logError(mcpError);

  if (resolvedUserId) {
    MonitoringService.error(message, {
      error: error.message,
      ...context,
      executionTimeMs: executionTime,
      timestamp: new Date().toISOString(),
    }, 'excel', null, resolvedUserId);
  } else if (resolvedSessionId) {
    MonitoringService.error(message, {
      sessionId: resolvedSessionId,
      error: error.message,
      ...context,
      executionTimeMs: executionTime,
      timestamp: new Date().toISOString(),
    }, 'excel');
  }

  MonitoringService.trackMetric(`excel_${domain}_failure`, executionTime, {
    errorType: error.code || 'unknown',
    ...context,
    userId: resolvedUserId,
    timestamp: new Date().toISOString(),
  });

  throw mcpError;
}

function trackSuccess(metricName, executionTime, context, resolvedUserId, resolvedSessionId) {
  if (resolvedUserId) {
    MonitoringService.info(`Excel ${metricName} completed`, {
      ...context,
      executionTimeMs: executionTime,
      timestamp: new Date().toISOString(),
    }, 'excel', null, resolvedUserId);
  } else if (resolvedSessionId) {
    MonitoringService.info(`Excel ${metricName} completed`, {
      sessionId: resolvedSessionId,
      ...context,
      executionTimeMs: executionTime,
      timestamp: new Date().toISOString(),
    }, 'excel');
  }

  MonitoringService.trackMetric(`excel_${metricName}_success`, executionTime, {
    ...context,
    userId: resolvedUserId,
    timestamp: new Date().toISOString(),
  });
}

// --- Session Public API ---

/**
 * @param {string} fileId
 * @param {boolean} persistent
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<{sessionId: string}>}
 */
async function createWorkbookSession(fileId, persistent, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Creating workbook session', { fileId, persistent, timestamp: new Date().toISOString() }, 'excel');
    }

    const graphSessionId = await createSession(fileId, persistent, req, userId, sessionId);
    const cacheKey = `${rUserId}:${fileId}`;
    const entry = { sessionId: graphSessionId, createdAt: Date.now(), persistent };
    sessionCache.set(cacheKey, entry);

    trackSuccess('create_session', Date.now() - startTime, { fileId }, rUserId, rSessionId);
    return { sessionId: graphSessionId };
  } catch (error) {
    handleError('create_session', 'Failed to create workbook session', error, { fileId }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 */
async function closeWorkbookSession(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);
  const cacheKey = `${rUserId}:${fileId}`;

  try {
    const cached = sessionCache.get(cacheKey);
    if (cached) {
      await closeSession(fileId, cached.sessionId, req, userId, sessionId);
      invalidateSession(rUserId, fileId);
    }
    trackSuccess('close_session', Date.now() - startTime, { fileId }, rUserId, rSessionId);
  } catch (error) {
    invalidateSession(rUserId, fileId);
    handleError('close_session', 'Failed to close workbook session', error, { fileId }, rUserId, rSessionId, Date.now() - startTime);
  }
}

// --- Worksheets ---

/**
 * @param {string} fileId
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Array>}
 */
async function listWorksheets(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('list_worksheets', Date.now() - startTime, { fileId }, rUserId, rSessionId);
      return res.value || [];
    });
  } catch (error) {
    handleError('list_worksheets', 'Failed to list worksheets', error, { fileId }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} name
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function addWorksheet(fileId, name, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ name });
      trackSuccess('add_worksheet', Date.now() - startTime, { fileId, name }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('add_worksheet', 'Failed to add worksheet', error, { fileId, name }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function getWorksheet(fileId, sheetIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('get_worksheet', Date.now() - startTime, { fileId, sheetIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('get_worksheet', 'Failed to get worksheet', error, { fileId, sheetIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {Object} properties - { name, position, visibility }
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function updateWorksheet(fileId, sheetIdOrName, properties, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).patch(properties);
      trackSuccess('update_worksheet', Date.now() - startTime, { fileId, sheetIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('update_worksheet', 'Failed to update worksheet', error, { fileId, sheetIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 */
async function deleteWorksheet(fileId, sheetIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')`;
      await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).delete();
      trackSuccess('delete_worksheet', Date.now() - startTime, { fileId, sheetIdOrName }, rUserId, rSessionId);
    });
  } catch (error) {
    handleError('delete_worksheet', 'Failed to delete worksheet', error, { fileId, sheetIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

// --- Ranges ---

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function getRange(fileId, sheetIdOrName, address, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('get_range', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('get_range', 'Failed to get range', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {Array<Array>} values - 2D array
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function updateRange(fileId, sheetIdOrName, address, values, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).patch({ values });
      trackSuccess('update_range', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('update_range', 'Failed to update range', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function getRangeFormat(fileId, sheetIdOrName, address, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')/format`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('get_range_format', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('get_range_format', 'Failed to get range format', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {Object} format
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function updateRangeFormat(fileId, sheetIdOrName, address, format, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const basePath = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')/format`;

      // Graph requires sub-resources (font, fill, borders, protection) to be patched at their own endpoints
      const subResources = ['font', 'fill', 'borders', 'protection'];
      const directProps = {};
      const results = {};

      for (const [key, value] of Object.entries(format)) {
        if (subResources.includes(key) && typeof value === 'object') {
          // Patch sub-resource at /format/{subResource}
          const subResult = await client.api(`${basePath}/${key}`, userId, sessionId)
            .header('workbook-session-id', wbSessionId).patch(value);
          results[key] = subResult;
        } else {
          directProps[key] = value;
        }
      }

      // Patch direct properties (columnWidth, rowHeight, horizontalAlignment, etc.) at /format
      if (Object.keys(directProps).length > 0) {
        const directResult = await client.api(basePath, userId, sessionId)
          .header('workbook-session-id', wbSessionId).patch(directProps);
        results._format = directResult;
      }

      trackSuccess('update_range_format', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return results;
    });
  } catch (error) {
    handleError('update_range_format', 'Failed to update range format', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {Array} fields - Sort fields
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function sortRange(fileId, sheetIdOrName, address, fields, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')/sort/apply`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ fields });
      trackSuccess('sort_range', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('sort_range', 'Failed to sort range', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {boolean} across
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function mergeRange(fileId, sheetIdOrName, address, across, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')/merge`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ across });
      trackSuccess('merge_range', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('merge_range', 'Failed to merge range', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function unmergeRange(fileId, sheetIdOrName, address, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/range(address='${address}')/unmerge`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({});
      trackSuccess('unmerge_range', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('unmerge_range', 'Failed to unmerge range', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

// --- Tables ---

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Array>}
 */
async function listTables(fileId, sheetIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/tables`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('list_tables', Date.now() - startTime, { fileId, sheetIdOrName }, rUserId, rSessionId);
      return res.value || [];
    });
  } catch (error) {
    handleError('list_tables', 'Failed to list tables', error, { fileId, sheetIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} sheetIdOrName
 * @param {string} address
 * @param {boolean} hasHeaders
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function createTable(fileId, sheetIdOrName, address, hasHeaders, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/worksheets('${enc(sheetIdOrName)}')/tables/add`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ address, hasHeaders });
      trackSuccess('create_table', Date.now() - startTime, { fileId, sheetIdOrName, address }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('create_table', 'Failed to create table', error, { fileId, sheetIdOrName, address }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {Object} properties
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function updateTable(fileId, tableIdOrName, properties, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).patch(properties);
      trackSuccess('update_table', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('update_table', 'Failed to update table', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 */
async function deleteTable(fileId, tableIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')`;
      await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).delete();
      trackSuccess('delete_table', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
    });
  } catch (error) {
    handleError('delete_table', 'Failed to delete table', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Array>}
 */
async function listTableRows(fileId, tableIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/rows`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('list_table_rows', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res.value || [];
    });
  } catch (error) {
    handleError('list_table_rows', 'Failed to list table rows', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {Array} values - Row values
 * @param {number} [index]
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function addTableRow(fileId, tableIdOrName, values, index, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/rows`;
      const body = { values: [values] };
      if (index !== undefined && index !== null) {
        body.index = index;
      }
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post(body);
      trackSuccess('add_table_row', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('add_table_row', 'Failed to add table row', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {number} index
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 */
async function deleteTableRow(fileId, tableIdOrName, index, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/rows/itemAt(index=${index})`;
      await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).delete();
      trackSuccess('delete_table_row', Date.now() - startTime, { fileId, tableIdOrName, index }, rUserId, rSessionId);
    });
  } catch (error) {
    handleError('delete_table_row', 'Failed to delete table row', error, { fileId, tableIdOrName, index }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Array>}
 */
async function listTableColumns(fileId, tableIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/columns`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).get();
      trackSuccess('list_table_columns', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res.value || [];
    });
  } catch (error) {
    handleError('list_table_columns', 'Failed to list table columns', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {Array} values - Column values
 * @param {number} [index]
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function addTableColumn(fileId, tableIdOrName, values, index, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/columns`;
      const body = { values };
      if (index !== undefined && index !== null) {
        body.index = index;
      }
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post(body);
      trackSuccess('add_table_column', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('add_table_column', 'Failed to add table column', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {string} columnIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 */
async function deleteTableColumn(fileId, tableIdOrName, columnIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/columns('${enc(columnIdOrName)}')`;
      await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).delete();
      trackSuccess('delete_table_column', Date.now() - startTime, { fileId, tableIdOrName, columnIdOrName }, rUserId, rSessionId);
    });
  } catch (error) {
    handleError('delete_table_column', 'Failed to delete table column', error, { fileId, tableIdOrName, columnIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {Array} fields - Sort fields
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function sortTable(fileId, tableIdOrName, fields, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/sort/apply`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ fields });
      trackSuccess('sort_table', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('sort_table', 'Failed to sort table', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {string} columnId
 * @param {Object} criteria
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function filterTable(fileId, tableIdOrName, columnId, criteria, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/columns(id='${columnId}')/filter/apply`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ criteria });
      trackSuccess('filter_table', Date.now() - startTime, { fileId, tableIdOrName, columnId }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('filter_table', 'Failed to filter table', error, { fileId, tableIdOrName, columnId }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {string} columnId
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function clearTableFilter(fileId, tableIdOrName, columnId, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/columns(id='${columnId}')/filter/clear`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({});
      trackSuccess('clear_table_filter', Date.now() - startTime, { fileId, tableIdOrName, columnId }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('clear_table_filter', 'Failed to clear table filter', error, { fileId, tableIdOrName, columnId }, rUserId, rSessionId, Date.now() - startTime);
  }
}

/**
 * @param {string} fileId
 * @param {string} tableIdOrName
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function convertTableToRange(fileId, tableIdOrName, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/tables('${enc(tableIdOrName)}')/convertToRange`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({});
      trackSuccess('convert_table_to_range', Date.now() - startTime, { fileId, tableIdOrName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('convert_table_to_range', 'Failed to convert table to range', error, { fileId, tableIdOrName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

// --- Functions ---

/**
 * @param {string} fileId
 * @param {string} functionName
 * @param {Object} args
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function callWorkbookFunction(fileId, functionName, args, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/functions/${functionName}`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post(args);
      trackSuccess('call_function', Date.now() - startTime, { fileId, functionName }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('call_function', 'Failed to call workbook function', error, { fileId, functionName }, rUserId, rSessionId, Date.now() - startTime);
  }
}

// --- Workbook ---

/**
 * @param {string} fileId
 * @param {string} [calculationType='Full']
 * @param {object} req
 * @param {string} [userId]
 * @param {string} [sessionId]
 * @returns {Promise<Object>}
 */
async function calculateWorkbook(fileId, calculationType, req, userId, sessionId) {
  const startTime = Date.now();
  const { userId: rUserId, sessionId: rSessionId } = resolveContext(req, userId, sessionId);
  const calcType = calculationType || 'Full';

  try {
    return await withSession(fileId, req, userId, sessionId, async (wbSessionId) => {
      const client = await graphClientFactory.createClient(req, userId, sessionId);
      const path = `${WB_PREFIX(fileId)}/application/calculate`;
      const res = await client.api(path, userId, sessionId).header('workbook-session-id', wbSessionId).post({ calculationType: calcType });
      trackSuccess('calculate_workbook', Date.now() - startTime, { fileId, calculationType: calcType }, rUserId, rSessionId);
      return res;
    });
  } catch (error) {
    handleError('calculate_workbook', 'Failed to calculate workbook', error, { fileId, calculationType: calcType }, rUserId, rSessionId, Date.now() - startTime);
  }
}

module.exports = {
  // Sessions
  createWorkbookSession,
  closeWorkbookSession,
  // Worksheets
  listWorksheets,
  addWorksheet,
  getWorksheet,
  updateWorksheet,
  deleteWorksheet,
  // Ranges
  getRange,
  updateRange,
  getRangeFormat,
  updateRangeFormat,
  sortRange,
  mergeRange,
  unmergeRange,
  // Tables
  listTables,
  createTable,
  updateTable,
  deleteTable,
  listTableRows,
  addTableRow,
  deleteTableRow,
  listTableColumns,
  addTableColumn,
  deleteTableColumn,
  sortTable,
  filterTable,
  clearTableFilter,
  convertTableToRange,
  // Functions
  callWorkbookFunction,
  // Workbook
  calculateWorkbook,
};
