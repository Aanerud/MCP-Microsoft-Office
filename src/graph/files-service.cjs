/**
 * @fileoverview FilesService - Microsoft Graph Files (OneDrive) API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

/**
 * Normalizes a file object from Graph API to a standard format
 * @param {Object} graphFile - File object from Graph API
 * @returns {Object} Normalized file object
 */
function normalizeFile(graphFile) {
  return {
    id: graphFile.id,
    name: graphFile.name,
    size: graphFile.size,
    lastModified: graphFile.lastModifiedDateTime,
    isFolder: !!graphFile.folder,
    webUrl: graphFile.webUrl
  };
}

/**
 * Redacts sensitive data from objects before logging
 * @param {Object} data - The data object to redact
 * @returns {Object} Redacted copy of the data
 * @private
 */
function redactSensitiveData(data) {
  if (!data || typeof data !== 'object') {
    return data;
  }
  
  // Create a deep copy to avoid modifying the original
  const result = Array.isArray(data) ? [...data] : {...data};
  
  // Fields that should be redacted
  const sensitiveFields = [
    'user', 'email', 'mail', 'address', 'emailAddress', 'password', 'token', 'accessToken',
    'refreshToken', 'content', 'body', 'contentBytes'
  ];
  
  // Recursively process the object
  for (const key in result) {
    if (Object.prototype.hasOwnProperty.call(result, key)) {
      // Check if this is a sensitive field
      if (sensitiveFields.includes(key.toLowerCase())) {
        if (typeof result[key] === 'string') {
          result[key] = 'REDACTED';
        } else if (Array.isArray(result[key])) {
          result[key] = `[${result[key].length} items]`;
        } else if (typeof result[key] === 'object' && result[key] !== null) {
          result[key] = '{REDACTED}';
        }
      } 
      // Recursively process nested objects
      else if (typeof result[key] === 'object' && result[key] !== null) {
        result[key] = redactSensitiveData(result[key]);
      }
    }
  }
  
  return result;
}

/**
 * Lists files and folders in a directory (defaults to root).
 * @param {string} [parentId]
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function listFiles(parentId, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Files listing requested', {
        parentId: parentId || 'root',
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    const path = parentId ? `/me/drive/items/${parentId}/children` : '/me/drive/root/children';
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for files listing', {
        path,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const res = await client.api(path).get();
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('Files listing completed successfully', {
        fileCount: res.value ? res.value.length : 0,
        parentId: parentId || 'root',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('Files listing completed with session', {
        sessionId: resolvedSessionId,
        fileCount: res.value ? res.value.length : 0,
        parentId: parentId || 'root',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('files_list_success', executionTime, {
      fileCount: res.value ? res.value.length : 0,
      parentId: parentId || 'root',
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return (res.value || []).map(normalizeFile);
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to list files: ${error.message}`,
      'error',
      { 
        parentId: parentId || 'root',
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('Files listing failed', {
        error: error.message,
        parentId: parentId || 'root',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files listing failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        parentId: parentId || 'root',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('files_list_failure', executionTime, {
      errorType: error.code || 'unknown',
      parentId: parentId || 'root',
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Searches files by name.
 * @param {string} query
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function searchFiles(query, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!query) {
    const mcpError = ErrorService.createError(
      'files',
      'Search query is required for searchFiles',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_query'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Files search validation failed: missing query', {
        validationError: 'missing_query',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files search validation failed: missing query', {
        sessionId: resolvedSessionId,
        validationError: 'missing_query',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Files search requested', {
        query: redactSensitiveData(query),
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    const searchUrl = `/me/drive/root/search(q='${encodeURIComponent(query)}')`;
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for files search', {
        searchUrl: redactSensitiveData(searchUrl),
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const res = await client.api(searchUrl).get();
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('Files search completed successfully', {
        resultCount: res.value ? res.value.length : 0,
        queryLength: query.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('Files search completed with session', {
        sessionId: resolvedSessionId,
        resultCount: res.value ? res.value.length : 0,
        queryLength: query.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('files_search_success', executionTime, {
      resultCount: res.value ? res.value.length : 0,
      queryLength: query.length,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return (res.value || []).map(normalizeFile);
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to search files: ${error.message}`,
      'error',
      { 
        query: redactSensitiveData(query),
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('Files search failed', {
        error: error.message,
        queryLength: query.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files search failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        queryLength: query.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('files_search_failure', executionTime, {
      errorType: error.code || 'unknown',
      queryLength: query.length,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Downloads a file by ID.
 * @param {string} id
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Buffer>}
 */
async function downloadFile(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for downloadFile',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Files download validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files download validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('File download requested', {
        fileId: id,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for file download', {
        fileId: id,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Get the raw content from the Graph API
    const content = await client.api(`/me/drive/items/${id}/content`).get();
    
    // Ensure content is always returned as a Buffer
    let bufferContent;
    if (Buffer.isBuffer(content)) {
      bufferContent = content;
    } else if (typeof content === 'string') {
      bufferContent = Buffer.from(content);
    } else if (content && typeof content === 'object') {
      // If it's an object (like a response object), stringify it
      bufferContent = Buffer.from(JSON.stringify(content));
    } else {
      // Fallback for null/undefined
      bufferContent = Buffer.from('');
    }
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('File download completed successfully', {
        fileId: id,
        contentSize: bufferContent.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('File download completed with session', {
        sessionId: resolvedSessionId,
        fileId: id,
        contentSize: bufferContent.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_download_success', executionTime, {
      fileId: id,
      contentSize: bufferContent.length,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return bufferContent;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to download file: ${error.message}`,
      'error',
      { 
        fileId: id,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('File download failed', {
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File download failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_download_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId: id,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Uploads a file to root directory.
 * @param {string} name - The name of the file to upload
 * @param {string|Buffer} content - The content of the file
 * @param {object} req - Express request object for authentication
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<object>} - The uploaded file metadata
 */
async function uploadFile(name, content, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!name) {
    const mcpError = ErrorService.createError(
      'files',
      'File name is required for uploadFile',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_name'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Files upload validation failed: missing file name', {
        validationError: 'missing_file_name',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files upload validation failed: missing file name', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_name',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  if (!content) {
    const mcpError = ErrorService.createError(
      'files',
      'File content is required for uploadFile',
      'error',
      { 
        fileName: name,
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_content'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Files upload validation failed: missing file content', {
        fileName: name,
        validationError: 'missing_file_content',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Files upload validation failed: missing file content', {
        sessionId: resolvedSessionId,
        fileName: name,
        validationError: 'missing_file_content',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('File upload requested', {
        fileName: name,
        fileSize: content.length || 'unknown',
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Ensure we have a valid client with authentication
    const client = await graphClientFactory.createClient(req);
    
    // Convert string content to Buffer if needed
    const contentBuffer = typeof content === 'string' ? Buffer.from(content) : content;
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for file upload', {
        fileName: name,
        fileSize: contentBuffer.length,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Make the API call to upload the file
    const result = await client.api(`/me/drive/root:/${encodeURIComponent(name)}:/content`).put(contentBuffer);
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('File upload completed successfully', {
        fileName: name,
        fileId: result.id,
        fileSize: contentBuffer.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('File upload completed with session', {
        sessionId: resolvedSessionId,
        fileName: name,
        fileId: result.id,
        fileSize: contentBuffer.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_upload_success', executionTime, {
      fileName: name,
      fileSize: contentBuffer.length,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to upload file: ${error.message}`,
      'error',
      {
        fileName: name,
        fileSize: content.length || 'unknown',
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('File upload failed', {
        error: error.message,
        fileName: name,
        fileSize: content.length || 'unknown',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File upload failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileName: name,
        fileSize: content.length || 'unknown',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_upload_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileName: name,
      fileSize: content.length || 'unknown',
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Retrieves metadata for a file by ID.
 * @param {string} id
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<object>}
 */
async function getFileMetadata(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for getFileMetadata',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('File metadata validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File metadata validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('File metadata requested', {
        fileId: id,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for file metadata', {
        fileId: id,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const meta = await client.api(`/me/drive/items/${id}`).get();
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('File metadata retrieved successfully', {
        fileId: id,
        fileName: meta.name,
        fileSize: meta.size,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('File metadata retrieved with session', {
        sessionId: resolvedSessionId,
        fileId: id,
        fileName: meta.name,
        fileSize: meta.size,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_metadata_success', executionTime, {
      fileId: id,
      fileName: meta.name,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return normalizeFile(meta);
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to retrieve file metadata: ${error.message}`,
      'error',
      { 
        fileId: id,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('File metadata retrieval failed', {
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File metadata retrieval failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_metadata_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId: id,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Creates a sharing link for a file.
 * @param {string} id
 * @param {string} type - 'view', 'edit', etc.
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<object>} Normalized sharing link
 */
async function createSharingLink(id, type = 'view', req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for createSharingLink',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Sharing link creation validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Sharing link creation validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Sharing link creation requested', {
        fileId: id,
        linkType: type,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    // Microsoft Graph API expects specific parameters for createLink
    // Try organization scope first as anonymous might be blocked in demo tenants
    const body = {
      type: type === 'edit' ? 'edit' : 'view',
      scope: 'organization'  // Changed from 'anonymous' to 'organization'
    };
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for sharing link creation', {
        fileId: id,
        linkType: body.type,
        scope: body.scope,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const res = await client.api(`/me/drive/items/${id}/createLink`).post(body);
    const executionTime = Date.now() - startTime;
    
    // Return normalized response
    if (res && res.link && res.link.webUrl) {
      const result = { 
        webUrl: res.link.webUrl,
        type: res.link.type || body.type,
        scope: res.link.scope || body.scope
      };
      
      // Pattern 2: User Activity Logs
      if (resolvedUserId) {
        MonitoringService.info('Sharing link created successfully', {
          fileId: id,
          linkType: result.type,
          scope: result.scope,
          executionTimeMs: executionTime,
          timestamp: new Date().toISOString()
        }, 'files', null, resolvedUserId);
      } else if (resolvedSessionId) {
        MonitoringService.info('Sharing link created with session', {
          sessionId: resolvedSessionId,
          fileId: id,
          linkType: result.type,
          scope: result.scope,
          executionTimeMs: executionTime,
          timestamp: new Date().toISOString()
        }, 'files');
      }
      
      MonitoringService.trackMetric('sharing_link_create_success', executionTime, {
        fileId: id,
        linkType: result.type,
        userId: resolvedUserId,
        timestamp: new Date().toISOString()
      });
      
      return result;
    }
    
    // Handle case where no sharing link was returned
    const noLinkError = ErrorService.createError(
      'files',
      'No sharing link returned from API',
      'error',
      { 
        fileId: id,
        linkType: type,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(noLinkError);
    
    return { error: 'No sharing link returned from API' };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to create sharing link: ${error.message}`,
      'error',
      { 
        fileId: id,
        linkType: type,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('Sharing link creation failed', {
        error: error.message,
        fileId: id,
        linkType: type,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Sharing link creation failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        linkType: type,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('sharing_link_create_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId: id,
      linkType: type,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Gets sharing links (permissions) for a file.
 * @param {string} id
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized sharing links
 */
async function getSharingLinks(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for getSharingLinks',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Get sharing links validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Get sharing links validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Get sharing links requested', {
        fileId: id,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for sharing links', {
        fileId: id,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const res = await client.api(`/me/drive/items/${id}/permissions`).get();
    const filteredLinks = (res.value || [])
      .filter(p => p.link && p.link.webUrl)
      .map(p => ({ id: p.id, webUrl: p.link.webUrl, type: p.link.type }));
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('Sharing links retrieved successfully', {
        fileId: id,
        linkCount: filteredLinks.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('Sharing links retrieved with session', {
        sessionId: resolvedSessionId,
        fileId: id,
        linkCount: filteredLinks.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('sharing_links_get_success', executionTime, {
      fileId: id,
      linkCount: filteredLinks.length,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return filteredLinks;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to get sharing links: ${error.message}`,
      'error',
      { 
        fileId: id,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('Get sharing links failed', {
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Get sharing links failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('sharing_links_get_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId: id,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Removes a sharing permission from a file.
 * @param {string} fileId
 * @param {string} permissionId
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<object>}
 */
async function removeSharingPermission(fileId, permissionId, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input parameters
  if (!fileId) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for removeSharingPermission',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Remove sharing permission validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Remove sharing permission validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  if (!permissionId) {
    const mcpError = ErrorService.createError(
      'files',
      'Permission ID is required for removeSharingPermission',
      'error',
      { 
        fileId,
        timestamp: new Date().toISOString(),
        validationError: 'missing_permission_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('Remove sharing permission validation failed: missing permission ID', {
        fileId,
        validationError: 'missing_permission_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Remove sharing permission validation failed: missing permission ID', {
        sessionId: resolvedSessionId,
        fileId,
        validationError: 'missing_permission_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Remove sharing permission requested', {
        fileId,
        permissionId,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API to remove sharing permission', {
        fileId,
        permissionId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Delete the permission
    const result = await client.api(`/me/drive/items/${fileId}/permissions/${permissionId}`).delete();
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('Sharing permission removed successfully', {
        fileId,
        permissionId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('Sharing permission removed with session', {
        sessionId: resolvedSessionId,
        fileId,
        permissionId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('sharing_permission_remove_success', executionTime, {
      fileId,
      permissionId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return { success: true, fileId, permissionId };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to remove sharing permission: ${error.message}`,
      'error',
      {
        fileId,
        permissionId,
        error: error.message,
        statusCode: error.statusCode,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('Remove sharing permission failed', {
        error: error.message,
        fileId,
        permissionId,
        statusCode: error.statusCode,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('Remove sharing permission failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId,
        permissionId,
        statusCode: error.statusCode,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('sharing_permission_remove_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      permissionId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Gets file content by ID from OneDrive/SharePoint
 * @param {string} id - The file ID to retrieve
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<{content: Buffer, contentType: string}>} File content and content type
 */
async function getFileContent(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for getFileContent',
      'error',
      { 
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error('File content retrieval validation failed: missing file ID', {
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File content retrieval validation failed: missing file ID', {
        sessionId: resolvedSessionId,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('File content requested', {
        fileId: id,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for file metadata', {
        fileId: id,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // First get file metadata to determine content type
    const fileMetadata = await client.api(`/me/drive/items/${id}`).get();
    const contentType = fileMetadata.file ? fileMetadata.file.mimeType : 'application/octet-stream';
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Calling Graph API for file content', {
        fileId: id,
        contentType,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Get the actual file content
    const content = await client.api(`/me/drive/items/${id}/content`).get();
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info('File content retrieved successfully', {
        fileId: id,
        contentType,
        contentSize: content ? content.length : 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info('File content retrieved with session', {
        sessionId: resolvedSessionId,
        fileId: id,
        contentType,
        contentSize: content ? content.length : 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_content_success', executionTime, {
      fileId: id,
      contentType,
      contentSize: content ? content.length : 0,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    // Return both the content and content type
    return {
      content: content,
      contentType: contentType
    };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to retrieve file content: ${error.message}`,
      'error',
      { 
        fileId: id,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error('File content retrieval failed', {
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error('File content retrieval failed', {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric('file_content_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId: id,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

// setFileContent functionality has been consolidated into updateFileContent with options.setContent=true

/**
 * Updates file content by ID using PUT method as required by Microsoft Graph API.
 * This function handles both update and set operations (they use the same Graph API endpoint).
 * 
 * @param {string} id - File ID
 * @param {Buffer} content - File content
 * @param {object} req - Express request object
 * @param {object} [options] - Additional options
 * @param {boolean} [options.setContent=false] - If true, operation is considered a 'set' operation (for metrics/logging)
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<object>} - Updated file metadata
 */
async function updateFileContent(id, content, req, options = {}, userId, sessionId) {
  const startTime = Date.now();
  const operationType = options.setContent ? 'set' : 'update';
  
  // Extract user context from request if not provided
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;
  
  // Validate input
  if (!id) {
    const mcpError = ErrorService.createError(
      'files',
      'File ID is required for updateFileContent',
      'error',
      { 
        operationType,
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_id'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error(`File content ${operationType} validation failed: missing file ID`, {
        operationType,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error(`File content ${operationType} validation failed: missing file ID`, {
        sessionId: resolvedSessionId,
        operationType,
        validationError: 'missing_file_id',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  if (!content) {
    const mcpError = ErrorService.createError(
      'files',
      'File content is required for updateFileContent',
      'error',
      { 
        fileId: id,
        operationType,
        timestamp: new Date().toISOString(),
        validationError: 'missing_file_content'
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking for validation
    if (resolvedUserId) {
      MonitoringService.error(`File content ${operationType} validation failed: missing file content`, {
        fileId: id,
        operationType,
        validationError: 'missing_file_content',
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error(`File content ${operationType} validation failed: missing file content`, {
        sessionId: resolvedSessionId,
        fileId: id,
        operationType,
        validationError: 'missing_file_content',
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    throw mcpError;
  }
  
  try {
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug(`File content ${operationType} requested`, {
        fileId: id,
        contentSize: content?.length || 0,
        operationType,
        sessionId: resolvedSessionId,
        userAgent: req?.get('User-Agent'),
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    const client = await graphClientFactory.createClient(req);
    
    // Ensure content is properly formatted for Graph API
    let formattedContent = content;
    if (typeof content === 'string') {
      formattedContent = Buffer.from(content, 'utf8');
    } else if (!Buffer.isBuffer(content)) {
      formattedContent = Buffer.from(String(content), 'utf8');
    }
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Getting file metadata for content update', {
        fileId: id,
        operationType,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // First get the file metadata to get the file name/path
    const fileMetadata = await client.api(`/me/drive/items/${id}`).get();
    
    if (!fileMetadata || !fileMetadata.name) {
      throw new Error('Could not retrieve file metadata for content update');
    }
    
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug(`Calling Graph API for file content ${operationType}`, {
        fileId: id,
        fileName: fileMetadata.name,
        contentSize: formattedContent.length,
        operationType,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    // Use the same approach as uploadFile - path-based API which works better
    const result = await client.api(`/me/drive/root:/${encodeURIComponent(fileMetadata.name)}:/content`)
      .put(formattedContent);
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (resolvedUserId) {
      MonitoringService.info(`File content ${operationType} completed successfully`, {
        fileId: id,
        fileName: fileMetadata.name,
        contentSize: formattedContent.length,
        operationType,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.info(`File content ${operationType} completed with session`, {
        sessionId: resolvedSessionId,
        fileId: id,
        fileName: fileMetadata.name,
        contentSize: formattedContent.length,
        operationType,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric(`file_content_${operationType}_success`, executionTime, {
      fileId: id,
      contentSize: formattedContent.length,
      operationType,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'files',
      `Failed to ${operationType} file content: ${error.message}`,
      'error',
      {
        fileId: id,
        operationType,
        error: error.message,
        errorCode: error.statusCode || error.code,
        contentSize: content?.length || 0,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (resolvedUserId) {
      MonitoringService.error(`File content ${operationType} failed`, {
        error: error.message,
        fileId: id,
        operationType,
        errorCode: error.statusCode || error.code,
        contentSize: content?.length || 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files', null, resolvedUserId);
    } else if (resolvedSessionId) {
      MonitoringService.error(`File content ${operationType} failed`, {
        sessionId: resolvedSessionId,
        error: error.message,
        fileId: id,
        operationType,
        errorCode: error.statusCode || error.code,
        contentSize: content?.length || 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'files');
    }
    
    MonitoringService.trackMetric(`file_content_${operationType}_failure`, executionTime, {
      fileId: id,
      operationType,
      errorCode: error.statusCode || error.code || 'unknown',
      contentSize: content?.length || 0,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Resolves a file by OneDrive path (e.g., 'Documents/report.xlsx')
 * @param {string} filePath - OneDrive-rooted path
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for context
 * @param {string} [sessionId] - Session ID for context
 * @returns {Promise<object>} Graph driveItem object
 */
async function getFileByPath(filePath, req, userId = null, sessionId = null) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const result = await client.api(`/me/drive/root:/${filePath}`, resolvedUserId, resolvedSessionId).get();

    const executionTime = Date.now() - startTime;
    MonitoringService.trackMetric('file_get_by_path', executionTime, {
      path: filePath.substring(0, 50),
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = error.id ? error : ErrorService.createError(
      'files',
      `Failed to resolve file by path: ${error.message}`,
      'error',
      { path: filePath, error: error.message, stack: error.stack, timestamp: new Date().toISOString() }
    );
    MonitoringService.logError(mcpError);
    MonitoringService.trackMetric('file_get_by_path_failure', executionTime, {
      path: filePath.substring(0, 50),
      errorCode: error.statusCode || error.code || 'unknown',
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });
    throw mcpError;
  }
}

/**
 * Gets file size without downloading the file content
 * @param {string} fileId - Drive item ID
 * @param {object} req - Express request object
 * @param {string} [userId] - User ID for context
 * @param {string} [sessionId] - Session ID for context
 * @returns {Promise<number>} File size in bytes
 */
async function getFileSize(fileId, req, userId = null, sessionId = null) {
  const metadata = await getFileMetadata(fileId, req, userId, sessionId);
  return metadata.size || 0;
}

// Export all file service methods
module.exports = {
  listFiles,              // List files and folders in a directory
  searchFiles,            // Search files by name
  downloadFile,           // Handles both file content download and metadata retrieval with options.metadataOnly
  uploadFile,             // Handles file uploads
  getFileMetadata,        // Get file metadata by ID
  getFileContent,         // Get file content by ID
  updateFileContent,      // Handles both update and set operations with options.setContent
  createSharingLink,      // Create sharing links for files
  getSharingLinks,        // Get sharing links for files
  removeSharingPermission, // Remove sharing permissions from files
  getFileByPath,          // Resolve file by OneDrive path
  getFileSize             // Get file size without downloading
};
