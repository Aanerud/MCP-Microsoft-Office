/**
 * @fileoverview Handles /api/files endpoints for file operations.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const filesModule = require('../../modules/files');

/**
 * Helper function to call module methods with fallback support
 * @param {string} methodName - Name of the method to call
 * @param {Array} args - Arguments to pass to the method
 * @param {Function} fallback - Fallback function to call if method fails
 * @param {object} req - Express request object for context
 * @returns {Promise} Result of the method call or fallback
 */
async function callModuleWithFallback(methodName, args, fallback, req) {
    const { userId = null, deviceId = null } = req.user || {};
    
    try {
        // Try to call the module method
        if (typeof filesModule[methodName] === 'function') {
            return await filesModule[methodName](...args);
        } else {
            throw new Error(`Method ${methodName} not found on filesModule`);
        }
    } catch (error) {
        // Log the error
        MonitoringService.warn(`Module method ${methodName} failed, using fallback`, {
            methodName,
            error: error.message,
            userId,
            deviceId,
            timestamp: new Date().toISOString()
        }, 'files', null, userId, deviceId);
        
        // Call fallback function
        return await fallback();
    }
}

/**
 * Factory for files controller with dependency injection.
 * @param {object} deps - { filesModule, errorService, monitoringService }
 */
module.exports = ({ filesModule, errorService = ErrorService, monitoringService = MonitoringService }) => {
    return {
    /**
     * GET /api/files
     */
    async listFiles(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Files listing requested', {
                    parentId: req.query.parentId || 'root',
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }, 'files');
            }
            
            // Call the module method directly with req as separate parameter
            const files = await filesModule.listFiles(req.query.parentId, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('Files listing completed successfully', {
                    fileCount: Array.isArray(files) ? files.length : 0,
                    parentId: req.query.parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('Files listing completed with session', {
                    sessionId: req.session.id,
                    fileCount: Array.isArray(files) ? files.length : 0,
                    parentId: req.query.parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_list_api_success', executionTime, {
                fileCount: Array.isArray(files) ? files.length : 0,
                parentId: req.query.parentId || 'root',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(files);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `Files listing error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files',
                    parentId: req.query.parentId || 'root',
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Files listing failed', {
                    error: err.message,
                    parentId: req.query.parentId || 'root',
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('Files listing failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    parentId: req.query.parentId || 'root',
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_list_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_LIST_FAILED',
                error_description: 'Failed to retrieve files list'
            });
        }
    },
    /**
     * POST /api/files/upload
     */
    async uploadFile(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Log the request with user context
        monitoringService.info(`Processing ${req.method} ${req.path}`, {
            method: req.method,
            path: req.path,
            fileName: req.body?.name,
            ip: req.ip,
            userId,
            deviceId
        }, 'files', null, userId, deviceId);
        
        try {
            // Validate input
            const schema = Joi.object({
                name: Joi.string().min(1).max(255).pattern(/^[^\u003c\u003e:"/\\|?*]+$/).required(),
                content: Joi.string().min(1).max(10485760).required() // 10MB limit
            });
            
            const { error, value } = schema.validate(req.body);
            if (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'File upload validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/upload',
                        error: 'Invalid request parameters',
                        validationDetails: error.details,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('File upload validation failed', {
                        error: 'Invalid request parameters',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('File upload validation failed', {
                        sessionId: req.session.id,
                        error: 'Invalid request parameters',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILE_UPLOAD_INVALID_REQUEST',
                    error_description: 'Invalid file upload parameters'
                });
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('File upload requested', { 
                    fileName: value.name,
                    fileSize: value.content.length,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method directly with req as separate parameter
            const result = await filesModule.uploadFile(value.name, value.content, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('File upload completed successfully', {
                    fileName: value.name,
                    fileSize: value.content.length,
                    fileId: result.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('File upload completed with session', {
                    sessionId: req.session.id,
                    fileName: value.name,
                    fileSize: value.content.length,
                    fileId: result.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_upload_api_success', executionTime, {
                fileName: value.name,
                fileSize: value.content.length,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `File upload error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/upload',
                    fileName: req.body?.name,
                    fileSize: req.body?.content?.length,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('File upload failed', {
                    error: err.message,
                    fileName: req.body?.name,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('File upload failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileName: req.body?.name,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_upload_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILE_UPLOAD_FAILED',
                error_description: 'Failed to upload file'
            });
        }
    },
    /**
     * GET /api/files/search
     * Searches for files in OneDrive/SharePoint using the query provided
     */
    async searchFiles(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!req.query.q) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'Files search validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/search',
                        error: 'Search query is required',
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('Files search validation failed', {
                        error: 'Search query is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('Files search validation failed', {
                        sessionId: req.session.id,
                        error: 'Search query is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_SEARCH_INVALID_REQUEST',
                    error_description: 'Search query is required'
                });
            }
            
            const query = req.query.q;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Files search requested', { 
                    query,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method directly with req as separate parameter
            const files = await filesModule.searchFiles(query, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('Files search completed successfully', {
                    query: req.query.q,
                    resultCount: Array.isArray(files) ? files.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('Files search completed with session', {
                    sessionId: req.session.id,
                    query: req.query.q,
                    resultCount: Array.isArray(files) ? files.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_search_api_success', executionTime, {
                query: req.query.q,
                resultCount: Array.isArray(files) ? files.length : 0,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(files);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `Files search error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/search',
                    query: req.query.q,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Files search failed', {
                    error: err.message,
                    query: req.query.q,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('Files search failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    query: req.query.q,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_search_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_SEARCH_FAILED',
                error_description: 'Failed to search files'
            });
        }
    },
    /**
     * GET /api/files/download
     */
    async downloadFile(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!req.query.id) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'File download validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/download',
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('File download validation failed', {
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('File download validation failed', {
                        sessionId: req.session.id,
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_DOWNLOAD_INVALID_REQUEST',
                    error_description: 'File ID is required'
                });
            }
            
            const fileId = req.query.id;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('File download requested', { 
                    fileId,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Check if metadata only is requested
            const metadataOnly = req.query.metadataOnly === 'true';
            
            // Log if metadata only is requested
            if (metadataOnly && process.env.NODE_ENV === 'development') {
                monitoringService.debug('File metadata only requested', { 
                    fileId,
                    sessionId: req.session?.id,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Create options object for downloadFile
            const options = { metadataOnly };
            
            // Try to download the file from the module, or return mock data if it fails
            const result = await callModuleWithFallback(
                'downloadFile',
                [fileId, options, req],
                () => {
                    // Return mock data for development/testing
                    if (metadataOnly) {
                        return {
                            id: fileId,
                            name: `Mock File ${fileId}`,
                            size: 12345,
                            createdDateTime: new Date().toISOString(),
                            lastModifiedDateTime: new Date().toISOString(),
                            webUrl: `https://example.com/files/${fileId}`
                        };
                    } else {
                        return Buffer.from(`This is mock content for file ${fileId}`);
                    }
                },
                req
            );
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Prepare response based on whether we got metadata or file content
            let response;
            
            if (metadataOnly) {
                // For metadata request, return the metadata directly
                response = result;
                
                // Log success metrics with user context for metadata
                monitoringService.trackMetric('files_metadata_api_success', executionTime, {
                    fileId: fileId,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                });
                
                // Log success with result summary and user context
                monitoringService.info('Files metadata API completed successfully', {
                    fileId: fileId,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }, 'files', null, userId, deviceId);
            } else {
                // For file content request, handle as before
                // Log success metrics with user context
                monitoringService.trackMetric('files_download_api_success', executionTime, {
                    fileId: fileId,
                    contentLength: result?.length || 0,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                });
                
                // Log success with result summary and user context
                monitoringService.info('Files download API completed successfully', {
                    fileId: fileId,
                    contentLength: result?.length || 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }, 'files', null, userId, deviceId);
                
                // For MCP adapter compatibility, return a JSON response with base64-encoded file content
                // instead of sending the raw binary data directly
                let base64Content;
                
                // Handle different content types properly
                if (Buffer.isBuffer(result)) {
                    // If it's already a buffer, convert to base64
                    base64Content = result.toString('base64');
                } else if (typeof result === 'string') {
                    // If it's a string, convert to buffer then to base64
                    base64Content = Buffer.from(result).toString('base64');
                } else if (result && typeof result === 'object') {
                    // If it's an object (like a response object), stringify it first
                    base64Content = Buffer.from(JSON.stringify(result)).toString('base64');
                } else {
                    // Fallback for null/undefined
                    base64Content = Buffer.from('').toString('base64');
                }
                
                response = {
                    id: fileId,
                    content: base64Content,
                    contentType: 'application/octet-stream',
                    filename: `file-${fileId}.bin`
                };
            }
            
            res.json(response);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Create standardized error with user context
            const mcpError = ErrorService.createError(
                ErrorService.CATEGORIES.API, 
                `Files download error: ${err.message}`, 
                ErrorService.SEVERITIES.ERROR, 
                { 
                    fileId: req.query.id,
                    stack: err.stack,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }
            );
            
            // Log the error
            monitoringService.logError(mcpError);
            
            // Track failure metrics with user context
            monitoringService.trackMetric('files_download_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ error: 'Internal error', message: err.message });
        }
    },
    /**
     * GET /api/files/metadata
     */
    async getFileMetadata(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!req.query.id) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'File metadata validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/metadata',
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('File metadata validation failed', {
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('File metadata validation failed', {
                        sessionId: req.session.id,
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_METADATA_INVALID_REQUEST',
                    error_description: 'File ID is required'
                });
            }
            
            const fileId = req.query.id;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('File metadata requested', { 
                    fileId,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method directly with req as separate parameter
            const metadata = await filesModule.getFileMetadata(fileId, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('File metadata retrieval completed successfully', {
                    fileId: fileId,
                    fileName: metadata?.name,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('File metadata retrieval completed with session', {
                    sessionId: req.session.id,
                    fileId: fileId,
                    fileName: metadata?.name,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_metadata_api_success', executionTime, {
                fileId: fileId,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(metadata);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `File metadata error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/metadata',
                    fileId: req.query.id,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('File metadata retrieval failed', {
                    error: err.message,
                    fileId: req.query.id,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('File metadata retrieval failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.query.id,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_metadata_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_METADATA_FAILED',
                error_description: 'Failed to retrieve file metadata'
            });
        }
    },

    /**
     * PUT /api/files/content
     */
    async setFileContent(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            const schema = Joi.object({
                fileId: Joi.string().required(),
                content: Joi.string().required()
            });
            
            const { error, value } = schema.validate(req.body);
            if (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'File content validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/content',
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('File content validation failed', {
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('File content validation failed', {
                        sessionId: req.session.id,
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_CONTENT_INVALID_REQUEST',
                    error_description: error.details[0].message
                });
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('File content update requested', { 
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method
            const result = await filesModule.setFileContent(value.fileId, value.content, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('File content update completed successfully', {
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('File content update completed with session', {
                    sessionId: req.session.id,
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_content_set_api_success', executionTime, {
                fileId: value.fileId,
                contentLength: value.content.length,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `File content update error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/content',
                    fileId: req.body?.fileId,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('File content update failed', {
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('File content update failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_content_set_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_CONTENT_UPDATE_FAILED',
                error_description: 'Failed to update file content'
            });
        }
    },

    /**
     * PATCH /api/files/content
     */
    async updateFileContent(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            const schema = Joi.object({
                fileId: Joi.string().required(),
                content: Joi.string().required()
            });
            
            const { error, value } = schema.validate(req.body);
            if (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'File content update validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/content',
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('File content update validation failed', {
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('File content update validation failed', {
                        sessionId: req.session.id,
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_CONTENT_UPDATE_INVALID_REQUEST',
                    error_description: error.details[0].message
                });
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('File content update requested', { 
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method
            const result = await filesModule.updateFileContent(value.fileId, value.content, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('File content update completed successfully', {
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('File content update completed with session', {
                    sessionId: req.session.id,
                    fileId: value.fileId,
                    contentLength: value.content.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_content_update_api_success', executionTime, {
                fileId: value.fileId,
                contentLength: value.content.length,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `File content update error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/content',
                    fileId: req.body?.fileId,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('File content update failed', {
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('File content update failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_content_update_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_CONTENT_UPDATE_FAILED',
                error_description: 'Failed to update file content'
            });
        }
    },

    /**
     * POST /api/files/sharing/link
     */
    async createSharingLink(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            const schema = Joi.object({
                fileId: Joi.string().required(),
                type: Joi.string().valid('view', 'edit').default('view'),
                scope: Joi.string().valid('anonymous', 'organization').default('anonymous')
            });
            
            const { error, value } = schema.validate(req.body);
            if (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'Sharing link creation validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/sharing/link',
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('Sharing link creation validation failed', {
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('Sharing link creation validation failed', {
                        sessionId: req.session.id,
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_SHARING_INVALID_REQUEST',
                    error_description: error.details[0].message
                });
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Sharing link creation requested', { 
                    fileId: value.fileId,
                    type: value.type,
                    scope: value.scope,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method
            const result = await filesModule.createSharingLink(value.fileId, value.type, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('Sharing link creation completed successfully', {
                    fileId: value.fileId,
                    type: value.type,
                    scope: value.scope,
                    linkId: result?.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('Sharing link creation completed with session', {
                    sessionId: req.session.id,
                    fileId: value.fileId,
                    type: value.type,
                    scope: value.scope,
                    linkId: result?.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_sharing_link_create_api_success', executionTime, {
                fileId: value.fileId,
                type: value.type,
                scope: value.scope,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `Sharing link creation error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/sharing/link',
                    fileId: req.body?.fileId,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Sharing link creation failed', {
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('Sharing link creation failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.body?.fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_sharing_link_create_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_SHARING_LINK_CREATE_FAILED',
                error_description: 'Failed to create sharing link'
            });
        }
    },

    /**
     * GET /api/files/sharing/links
     */
    async getSharingLinks(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!req.query.fileId) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'Get sharing links validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/sharing/links',
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('Get sharing links validation failed', {
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('Get sharing links validation failed', {
                        sessionId: req.session.id,
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_SHARING_INVALID_REQUEST',
                    error_description: 'File ID is required'
                });
            }
            
            const fileId = req.query.fileId;
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Get sharing links requested', { 
                    fileId,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method
            const result = await filesModule.getSharingLinks(fileId, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('Get sharing links completed successfully', {
                    fileId: fileId,
                    linksCount: result?.length || 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('Get sharing links completed with session', {
                    sessionId: req.session.id,
                    fileId: fileId,
                    linksCount: result?.length || 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track success metrics
            monitoringService.trackMetric('files_sharing_links_get_api_success', executionTime, {
                fileId: fileId,
                linksCount: result?.length || 0,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'files',
                `Get sharing links error: ${err.message}`,
                'error',
                { 
                    endpoint: '/api/files/sharing/links',
                    fileId: req.query.fileId,
                    error: err.message,
                    stack: err.stack,
                    timestamp: new Date().toISOString()
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Get sharing links failed', {
                    error: err.message,
                    fileId: req.query.fileId,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('Get sharing links failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.query.fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics
            monitoringService.trackMetric('files_sharing_links_get_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ 
                error: 'FILES_SHARING_LINKS_GET_FAILED',
                error_description: 'Failed to retrieve sharing links'
            });
        }
    },

    /**
     * DELETE /api/files/sharing/permission
     */
    async removeSharingPermission(req, res) {
        // Extract user context from auth middleware
        const { userId = null, deviceId = null } = req.user || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            const schema = Joi.object({
                fileId: Joi.string().min(1).max(255).required(),
                permissionId: Joi.string().min(1).max(255).required()
            });
            
            const { error, value } = schema.validate(req.body);
            if (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'files',
                    'Remove sharing permission validation failed',
                    'error',
                    { 
                        endpoint: '/api/files/sharing/permission',
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }
                );
                monitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService.error('Remove sharing permission validation failed', {
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (req.session?.id) {
                    monitoringService.error('Remove sharing permission validation failed', {
                        sessionId: req.session.id,
                        error: error.details[0].message,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                return res.status(400).json({ 
                    error: 'FILES_SHARING_PERMISSION_INVALID_REQUEST',
                    error_description: error.details[0].message
                });
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService.debug('Remove sharing permission requested', { 
                    fileId: value.fileId,
                    permissionId: value.permissionId,
                    sessionId: req.session?.id,
                    userAgent: req.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Call the module method
            const result = await filesModule.removeSharingPermission(value.fileId, value.permissionId, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService.info('Remove sharing permission completed successfully', {
                    fileId: value.fileId,
                    permissionId: value.permissionId,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.info('Remove sharing permission completed with session', {
                    sessionId: req.session.id,
                    fileId: value.fileId,
                    permissionId: value.permissionId,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Log success metrics with user context
            monitoringService.trackMetric('files_remove_sharing_api_success', executionTime, {
                fileId: value.fileId,
                permissionId: value.permissionId,
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.json(result);
        } catch (err) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Create standardized error with user context
            const mcpError = ErrorService.createError(
                ErrorService.CATEGORIES.API, 
                `Files sharing permission removal error: ${err.message}`, 
                ErrorService.SEVERITIES.ERROR, 
                { 
                    fileId: req.body?.fileId,
                    permissionId: req.body?.permissionId,
                    stack: err.stack,
                    timestamp: new Date().toISOString(),
                    userId,
                    deviceId
                }
            );
            monitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService.error('Remove sharing permission failed', {
                    error: err.message,
                    fileId: req.body?.fileId,
                    permissionId: req.body?.permissionId,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (req.session?.id) {
                monitoringService.error('Remove sharing permission failed', {
                    sessionId: req.session.id,
                    error: err.message,
                    fileId: req.body?.fileId,
                    permissionId: req.body?.permissionId,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Track failure metrics with user context
            monitoringService.trackMetric('files_remove_sharing_api_failure', executionTime, {
                errorType: err.code || 'unknown',
                timestamp: new Date().toISOString(),
                userId,
                deviceId
            });
            
            res.status(500).json({ error: 'Internal error', message: err.message });
        }
    }
    };
};
