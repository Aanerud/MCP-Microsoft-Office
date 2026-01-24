/**
 * @fileoverview MCP Files Module - Handles file-related intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system and phase1_architecture.md.
 */

const { normalizeFile } = require('../../graph/normalizers.cjs');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

// Define the capabilities supported by this module
// This array is used by the module registry to find appropriate modules for different intents
const FILES_CAPABILITIES = [
    'listFiles',            // List files and folders in a directory
    'searchFiles',          // Search files by name or content
    'downloadFile',         // Download file content or get metadata with options.metadataOnly
    'uploadFile',           // Upload new files
    'getFileMetadata',      // Get file metadata by ID
    'getFileContent',       // Get file content by ID
    'setFileContent',       // Set file content (mapped to updateFileContent with setContent: true)
    'updateFileContent',    // Update or set file content with options.setContent
    'createSharingLink',    // Create sharing links for files
    'getSharingLinks',      // Get existing sharing links for files
    'removeSharingPermission' // Remove sharing permissions from files
];

// Log module initialization
MonitoringService.info('Files Module initialized', {
    serviceName: 'files-module',
    capabilities: FILES_CAPABILITIES.length, // Dynamic count of capabilities
    timestamp: new Date().toISOString()
}, 'files');

// Create a consolidated FilesModule with only the essential file tools
const FilesModule = {
    /**
     * Helper method to redact sensitive data from objects before logging
     * @param {object} data - The data object to redact
     * @returns {object} Redacted copy of the data
     * @private
     */
    redactSensitiveData(data, visited = new WeakSet()) {
        if (!data || typeof data !== 'object') {
            return data;
        }
        
        // Check for circular references
        if (visited.has(data)) {
            return '[Circular Reference]';
        }
        
        // Add current object to visited set
        visited.add(data);
        
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
                    result[key] = this.redactSensitiveData(result[key], visited);
                }
            }
        }
        
        return result;
    },
    
    id: 'files',
    name: 'OneDrive Files',
    capabilities: FILES_CAPABILITIES,
    
    /**
     * Initializes the files module with dependencies.
     * @param {object} services - { graphService, errorService, monitoringService }
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {object} Initialized module
     */
    init(services, userId, sessionId) {
        const startTime = Date.now();
        
        // Validate that required services are provided
        const requiredServices = ['graphService', 'errorService', 'monitoringService']; 
        
        // Use imported services as fallbacks during initialization
        const errorService = services?.errorService || ErrorService;
        const monitoringService = services?.monitoringService || MonitoringService;

        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService?.debug('Initializing Files Module', {
                    userId: userId ? userId.substring(0, 20) + '...' : 'system',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    requiredServices: requiredServices.length,
                    timestamp: new Date().toISOString()
                }, 'files');
            }

            if (!services) {
                const error = errorService?.createError(
                    'files',
                    'FilesModule init requires a services object',
                    'error',
                    { 
                        userId: userId ? userId.substring(0, 20) + '...' : 'system',
                        sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString() 
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(error);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    monitoringService?.error('Files module initialization failed', {
                        error: 'Missing services object',
                        timestamp: new Date().toISOString()
                    }, 'files', null, userId);
                } else if (sessionId) {
                    monitoringService?.error('Files module initialization failed', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        error: 'Missing services object',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                    
                throw error;
            }

            // Validate required services
            for (const serviceName of requiredServices) {
                if (!services[serviceName]) {
                    const error = errorService?.createError(
                        'files',
                        `FilesModule init failed: Required service '${serviceName}' is missing`,
                        'error',
                        { 
                            missingService: serviceName,
                            userId: userId ? userId.substring(0, 20) + '...' : 'system',
                            sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                            timestamp: new Date().toISOString() 
                        }
                    );
                    
                    // Pattern 3: Infrastructure Error Logging
                    monitoringService?.logError(error);
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        monitoringService?.error('Files module initialization failed', {
                            error: `Missing required service: ${serviceName}`,
                            timestamp: new Date().toISOString()
                        }, 'files', null, userId);
                    } else if (sessionId) {
                        monitoringService?.error('Files module initialization failed', {
                            sessionId: sessionId.substring(0, 8) + '...',
                            error: `Missing required service: ${serviceName}`,
                            timestamp: new Date().toISOString()
                        }, 'files');
                    }
                        
                    throw error;
                }
            }

            this.services = services;
            
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (userId) {
                monitoringService?.info('Files module initialized successfully', {
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (sessionId) {
                monitoringService?.info('Files module initialized successfully', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
                
            return this; // Return the module instance, now containing validated services
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = error.id ? error : errorService.createError(
                'files',
                `Files module initialization failed: ${error.message}`,
                'error',
                {
                    userId: userId ? userId.substring(0, 20) + '...' : 'system',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    executionTimeMs: executionTime,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                monitoringService?.error('Files module initialization failed', {
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, userId);
            } else if (sessionId) {
                monitoringService?.error('Files module initialization failed', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    error: error.message,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            throw mcpError;
        }
    },
    /**
     * Lists files and folders in a directory (defaults to root)
     * @param {string} [parentId] - Parent folder ID (null for root)
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Array<object>>} List of files and folders
     */
    async listFiles(parentId, req, userId, sessionId) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Extract user context from req if not provided
        const contextUserId = userId || req?.user?.userId;
        const contextSessionId = sessionId || req?.session?.id;
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService?.debug('Files listing requested', {
                    parentId: parentId || 'root',
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    userAgent: req?.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Validate that GraphService is available
            if (!graphService || typeof graphService.listFiles !== 'function') {
                const err = errorService.createError(
                    'files',
                    'GraphService.listFiles not implemented',
                    'error',
                    { 
                        parentId: parentId || 'root',
                        userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(err);
                
                // Pattern 4: User Error Tracking
                if (contextUserId) {
                    monitoringService?.error('Files listing failed', {
                        error: 'GraphService.listFiles not implemented',
                        parentId: parentId || 'root',
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.error('Files listing failed', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        error: 'GraphService.listFiles not implemented',
                        parentId: parentId || 'root',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                throw err;
            }
            
            // Call the Graph service with user context
            const result = await graphService.listFiles(parentId, req, contextUserId, contextSessionId);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (contextUserId) {
                monitoringService?.info('Files listing completed successfully', {
                    fileCount: Array.isArray(result) ? result.length : 0,
                    parentId: parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, contextUserId);
            } else if (contextSessionId) {
                monitoringService?.info('Files listing completed successfully', {
                    sessionId: contextSessionId.substring(0, 8) + '...',
                    fileCount: Array.isArray(result) ? result.length : 0,
                    parentId: parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = error.id ? error : errorService.createError(
                'files',
                `Failed to list files: ${error.message}`,
                'error',
                { 
                    parentId: parentId || 'root',
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    executionTimeMs: executionTime,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (contextUserId) {
                monitoringService?.error('Files listing failed', {
                    error: error.message,
                    parentId: parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, contextUserId);
            } else if (contextSessionId) {
                monitoringService?.error('Files listing failed', {
                    sessionId: contextSessionId.substring(0, 8) + '...',
                    error: error.message,
                    parentId: parentId || 'root',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Searches files by name
     * @param {string} query - Search query
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Array<object>>} List of matching files
     */
    async searchFiles(query, req, userId, sessionId) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Extract user context from req if not provided
        const contextUserId = userId || req?.user?.userId;
        const contextSessionId = sessionId || req?.session?.id;
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!query) {
                const err = errorService.createError(
                    'files',
                    'Search query is required for searchFiles',
                    'error',
                    { 
                        userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_query'
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(err);
                
                // Pattern 4: User Error Tracking
                if (contextUserId) {
                    monitoringService?.error('Files search validation failed', {
                        error: 'Search query is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.error('Files search validation failed', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        error: 'Search query is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                throw err;
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService?.debug('Files search requested', {
                    query: query.substring(0, 50) + '...',
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    userAgent: req?.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            // Validate that GraphService is available
            if (!graphService) {
                const err = errorService.createError(
                    'files',
                    'GraphService not available',
                    'error',
                    { 
                        userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(err);
                
                // Pattern 4: User Error Tracking
                if (contextUserId) {
                    monitoringService?.error('Files search failed', {
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.error('Files search failed', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        error: 'GraphService not available',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                throw err;
            }
            
            if (typeof graphService.searchFiles !== 'function') {
                const err = errorService.createError(
                    'files',
                    'GraphService.searchFiles not implemented',
                    'error',
                    { 
                        userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_method'
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(err);
                
                // Pattern 4: User Error Tracking
                if (contextUserId) {
                    monitoringService?.error('Files search failed', {
                        error: 'GraphService.searchFiles not implemented',
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.error('Files search failed', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        error: 'GraphService.searchFiles not implemented',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                throw err;
            }
            
            // Call the Graph service with user context
            const results = await graphService.searchFiles(query, req, contextUserId, contextSessionId);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Pattern 2: User Activity Logs
            if (contextUserId) {
                monitoringService?.info('Files search completed successfully', {
                    resultCount: Array.isArray(results) ? results.length : 0,
                    queryLength: query.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, contextUserId);
            } else if (contextSessionId) {
                monitoringService?.info('Files search completed successfully', {
                    sessionId: contextSessionId.substring(0, 8) + '...',
                    resultCount: Array.isArray(results) ? results.length : 0,
                    queryLength: query.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            return results;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = error.id ? error : errorService.createError(
                'files',
                `Failed to search files: ${error.message}`,
                'error',
                { 
                    query: query ? query.substring(0, 50) + '...' : 'none',
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    executionTimeMs: executionTime,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (contextUserId) {
                monitoringService?.error('Files search failed', {
                    error: error.message,
                    queryLength: query ? query.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, contextUserId);
            } else if (contextSessionId) {
                monitoringService?.error('Files search failed', {
                    sessionId: contextSessionId.substring(0, 8) + '...',
                    error: error.message,
                    queryLength: query ? query.length : 0,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Downloads a file by ID or gets its metadata
     * @param {string} id - File ID
     * @param {object} [options={}] - Options for the download
     * @param {boolean} [options.metadataOnly=false] - If true, returns only file metadata without content
     * @param {object} req - Express request object (optional)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {Promise<Buffer|object>} File content as Buffer or file metadata object if metadataOnly is true
     */
    async downloadFile(id, options = {}, req, userId, sessionId) {
        // Handle case where req is passed as second parameter (backward compatibility)
        if (req === undefined && options && !options.metadataOnly && typeof options === 'object') {
            req = options;
            options = {};
        }
        
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Extract user context from req if not provided
        const contextUserId = userId || req?.user?.userId;
        const contextSessionId = sessionId || req?.session?.id;
        
        // Start tracking execution time
        const startTime = Date.now();
        
        try {
            // Validate input
            if (!id) {
                const err = errorService.createError(
                    'files',
                    'File ID is required for downloadFile',
                    'error',
                    { 
                        userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_file_id'
                    }
                );
                
                // Pattern 3: Infrastructure Error Logging
                monitoringService?.logError(err);
                
                // Pattern 4: User Error Tracking
                if (contextUserId) {
                    monitoringService?.error('File download validation failed', {
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.error('File download validation failed', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        error: 'File ID is required',
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
                
                throw err;
            }
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                monitoringService?.debug('File download requested', {
                    fileId: id.substring(0, 20) + '...',
                    metadataOnly: options.metadataOnly || false,
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    userAgent: req?.get('User-Agent'),
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            let result;
            
            // Determine which operation to perform based on options
            if (options.metadataOnly) {
                // If metadataOnly is true, get file metadata instead of content
                // Validate that GraphService.getFileMetadata is available
                if (!graphService || typeof graphService.getFileMetadata !== 'function') {
                    const err = errorService.createError(
                        'files',
                        'GraphService.getFileMetadata not implemented',
                        'error',
                        { 
                            fileId: id.substring(0, 20) + '...',
                            userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                            sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                            timestamp: new Date().toISOString(),
                            serviceError: 'missing_graph_service'
                        }
                    );
                    
                    // Pattern 3: Infrastructure Error Logging
                    monitoringService?.logError(err);
                    
                    // Pattern 4: User Error Tracking
                    if (contextUserId) {
                        monitoringService?.error('File metadata retrieval failed', {
                            error: 'GraphService.getFileMetadata not implemented',
                            fileId: id.substring(0, 20) + '...',
                            timestamp: new Date().toISOString()
                        }, 'files', null, contextUserId);
                    } else if (contextSessionId) {
                        monitoringService?.error('File metadata retrieval failed', {
                            sessionId: contextSessionId.substring(0, 8) + '...',
                            error: 'GraphService.getFileMetadata not implemented',
                            fileId: id.substring(0, 20) + '...',
                            timestamp: new Date().toISOString()
                        }, 'files');
                    }
                    
                    throw err;
                }
                
                // Call the Graph service to get metadata with user context
                result = await graphService.getFileMetadata(id, req, contextUserId, contextSessionId);
                
                // Calculate execution time
                const executionTime = Date.now() - startTime;
                
                // Pattern 2: User Activity Logs
                if (contextUserId) {
                    monitoringService?.info('File metadata retrieved successfully', {
                        fileId: id.substring(0, 20) + '...',
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.info('File metadata retrieved successfully', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        fileId: id.substring(0, 20) + '...',
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
            } else {
                // Get file content (default behavior)
                // Validate that GraphService.downloadFile is available
                if (!graphService || typeof graphService.downloadFile !== 'function') {
                    const err = errorService.createError(
                        'files',
                        'GraphService.downloadFile not implemented',
                        'error',
                        { 
                            fileId: id.substring(0, 20) + '...',
                            userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                            sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                            timestamp: new Date().toISOString(),
                            serviceError: 'missing_graph_service'
                        }
                    );
                    
                    // Pattern 3: Infrastructure Error Logging
                    monitoringService?.logError(err);
                    
                    // Pattern 4: User Error Tracking
                    if (contextUserId) {
                        monitoringService?.error('File download failed', {
                            error: 'GraphService.downloadFile not implemented',
                            fileId: id.substring(0, 20) + '...',
                            timestamp: new Date().toISOString()
                        }, 'files', null, contextUserId);
                    } else if (contextSessionId) {
                        monitoringService?.error('File download failed', {
                            sessionId: contextSessionId.substring(0, 8) + '...',
                            error: 'GraphService.downloadFile not implemented',
                            fileId: id.substring(0, 20) + '...',
                            timestamp: new Date().toISOString()
                        }, 'files');
                    }
                    
                    throw err;
                }
                
                // Call the Graph service to download file content with user context
                result = await graphService.downloadFile(id, req, contextUserId, contextSessionId);
                
                // Calculate execution time
                const executionTime = Date.now() - startTime;
                
                // Pattern 2: User Activity Logs
                if (contextUserId) {
                    monitoringService?.info('File downloaded successfully', {
                        fileId: id.substring(0, 20) + '...',
                        contentSize: result?.length || 0,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'files', null, contextUserId);
                } else if (contextSessionId) {
                    monitoringService?.info('File downloaded successfully', {
                        sessionId: contextSessionId.substring(0, 8) + '...',
                        fileId: id.substring(0, 20) + '...',
                        contentSize: result?.length || 0,
                        executionTimeMs: executionTime,
                        timestamp: new Date().toISOString()
                    }, 'files');
                }
            }
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = error.id ? error : errorService.createError(
                'files',
                `Failed to ${options.metadataOnly ? 'get file metadata' : 'download file'}: ${error.message}`,
                'error',
                { 
                    fileId: id.substring(0, 20) + '...',
                    metadataOnly: options.metadataOnly || false,
                    userId: contextUserId ? contextUserId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: contextSessionId ? contextSessionId.substring(0, 8) + '...' : 'none',
                    executionTimeMs: executionTime,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            monitoringService?.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (contextUserId) {
                monitoringService?.error(`File ${options.metadataOnly ? 'metadata retrieval' : 'download'} failed`, {
                    error: error.message,
                    fileId: id.substring(0, 20) + '...',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files', null, contextUserId);
            } else if (contextSessionId) {
                monitoringService?.error(`File ${options.metadataOnly ? 'metadata retrieval' : 'download'} failed`, {
                    sessionId: contextSessionId.substring(0, 8) + '...',
                    error: error.message,
                    fileId: id.substring(0, 20) + '...',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'files');
            }
            
            throw mcpError;
        }
    },
    
    /**
     * Uploads a file to root directory
     * @param {string} name - File name
     * @param {Buffer} content - File content
     * @param {object} req - Express request object (optional)
     * @returns {Promise<object>} File metadata
     */
    async uploadFile(name, content, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Log the request with detailed parameters
        monitoringService?.debug('File upload requested', { 
            fileName: name,
            contentSize: content?.length || 0,
            timestamp: new Date().toISOString(),
            source: 'files.uploadFile'
        }, 'files');
        
        try {
            // Validate input
            if (!name) {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.VALIDATION,
                    'File name is required for uploadFile',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_file_name'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            if (!content) {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.VALIDATION,
                    'File content is required for uploadFile',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        fileName: name,
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_file_content'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Validate that GraphService is available
            if (!graphService || typeof graphService.uploadFile !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.uploadFile not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        fileName: name,
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Call the Graph service
            const result = await graphService.uploadFile(name, content, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Log success metrics
            monitoringService?.trackMetric('files_upload_success', executionTime, {
                fileName: name,
                contentSize: content?.length || 0,
                timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Track failure metrics
            monitoringService?.trackMetric('files_upload_failure', executionTime, {
                errorType: error.code || 'unknown',
                fileName: name,
                contentSize: content?.length || 0,
                timestamp: new Date().toISOString()
            });
            
            // Create standardized error if not already one
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to upload file: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { 
                    fileName: name,
                    contentSize: content?.length || 0,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            // Log the error
            monitoringService?.logError(mcpError);
            
            // Rethrow the error for the caller to handle
            throw mcpError;
        }
    },
    
    /**
     * Retrieves metadata for a file by ID
     * @param {string} id - File ID
     * @param {object} req - Express request object (optional)
     * @returns {Promise<object>} File metadata
     */
    async getFileMetadata(id, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Log the request
        monitoringService?.debug('File metadata requested', { 
            fileId: id,
            timestamp: new Date().toISOString(),
            source: 'files.getFileMetadata'
        }, 'files');
        
        try {
            // Validate input
            if (!id) {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.VALIDATION,
                    'File ID is required for getFileMetadata',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_file_id'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Validate that GraphService is available
            if (!graphService || typeof graphService.getFileMetadata !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.getFileMetadata not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        fileId: id,
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Call the Graph service
            const result = await graphService.getFileMetadata(id, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Log success metrics
            monitoringService?.trackMetric('files_metadata_success', executionTime, {
                fileId: id,
                timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Track failure metrics
            monitoringService?.trackMetric('files_metadata_failure', executionTime, {
                errorType: error.code || 'unknown',
                fileId: id,
                timestamp: new Date().toISOString()
            });
            
            // Create standardized error if not already one
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to get file metadata: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { 
                    fileId: id,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            // Log the error
            monitoringService?.logError(mcpError);
            
            // Rethrow the error for the caller to handle
            throw mcpError;
        }
    },
    
    /**
     * Creates a sharing link for a file
     * @param {string} id - File ID
     * @param {string} type - Link type ('view', 'edit', etc.)
     * @param {object} req - Express request object (optional)
     * @returns {Promise<object>} Sharing link
     */
    async createSharingLink(id, type = 'view', req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Log the request
        monitoringService?.debug('Create sharing link requested', { 
            fileId: id,
            linkType: type,
            timestamp: new Date().toISOString(),
            source: 'files.createSharingLink'
        }, 'files');
        
        try {
            // Validate input
            if (!id) {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.VALIDATION,
                    'File ID is required for createSharingLink',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        timestamp: new Date().toISOString(),
                        validationError: 'missing_file_id'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Validate that GraphService is available
            if (!graphService || typeof graphService.createSharingLink !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.createSharingLink not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { 
                        fileId: id,
                        linkType: type,
                        timestamp: new Date().toISOString(),
                        serviceError: 'missing_graph_service'
                    }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            // Call the Graph service
            const result = await graphService.createSharingLink(id, type, req);
            
            // Calculate execution time
            const executionTime = Date.now() - startTime;
            
            // Log success metrics
            monitoringService?.trackMetric('files_create_sharing_link_success', executionTime, {
                fileId: id,
                linkType: type,
                timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            // Calculate execution time even for failures
            const executionTime = Date.now() - startTime;
            
            // Track failure metrics
            monitoringService?.trackMetric('files_create_sharing_link_failure', executionTime, {
                errorType: error.code || 'unknown',
                fileId: id,
                linkType: type,
                timestamp: new Date().toISOString()
            });
            
            // Create standardized error if not already one
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to create sharing link: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { 
                    fileId: id,
                    linkType: type,
                    error: error.toString(),
                    stack: error.stack,
                    timestamp: new Date().toISOString()
                }
            );
            
            // Log the error
            monitoringService?.logError(mcpError);
            
            // Rethrow the error for the caller to handle
            throw mcpError;
        }
    },
    
    /**
     * Gets sharing links for a file
     * @param {string} id - File ID
     * @param {object} req - Express request object (optional)
     * @returns {Promise<Array<object>>} List of sharing links
     */
    async getSharingLinks(id, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Validate input and log request
        if (!id) {
            const err = errorService.createError(
                ErrorService.CATEGORIES.VALIDATION,
                'File ID is required for getSharingLinks',
                ErrorService.SEVERITIES.ERROR,
                { timestamp: new Date().toISOString() }
            );
            monitoringService?.logError(err);
            throw err;
        }
        
        try {
            if (!graphService || typeof graphService.getSharingLinks !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.getSharingLinks not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { fileId: id, timestamp: new Date().toISOString() }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            const result = await graphService.getSharingLinks(id, req);
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_get_sharing_links_success', executionTime, {
                fileId: id,
                linkCount: Array.isArray(result) ? result.length : 0,
                timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_get_sharing_links_failure', executionTime, {
                errorType: error.code || 'unknown',
                fileId: id,
                timestamp: new Date().toISOString()
            });
            
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to get sharing links: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { fileId: id, error: error.toString(), stack: error.stack, timestamp: new Date().toISOString() }
            );
            
            monitoringService?.logError(mcpError);
            throw mcpError;
        }
    },
    
    /**
     * Removes a sharing permission from a file
     * @param {string} fileId - File ID
     * @param {string} permissionId - Permission ID
     * @param {object} req - Express request object (optional)
     * @returns {Promise<object>} Result
     */
    async removeSharingPermission(fileId, permissionId, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        // Start tracking execution time
        const startTime = Date.now();
        
        // Validate input
        if (!fileId || !permissionId) {
            const err = errorService.createError(
                ErrorService.CATEGORIES.VALIDATION,
                'File ID and permission ID are required for removeSharingPermission',
                ErrorService.SEVERITIES.ERROR,
                { fileId, permissionId, timestamp: new Date().toISOString() }
            );
            monitoringService?.logError(err);
            throw err;
        }
        
        try {
            if (!graphService || typeof graphService.removeSharingPermission !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.removeSharingPermission not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { fileId, permissionId, timestamp: new Date().toISOString() }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            const result = await graphService.removeSharingPermission(fileId, permissionId, req);
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_remove_sharing_permission_success', executionTime, {
                fileId, permissionId, timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_remove_sharing_permission_failure', executionTime, {
                errorType: error.code || 'unknown',
                fileId, permissionId, timestamp: new Date().toISOString()
            });
            
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to remove sharing permission: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { fileId, permissionId, error: error.toString(), stack: error.stack, timestamp: new Date().toISOString() }
            );
            
            monitoringService?.logError(mcpError);
            throw mcpError;
        }
    },
    
    /**
     * Gets file content by ID
     * @param {string} id - File ID
     * @param {object} req - Express request object (optional)
     * @returns {Promise<Buffer>} File content
     */
    async getFileContent(id, req) {
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        const startTime = Date.now();
        
        // Validate input
        if (!id) {
            const err = errorService.createError(
                ErrorService.CATEGORIES.VALIDATION,
                'File ID is required for getFileContent',
                ErrorService.SEVERITIES.ERROR,
                { timestamp: new Date().toISOString() }
            );
            monitoringService?.logError(err);
            throw err;
        }
        
        try {
            if (!graphService || typeof graphService.getFileContent !== 'function') {
                const err = errorService.createError(
                    ErrorService.CATEGORIES.SYSTEM,
                    'GraphService.getFileContent not implemented',
                    ErrorService.SEVERITIES.ERROR,
                    { fileId: id, timestamp: new Date().toISOString() }
                );
                monitoringService?.logError(err);
                throw err;
            }
            
            const result = await graphService.getFileContent(id, req);
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_get_content_success', executionTime, {
                fileId: id, contentSize: result?.length || 0, timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            monitoringService?.trackMetric('files_get_content_failure', executionTime, {
                errorType: error.code || 'unknown', fileId: id, timestamp: new Date().toISOString()
            });
            
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to get file content: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { fileId: id, error: error.toString(), stack: error.stack, timestamp: new Date().toISOString() }
            );
            
            monitoringService?.logError(mcpError);
            throw mcpError;
        }
    },
    
    /**
     * Updates or sets file content by ID
     * @param {string} id - File ID
     * @param {Buffer} content - File content
     * @param {object} options - Options for the operation
     * @param {boolean} [options.setContent=false] - If true, treats operation as 'set' rather than 'update'
     * @param {object} req - Express request object (optional)
     * @returns {Promise<object>} Updated file metadata
     */
    async updateFileContent(id, content, options = {}, req) {
        // Handle backward compatibility - if options is the request object, shift parameters
        if (options && !req && (typeof options !== 'object' || options?.headers || options?.session)) {
            req = options;
            options = {};
        }
        
        // Get services with fallbacks
        const { graphService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
        
        const startTime = Date.now();
        const operationType = options.setContent ? 'set' : 'update';
        
        // Validate input
        if (!id) {
            const err = errorService.createError(
                ErrorService.CATEGORIES.VALIDATION,
                `File ID is required for ${operationType}FileContent`,
                ErrorService.SEVERITIES.ERROR,
                { timestamp: new Date().toISOString() }
            );
            monitoringService?.logError(err);
            throw err;
        }
        
        if (!content) {
            const err = errorService.createError(
                ErrorService.CATEGORIES.VALIDATION,
                `File content is required for ${operationType}FileContent`,
                ErrorService.SEVERITIES.ERROR,
                { fileId: id, timestamp: new Date().toISOString() }
            );
            monitoringService?.logError(err);
            throw err;
        }
        
        try {
            // Log operation details
            monitoringService?.debug(`${operationType} file content requested`, {
                fileId: id,
                operationType,
                contentSize: content?.length || 0,
                timestamp: new Date().toISOString()
            });
            
            // Call the consolidated updateFileContent method with options
            const result = await graphService.updateFileContent(id, content, req, options);
            const executionTime = Date.now() - startTime;
            
            // Track success metrics
            monitoringService?.trackMetric(`files_${operationType}_content_success`, executionTime, {
                fileId: id, 
                contentSize: content?.length || 0, 
                timestamp: new Date().toISOString()
            });
            
            return result;
        } catch (error) {
            const executionTime = Date.now() - startTime;
            
            // Track failure metrics
            monitoringService?.trackMetric(`files_${operationType}_content_failure`, executionTime, {
                errorType: error.code || 'unknown', 
                fileId: id, 
                contentSize: content?.length || 0, 
                timestamp: new Date().toISOString(),
                operationType
            });
            
            const mcpError = error.id ? error : errorService.createError(
                ErrorService.CATEGORIES.SYSTEM,
                `Failed to ${operationType} file content: ${error.message}`,
                ErrorService.SEVERITIES.ERROR,
                { 
                    fileId: id, 
                    contentSize: content?.length || 0, 
                    error: error.toString(), 
                    stack: error.stack, 
                    timestamp: new Date().toISOString(),
                    operationType 
                }
            );
            
            monitoringService?.logError(mcpError);
            throw mcpError;
        }
    },
    
    /**
     * Sets file content (wrapper around updateFileContent with setContent option).
     * @param {string} id - File ID
     * @param {string|Buffer} content - File content
     * @param {object} req - Express request object
     * @returns {Promise<object>} Updated file metadata
     */
    async setFileContent(id, content, req) {
        return await this.updateFileContent(id, content, { setContent: true }, req);
    },

/**
 * Handles file-related intents routed to this module.
 * @param {string} intent
 * @param {object} entities
 * @param {object} context
 * @returns {Promise<object>} Normalized response
 */
async handleIntent(intent, entities = {}, context = {}) {
    // Get services with fallbacks
    const { graphService, cacheService, errorService = ErrorService, monitoringService = MonitoringService } = this.services || {};
    
    // Start tracking execution time
    const startTime = Date.now();
    
    // Log the request with detailed parameters
    monitoringService?.debug('Files intent handling requested', { 
        intent,
        entities: this.redactSensitiveData(entities),
        timestamp: new Date().toISOString(),
        source: 'files.handleIntent'
    }, 'files');
    
    try {
        // Validate that GraphService is available
        if (!graphService) {
            const err = errorService.createError(
                'files',
                'GraphService not available',
                'error',
                { 
                    intent,
                    timestamp: new Date().toISOString(),
                    validationError: 'missing_graph_service'
                }
            );
            monitoringService?.logError(err);
            throw err;
        }
            
        let result;
            
        switch(intent) {
            case 'downloadFile': {
                // Handle both file content download and metadata-only requests
                const { fileId, metadataOnly } = entities;
                
                // Create options object for downloadFile
                const options = { metadataOnly: metadataOnly === true };
                
                // Log the specific operation being performed
                monitoringService?.debug(`Handling downloadFile intent with ${options.metadataOnly ? 'metadata only' : 'full content'}`, {
                    fileId,
                    metadataOnly: options.metadataOnly,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                // Call the consolidated downloadFile method
                const file = await this.downloadFile(fileId, options, context.req);
                
                // Return appropriate response based on operation type
                if (options.metadataOnly) {
                    result = { type: 'fileMetadata', file: normalizeFile(file) };
                } else {
                    result = { type: 'fileDownload', fileId, content: file };
                }
                break;
            }
            
            case 'uploadFile': {
                // Handle file upload requests
                const { name, content, parentId } = entities;
                
                // Log upload request details
                monitoringService?.debug('Handling uploadFile intent', {
                    fileName: name,
                    parentId: parentId || 'root',
                    contentSize: content?.length || 0,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                // Call the uploadFile method
                const uploadResult = await this.uploadFile(name, content, context.req);
                
                // Return normalized file metadata
                result = { type: 'fileUploadResult', file: normalizeFile(uploadResult) };
                break;
            }
            
            case 'updateFileContent': {
                // Handle both update and set file content operations
                const { fileId, content, setContent } = entities;
                
                // Create options object for updateFileContent
                const options = { setContent: setContent === true };
                
                // Log the specific operation being performed
                monitoringService?.debug(`Handling ${options.setContent ? 'setFileContent' : 'updateFileContent'} intent`, {
                    fileId,
                    contentSize: content?.length || 0,
                    setContent: options.setContent,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                // Call the consolidated updateFileContent method
                const updateResult = await this.updateFileContent(fileId, content, options, context.req);
                
                // Return normalized file metadata
                result = { 
                    type: options.setContent ? 'setFileContentResult' : 'updateFileContentResult', 
                    file: normalizeFile(updateResult) 
                };
                break;
            }
            
            // For backward compatibility, map deprecated intents to consolidated functions
            case 'getFileMetadata': {
                // Map to downloadFile with metadataOnly option
                const { fileId } = entities;
                
                monitoringService?.debug('Mapping getFileMetadata intent to downloadFile with metadataOnly', {
                    fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const file = await this.downloadFile(fileId, { metadataOnly: true }, context.req);
                result = { type: 'fileMetadata', file: normalizeFile(file) };
                break;
            }
            
            case 'getFileContent': {
                // Map to downloadFile
                const { fileId } = entities;
                
                monitoringService?.debug('Mapping getFileContent intent to downloadFile', {
                    fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const content = await this.downloadFile(fileId, { metadataOnly: false }, context.req);
                result = { type: 'fileContent', fileId, content };
                break;
            }
            
            case 'setFileContent': {
                // Map to updateFileContent with setContent option
                const { fileId, content } = entities;
                
                monitoringService?.debug('Mapping setFileContent intent to updateFileContent with setContent option', {
                    fileId,
                    contentSize: content?.length || 0,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const updateResult = await this.updateFileContent(fileId, content, { setContent: true }, context.req);
                result = { type: 'setFileContentResult', file: normalizeFile(updateResult) };
                break;
            }
            
            case 'listFiles': {
                // Handle file listing requests
                const { parentId } = entities;
                
                monitoringService?.debug('Handling listFiles intent', {
                    parentId: parentId || 'root',
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const files = await this.listFiles(parentId, context.req);
                result = { type: 'fileList', files: files.map(normalizeFile) };
                break;
            }
            
            case 'searchFiles': {
                // Handle file search requests
                const { query } = entities;
                
                monitoringService?.debug('Handling searchFiles intent', {
                    query,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const files = await this.searchFiles(query, context.req);
                result = { type: 'searchResults', query, files: files.map(normalizeFile) };
                break;
            }
            
            case 'createSharingLink': {
                // Handle sharing link creation requests
                const { fileId, type = 'view' } = entities;
                
                monitoringService?.debug('Handling createSharingLink intent', {
                    fileId,
                    linkType: type,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const sharingLink = await this.createSharingLink(fileId, type, context.req);
                result = { type: 'sharingLink', fileId, sharingLink };
                break;
            }
            
            case 'getSharingLinks': {
                // Handle get sharing links requests
                const { fileId } = entities;
                
                monitoringService?.debug('Handling getSharingLinks intent', {
                    fileId,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const sharingLinks = await this.getSharingLinks(fileId, context.req);
                result = { type: 'sharingLinks', fileId, sharingLinks };
                break;
            }
            
            case 'removeSharingPermission': {
                // Handle remove sharing permission requests
                const { fileId, permissionId } = entities;
                
                monitoringService?.debug('Handling removeSharingPermission intent', {
                    fileId,
                    permissionId,
                    timestamp: new Date().toISOString()
                }, 'files');
                
                const removeResult = await this.removeSharingPermission(fileId, permissionId, context.req);
                result = { type: 'permissionRemoved', fileId, permissionId, result: removeResult };
                break;
            }
            default: {
                // Create a standardized error for unsupported intent
                const unsupportedError = errorService.createError(
                    'files',
                    `The files module does not support the intent: ${intent}`,
                    'warn',
                    { 
                        intent, 
                        moduleId: this.id,
                        timestamp: new Date().toISOString()
                    }
                );
                
                monitoringService?.logError(unsupportedError);
                monitoringService?.warn(`Unsupported files intent received: ${intent}`, {
                    intent,
                    timestamp: new Date().toISOString()
                }, 'files');
                    
                // Track metric for unsupported intent
                monitoringService?.trackMetric('files_unsupported_intent', 1, {
                    intent,
                    timestamp: new Date().toISOString()
                });
                
                throw unsupportedError; // Throw error to signal unsupported operation
            }
        }
            
        // Calculate execution time
        const executionTime = Date.now() - startTime;
            
        // Log success metrics
        monitoringService?.trackMetric('files_intent_success', executionTime, {
            intent,
            timestamp: new Date().toISOString()
        });
            
        // Log success with result summary
        monitoringService?.info('Files intent handled successfully', {
            intent,
            executionTimeMs: executionTime,
            timestamp: new Date().toISOString()
        }, 'files');
            
        return result;
    } catch (error) {
        // Calculate execution time even for failures
        const executionTime = Date.now() - startTime;
            
        // Track failure metrics
        monitoringService?.trackMetric('files_intent_failure', executionTime, {
            intent,
            errorType: error.code || 'unknown',
            timestamp: new Date().toISOString()
        });
            
        // Create standardized error if not already one
        const mcpError = error.id ? error : errorService.createError(
            'files',
            `Failed to handle files intent: ${error.message}`,
            'error',
            { 
                intent,
                entities: this.redactSensitiveData(entities),
                error: error.toString(),
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
            
        // Log the error
        monitoringService?.logError(mcpError);
            
        // Rethrow the error for the caller to handle
        throw mcpError;
        }
    }
};

module.exports = FilesModule;
