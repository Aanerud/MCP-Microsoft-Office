/**
 * @fileoverview GraphClientFactory creates authenticated Microsoft Graph API clients.
 * Uses user-provided access tokens from AuthService. No secrets, fully async, testable.
 */

const graphNormalizer = require('./normalizers.cjs');
const msalService = require('../auth/msal-service.cjs');
const cacheService = require('../core/cache-service.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const fetch = require('node-fetch');

// Log service initialization - use optional chaining to prevent errors if service not fully initialized
MonitoringService?.info?.('Graph Client initialized', {
    serviceName: 'graph-client',
    timestamp: new Date().toISOString()
}, 'graph');

/**
 * Creates an authenticated Graph client.
 * @param {Object} req - Express request object (optional)
 * @param {string} userId - User ID for logging context (optional)
 * @param {string} sessionId - Session ID for logging context (optional)
 * @returns {Promise<GraphClient>}
 */
async function createClient(req, userId, sessionId) {
    const startTime = Date.now();
    
    // Extract user context from request if not provided
    const contextUserId = userId || req?.user?.userId || req?.user?.deviceId;
    const contextSessionId = sessionId || req?.session?.id;
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Graph client creation started', {
            method: 'createClient',
            hasRequest: !!req,
            userId: contextUserId,
            sessionId: contextSessionId,
            userAgent: req?.get?.('User-Agent'),
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    try {
        // Determine the type of request
        const isApiCall = req && (
            req.isApiCall || 
            req.path?.startsWith('/v1/') || 
            req.headers?.['x-mcp-internal-call'] === 'true' ||
            (req.user && req.user.deviceId && req.user.deviceId.startsWith('mcp-token-'))
        );
        
        // Pattern 1: Development Debug Logs - Request type determination
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Graph client creation type determined', {
                isApiCall,
                requestType: isApiCall ? 'API call' : 'regular request',
                hasRequest: !!req,
                requestPath: req?.path,
                hasHeaders: !!req?.headers,
                mcpHeader: req?.headers?.['x-mcp-internal-call'],
                userId: contextUserId,
                sessionId: contextSessionId,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        let token;
        
        // For API calls (including internal MCP calls), use the stored token
        if (isApiCall) {
            // Extract user ID from request for MCP calls
            // Always use userId (ms365:email) for token lookup - this matches how external tokens are stored
            // The deviceId (mcp-token-xxx) is for device tracking, NOT for token lookup
            let tokenUserId = req?.user?.userId || null;
            
            // Pattern 1: Development Debug Logs - Token retrieval
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting stored token for API call', {
                    tokenUserId,
                    userId: contextUserId,
                    sessionId: contextSessionId,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            token = await msalService.getMostRecentToken(tokenUserId);
        } else {
            // Normal flow - get token from session
            // Pattern 1: Development Debug Logs - Session token retrieval
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Getting token from session', {
                    userId: contextUserId,
                    sessionId: contextSessionId,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            token = await msalService.getAccessToken(req);
        }
        
        if (!token) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'graph',
                'No access token available for Graph client',
                'error',
                {
                    service: 'graph-client',
                    method: 'createClient',
                    isApiCall,
                    userId: contextUserId,
                    sessionId: contextSessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (contextUserId) {
                MonitoringService.error('Graph client creation failed - no access token', {
                    error: 'No access token available',
                    isApiCall,
                    timestamp: new Date().toISOString()
                }, 'graph', null, contextUserId);
            } else if (contextSessionId) {
                MonitoringService.error('Graph client creation failed - no access token', {
                    sessionId: contextSessionId,
                    error: 'No access token available',
                    isApiCall,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            
            throw mcpError;
        }
        
        const client = new GraphClient(token);
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (contextUserId) {
            MonitoringService.info('Graph client created successfully', {
                isApiCall,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, contextUserId);
        } else if (contextSessionId) {
            MonitoringService.info('Graph client created successfully', {
                sessionId: contextSessionId,
                isApiCall,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        MonitoringService.trackMetric('graph_client_creation_success', executionTime, {
            service: 'graph-client',
            method: 'createClient',
            isApiCall,
            userId: contextUserId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 1: Development Debug Logs - Success
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Graph client created successfully', {
                executionTimeMs: executionTime,
                userId: contextUserId,
                sessionId: contextSessionId,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return client;
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // If it's already an MCP error, just track metrics and rethrow
        if (error.category) {
            MonitoringService.trackMetric('graph_client_creation_failure', executionTime, {
                service: 'graph-client',
                method: 'createClient',
                errorType: error.code || 'auth_error',
                userId: contextUserId,
                timestamp: new Date().toISOString()
            });
            
            // Pattern 4: User Error Tracking for existing MCP errors
            if (contextUserId) {
                MonitoringService.error('Graph client creation failed', {
                    error: error.message,
                    errorCode: error.code,
                    timestamp: new Date().toISOString()
                }, 'graph', null, contextUserId);
            } else if (contextSessionId) {
                MonitoringService.error('Graph client creation failed', {
                    sessionId: contextSessionId,
                    error: error.message,
                    errorCode: error.code,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            
            throw error;
        }
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Graph client creation failed: ${error.message}`,
            'error',
            {
                service: 'graph-client',
                method: 'createClient',
                stack: error.stack,
                userId: contextUserId,
                sessionId: contextSessionId,
                timestamp: new Date().toISOString()
            }
        );
        
        MonitoringService.logError(mcpError);
        MonitoringService.trackMetric('graph_client_creation_failure', executionTime, {
            service: 'graph-client',
            method: 'createClient',
            errorType: error.code || 'unknown',
            userId: contextUserId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (contextUserId) {
            MonitoringService.error('Graph client creation failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'graph', null, contextUserId);
        } else if (contextSessionId) {
            MonitoringService.error('Graph client creation failed', {
                sessionId: contextSessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
}

class GraphClient {
    constructor(token) {
        this.token = token;
    }

    /**
     * Returns a request builder for a Graph API endpoint.
     * @param {string} path - The Graph API path (e.g., '/me')
     * @param {string} userId - User ID for logging context (optional)
     * @param {string} sessionId - Session ID for logging context (optional)
     */
    api(path, userId, sessionId) {
        const self = this;
        let apiVersion = 'v1.0'; // Default version

        const requestBuilder = {
            /**
             * Set the API version (e.g., 'v1.0' or 'beta')
             * @param {string} ver - API version
             * @returns {object} - Returns this for chaining
             */
            version(ver) {
                apiVersion = ver;
                return requestBuilder;
            },
            async get(options = {}) {
                return await _fetchWithRetry(path, self.token, 'GET', null, options, 2, userId, sessionId, apiVersion);
            },
            async post(body, options = {}) {
                return await _fetchWithRetry(path, self.token, 'POST', body, options, 2, userId, sessionId, apiVersion);
            },
            async put(body, options = {}) {
                return await _fetchWithRetry(path, self.token, 'PUT', body, options, 2, userId, sessionId, apiVersion);
            },
            async patch(body, options = {}) {
                return await _fetchWithRetry(path, self.token, 'PATCH', body, options, 2, userId, sessionId, apiVersion);
            },
            async delete(options = {}) {
                return await _fetchWithRetry(path, self.token, 'DELETE', null, options, 2, userId, sessionId, apiVersion);
            }
        };

        return requestBuilder;
    }

    /**
     * Batch multiple Graph API requests.
     * @param {Array<{method: string, url: string, body?: any}>} requests
     * @returns {Promise<Array<any>>}
     */
    async batch(requests) {
        const response = await _fetchWithRetry('/$batch', this.token, 'POST', { requests }, {});
        return (response.responses || []).map(r => r.body);
    }
}

/**
 * Helper to fetch with retry logic and respect for Microsoft Graph rate limiting.
 * Retries on 429 (Too Many Requests) using the retry-after header if present.
 * @param {string} path
 * @param {string} token
 * @param {string} method
 * @param {object|null} body
 * @param {object} options
 * @param {number} retries
 * @param {string} userId - User ID for logging context (optional)
 * @param {string} sessionId - Session ID for logging context (optional)
 * @param {string} apiVersion - API version (e.g., 'v1.0' or 'beta'), defaults to 'v1.0'
 */
async function _fetchWithRetry(path, token, method, body, options, retries = 2, userId, sessionId, apiVersion = 'v1.0') {
    const startTime = Date.now();
    const url = (path.startsWith('http') ? path : `https://graph.microsoft.com/${apiVersion}${path}`);
    const headers = Object.assign({
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
    }, options.headers || {});
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Graph API request started', {
            method,
            path,
            url,
            hasBody: !!body,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    for (let attempt = 0; attempt <= retries; attempt++) {
        const requestStartTime = Date.now();
        
        try {
            const res = await fetch(url, {
                method,
                headers,
                body: body ? JSON.stringify(body) : undefined
            });
            
            const responseTime = Date.now() - requestStartTime;
            const totalTime = Date.now() - startTime;
            
            // Track all Graph API requests
            MonitoringService.trackMetric('graph_api_request', responseTime, {
                method: method,
                endpoint: path,
                statusCode: res.status,
                success: res.ok,
                attempt: attempt + 1,
                userId,
                timestamp: new Date().toISOString()
            });
            
            if (res.ok) {
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Graph API request successful', {
                        method,
                        path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                        status: res.status,
                        responseTimeMs: responseTime,
                        attempt: attempt + 1,
                        timestamp: new Date().toISOString()
                    }, 'graph', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Graph API request successful', {
                        sessionId,
                        method,
                        path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                        status: res.status,
                        responseTimeMs: responseTime,
                        attempt: attempt + 1,
                        timestamp: new Date().toISOString()
                    }, 'graph');
                }
                
                // Pattern 1: Development Debug Logs
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Graph API request successful', {
                        method,
                        path,
                        status: res.status,
                        responseTimeMs: responseTime,
                        totalTimeMs: totalTime,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'graph');
                }
                
                // Handle responses with no content (like DELETE operations)
                if (res.status === 204 || res.headers.get('content-length') === '0') {
                    return { success: true, status: res.status };
                }

                // Check content type to determine how to handle response
                const contentType = res.headers.get('content-type') || '';

                // Check if content type is text-based (should not be returned as binary)
                const isTextContent = contentType.includes('text/') ||
                    contentType.includes('application/json') ||
                    contentType.includes('application/xml');

                // For binary/file content responses (e.g., /content endpoint), return raw buffer
                // But exclude text-based content types that should be returned as strings
                const isBinaryContent = !isTextContent && (
                    path.includes('/content') ||
                    contentType.includes('application/octet-stream') ||
                    contentType.includes('application/pdf') ||
                    contentType.includes('image/') ||
                    contentType.includes('application/vnd.') ||
                    contentType.includes('application/msword') ||
                    contentType.includes('application/zip')
                );

                if (isBinaryContent) {
                    // Pattern 1: Development Debug Logs
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService?.debug('Graph API returning binary content', {
                            method,
                            path,
                            contentType,
                            status: res.status,
                            timestamp: new Date().toISOString()
                        }, 'graph');
                    }

                    // Return raw buffer for binary content
                    const buffer = await res.buffer();
                    return buffer;
                }

                // For text content types (but not JSON), return as text string
                if (isTextContent && !contentType.includes('application/json')) {
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService?.debug('Graph API returning text content', {
                            method,
                            path,
                            contentType,
                            status: res.status,
                            timestamp: new Date().toISOString()
                        }, 'graph');
                    }
                    return await res.text();
                }

                // For JSON responses or unknown content types, read as text first then try to parse
                // (node-fetch only allows reading body once, so we read as text and parse manually)
                const bodyText = await res.text();

                // Try to parse as JSON first
                if (contentType.includes('application/json') || !contentType) {
                    try {
                        return JSON.parse(bodyText);
                    } catch (jsonError) {
                        // Not valid JSON, return as text if we have content
                        if (bodyText && bodyText.length > 0) {
                            MonitoringService?.debug('Graph API response is not JSON, returning as text', {
                                method,
                                path,
                                status: res.status,
                                contentType,
                                bodyLength: bodyText.length,
                                timestamp: new Date().toISOString()
                            }, 'graph');
                            return bodyText;
                        }
                        return { success: true, status: res.status };
                    }
                }

                // For other content types, return as text
                if (bodyText && bodyText.length > 0) {
                    return bodyText;
                }

                return { success: true, status: res.status };
            }
            
            if (res.status === 429) {
                const retryAfter = Number(res.headers.get('retry-after')) || 1;
                
                MonitoringService.trackMetric('graph_api_throttled', responseTime, {
                    method,
                    endpoint: path,
                    attempt: attempt + 1,
                    retryAfter,
                    userId,
                    timestamp: new Date().toISOString()
                });
                
                if (attempt === retries) {
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'graph',
                        `Graph API throttled (429) after ${retries+1} attempts`,
                        'error',
                        {
                            service: 'graph-client',
                            method: '_fetchWithRetry',
                            path,
                            httpMethod: method,
                            attempts: retries + 1,
                            retryAfter,
                            userId,
                            sessionId,
                            timestamp: new Date().toISOString()
                        }
                    );
                    MonitoringService.logError(mcpError);
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('Graph API throttled after max retries', {
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            attempts: retries + 1,
                            retryAfter,
                            timestamp: new Date().toISOString()
                        }, 'graph', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('Graph API throttled after max retries', {
                            sessionId,
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            attempts: retries + 1,
                            retryAfter,
                            timestamp: new Date().toISOString()
                        }, 'graph');
                    }
                    
                    throw mcpError;
                }
                
                // Pattern 1: Development Debug Logs
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Graph API throttled, retrying', {
                        method,
                        path,
                        attempt: attempt + 1,
                        retryAfter,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'graph');
                }
                
                await new Promise(r => setTimeout(r, retryAfter * 1000));
                continue;
            }
            
            // For the last attempt, get detailed error information
            if (attempt === retries) {
                try {
                    const errorData = await res.json();
                    
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'graph',
                        errorData.error ? 
                            `Graph API request failed: ${res.status} - ${errorData.error.code}: ${errorData.error.message}` :
                            `Graph API request failed: ${res.status}`,
                        'error',
                        {
                            service: 'graph-client',
                            method: '_fetchWithRetry',
                            path,
                            httpMethod: method,
                            statusCode: res.status,
                            graphError: errorData.error,
                            attempts: retries + 1,
                            userId,
                            sessionId,
                            timestamp: new Date().toISOString()
                        }
                    );
                    
                    MonitoringService.logError(mcpError);
                    MonitoringService.trackMetric('graph_api_error', totalTime, {
                        method,
                        endpoint: path,
                        statusCode: res.status,
                        errorCode: errorData.error?.code || 'unknown',
                        userId,
                        timestamp: new Date().toISOString()
                    });
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('Graph API request failed', {
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            statusCode: res.status,
                            errorCode: errorData.error?.code || 'unknown',
                            errorMessage: errorData.error?.message || 'Unknown error',
                            timestamp: new Date().toISOString()
                        }, 'graph', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('Graph API request failed', {
                            sessionId,
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            statusCode: res.status,
                            errorCode: errorData.error?.code || 'unknown',
                            errorMessage: errorData.error?.message || 'Unknown error',
                            timestamp: new Date().toISOString()
                        }, 'graph');
                    }
                    
                    throw mcpError;
                } catch (parseError) {
                    // If we can't parse the error response as JSON
                    const errorText = await res.text().catch(() => 'Unable to read error response');
                    
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'graph',
                        `Graph API request failed: ${res.status} - ${errorText.substring(0, 200)}`,
                        'error',
                        {
                            service: 'graph-client',
                            method: '_fetchWithRetry',
                            path,
                            httpMethod: method,
                            statusCode: res.status,
                            errorText: errorText.substring(0, 200),
                            parseError: parseError.message,
                            attempts: retries + 1,
                            userId,
                            sessionId,
                            timestamp: new Date().toISOString()
                        }
                    );
                    
                    MonitoringService.logError(mcpError);
                    MonitoringService.trackMetric('graph_api_error', totalTime, {
                        method,
                        endpoint: path,
                        statusCode: res.status,
                        errorType: 'parse_error',
                        userId,
                        timestamp: new Date().toISOString()
                    });
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('Graph API request failed with parse error', {
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            statusCode: res.status,
                            errorType: 'parse_error',
                            timestamp: new Date().toISOString()
                        }, 'graph', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('Graph API request failed with parse error', {
                            sessionId,
                            method,
                            path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                            statusCode: res.status,
                            errorType: 'parse_error',
                            timestamp: new Date().toISOString()
                        }, 'graph');
                    }
                    
                    throw mcpError;
                }
            }
        } catch (error) {
            const responseTime = Date.now() - requestStartTime;
            
            // If it's already an MCP error, just rethrow
            if (error.category) {
                throw error;
            }
            
            // Network or other fetch errors
            if (attempt === retries) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'graph',
                    `Graph API network error: ${error.message}`,
                    'error',
                    {
                        service: 'graph-client',
                        method: '_fetchWithRetry',
                        path,
                        httpMethod: method,
                        networkError: error.message,
                        attempts: retries + 1,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }
                );
                
                MonitoringService.logError(mcpError);
                MonitoringService.trackMetric('graph_api_network_error', responseTime, {
                    method,
                    endpoint: path,
                    errorType: 'network',
                    userId,
                    timestamp: new Date().toISOString()
                });
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Graph API network error', {
                        method,
                        path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                        errorType: 'network',
                        errorMessage: error.message,
                        timestamp: new Date().toISOString()
                    }, 'graph', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Graph API network error', {
                        sessionId,
                        method,
                        path: path.substring(0, 50) + (path.length > 50 ? '...' : ''),
                        errorType: 'network',
                        errorMessage: error.message,
                        timestamp: new Date().toISOString()
                    }, 'graph');
                }
                
                throw mcpError;
            }
        }
    }
}

/**
 * Enhanced batch method: retries only failed requests (429) after retry-after delay.
 * @param {Array} requests - Array of batch requests
 * @param {number} retries - Number of retry attempts (default: 2)
 * @param {string} userId - User ID for logging context (optional)
 * @param {string} sessionId - Session ID for logging context (optional)
 */
GraphClient.prototype.batch = async function(requests, retries = 2, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Graph API batch request started', {
            method: 'batch',
            requestCount: requests.length,
            userId,
            sessionId,
            timestamp: new Date().toISOString()
        }, 'graph');
    }
    
    try {
        // BUG-3: Use unique IDs for batch requests to handle out-of-order responses
        // Graph API batch responses may not maintain request order
        let pending = requests.map((req, i) => ({
            ...req,
            id: req.id || `req_${i}`,  // Ensure each request has a unique ID
            _idx: i
        }));

        // Create a map of ID to original index for O(1) lookup
        const idToIndex = new Map(pending.map(req => [req.id, req._idx]));

        let results = new Array(requests.length);
        let attempts = 0;

        while (pending.length && attempts <= retries) {
            const batchStartTime = Date.now();
            const response = await _fetchWithRetry('/$batch', this.token, 'POST', { requests: pending }, {}, 2, userId, sessionId);
            const batchTime = Date.now() - batchStartTime;

            MonitoringService.trackMetric('graph_api_batch_request', batchTime, {
                requestCount: pending.length,
                attempt: attempts + 1,
                userId,
                timestamp: new Date().toISOString()
            });

            const retryRequests = [];
            let maxRetryAfter = 0;
            let successCount = 0;
            let throttledCount = 0;

            // BUG-3: Match responses by ID instead of array index
            const pendingById = new Map(pending.map(req => [req.id, req]));
            (response.responses || []).forEach((r) => {
                const responseId = r.id;
                const pendingReq = pendingById.get(responseId);

                if (!pendingReq) {
                    MonitoringService.warn('Batch response ID not found in pending requests', {
                        responseId,
                        pendingIds: Array.from(pendingById.keys()),
                        timestamp: new Date().toISOString()
                    }, 'graph');
                    return;
                }

                const origIdx = idToIndex.get(responseId);
                if (r.status === 429) {
                    const retryAfter = Number((r.headers && r.headers['retry-after']) || 1);
                    maxRetryAfter = Math.max(maxRetryAfter, retryAfter);
                    retryRequests.push(pendingReq);
                    throttledCount++;
                } else {
                    results[origIdx] = r.body;
                    successCount++;
                }
            });
            
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Graph API batch attempt completed', {
                    attempt: attempts + 1,
                    successCount,
                    throttledCount,
                    pendingCount: retryRequests.length,
                    maxRetryAfter,
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            
            if (!retryRequests.length) break;
            
            if (attempts === retries) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'graph',
                    'Graph API batch throttled (429) after max retries',
                    'error',
                    {
                        service: 'graph-client',
                        method: 'batch',
                        totalRequests: requests.length,
                        throttledRequests: retryRequests.length,
                        attempts: retries + 1,
                        maxRetryAfter,
                        userId,
                        sessionId,
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Graph API batch throttled after max retries', {
                        totalRequests: requests.length,
                        throttledRequests: retryRequests.length,
                        attempts: retries + 1,
                        maxRetryAfter,
                        timestamp: new Date().toISOString()
                    }, 'graph', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Graph API batch throttled after max retries', {
                        sessionId,
                        totalRequests: requests.length,
                        throttledRequests: retryRequests.length,
                        attempts: retries + 1,
                        maxRetryAfter,
                        timestamp: new Date().toISOString()
                    }, 'graph');
                }
                
                throw mcpError;
            }
            
            await new Promise(r => setTimeout(r, maxRetryAfter * 1000));
            pending = retryRequests;
            attempts++;
        }
        
        const executionTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Graph API batch completed successfully', {
                totalRequests: requests.length,
                totalAttempts: attempts + 1,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Graph API batch completed successfully', {
                sessionId,
                totalRequests: requests.length,
                totalAttempts: attempts + 1,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        MonitoringService.trackMetric('graph_api_batch_success', executionTime, {
            service: 'graph-client',
            method: 'batch',
            totalRequests: requests.length,
            totalAttempts: attempts + 1,
            userId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Graph API batch completed successfully', {
                totalRequests: requests.length,
                totalAttempts: attempts + 1,
                executionTimeMs: executionTime,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        return results;
    } catch (error) {
        const executionTime = Date.now() - startTime;
        
        // If it's already an MCP error, just track metrics and rethrow
        if (error.category) {
            MonitoringService.trackMetric('graph_api_batch_failure', executionTime, {
                service: 'graph-client',
                method: 'batch',
                totalRequests: requests.length,
                errorType: error.code || 'api_error',
                userId,
                timestamp: new Date().toISOString()
            });
            
            // Pattern 4: User Error Tracking for existing MCP errors
            if (userId) {
                MonitoringService.error('Graph API batch operation failed', {
                    totalRequests: requests.length,
                    errorType: error.code || 'api_error',
                    errorMessage: error.message,
                    timestamp: new Date().toISOString()
                }, 'graph', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Graph API batch operation failed', {
                    sessionId,
                    totalRequests: requests.length,
                    errorType: error.code || 'api_error',
                    errorMessage: error.message,
                    timestamp: new Date().toISOString()
                }, 'graph');
            }
            
            throw error;
        }
        
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'graph',
            `Graph API batch operation failed: ${error.message}`,
            'error',
            {
                service: 'graph-client',
                method: 'batch',
                totalRequests: requests.length,
                stack: error.stack,
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        
        MonitoringService.logError(mcpError);
        MonitoringService.trackMetric('graph_api_batch_failure', executionTime, {
            service: 'graph-client',
            method: 'batch',
            totalRequests: requests.length,
            errorType: error.code || 'unknown',
            userId,
            timestamp: new Date().toISOString()
        });
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Graph API batch operation failed', {
                totalRequests: requests.length,
                errorMessage: error.message,
                timestamp: new Date().toISOString()
            }, 'graph', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Graph API batch operation failed', {
                sessionId,
                totalRequests: requests.length,
                errorMessage: error.message,
                timestamp: new Date().toISOString()
            }, 'graph');
        }
        
        throw mcpError;
    }
};

module.exports = { createClient, GraphClient };
