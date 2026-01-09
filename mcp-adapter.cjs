/**
 * @fileoverview Microsoft 365 MCP Adapter
 *
 * This adapter implements the Model Context Protocol (MCP) to enable Claude
 * to access Microsoft 365 data (mail, calendar, files) through the MCP server.
 * 
 * It follows the official MCP specification at https://modelcontextprotocol.io/
 * 
 * This adapter uses a proxy approach to communicate with backend modules without
 * directly importing them, providing true isolation between the MCP protocol
 * handling and the backend implementation.
 */

const readline = require('readline');
const fs = require('fs');
const path = require('path');
const os = require('os');

// HTTP client for API communication
const http = require('http');
const https = require('https');
const { URLSearchParams } = require('url');
const Joi = require('joi'); // For parameter validation

// CRITICAL: Set MCP mode flag and silence console BEFORE importing any services
// This prevents service initialization logs from polluting stdout
process.env.MCP_SILENT_MODE = 'true';
console.log = () => {};
console.debug = () => {};
console.info = () => {};
console.warn = () => {};
console.error = () => {};

// Import services AFTER silencing console
const monitoringService = require('./src/core/monitoring-service.cjs');
const errorService = require('./src/core/error-service.cjs');
const createToolsService = require('./src/core/tools-service.cjs');

// Remote server configuration
const MCP_SERVER_URL = process.env.MCP_SERVER_URL || 'http://localhost:3000';
const API_HOST = process.env.API_HOST || 'localhost';
const API_PORT = process.env.API_PORT || 3000;
const API_BASE_PATH = process.env.API_BASE_PATH || '/api';
const API_TIMEOUT = process.env.API_TIMEOUT || 30000; // 30 seconds

// Simple bearer token mode (if MCP_BEARER_TOKEN is provided)
const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN;
const SIMPLE_MODE = !!MCP_BEARER_TOKEN;

// OAuth 2.0 and device credential configuration
const OAUTH_DISCOVERY_PATH = '/.well-known/oauth-protected-resource';
const DEVICE_REGISTRATION_PATH = '/api/auth/device/register';
const DEVICE_TOKEN_PATH = '/api/auth/device/token';

// Device credential storage
let deviceCredentials = {
    deviceId: SIMPLE_MODE ? 'simple-bearer-token' : null,
    deviceSecret: null,
    accessToken: SIMPLE_MODE ? MCP_BEARER_TOKEN : null,
    refreshToken: null,
    tokenExpiry: SIMPLE_MODE ? Date.now() + (24 * 60 * 60 * 1000) : null, // 24 hours from now
    serverUrl: MCP_SERVER_URL
};

// OAuth 2.0 discovery cache
let oauthDiscovery = {
    discovered: false,
    authorizationEndpoint: null,
    tokenEndpoint: null,
    deviceAuthorizationEndpoint: null,
    lastDiscovery: null
};

// MCP adapters should not attempt to write to files directly as they may not have permissions
// All file logging should be handled by the backend via the monitoring service
// Only use stderr for direct logging from the adapter

// Monkey-patch console methods to prevent JSON-RPC stream pollution
// Store original console methods
const originalConsole = {
    log: console.log,
    debug: console.debug,
    info: console.info,
    warn: console.warn,
    error: console.error
};

// Essential logging categories - only log what's needed for debugging
const ESSENTIAL_LOG_CATEGORIES = {
    MCP_REQUEST: true,    // Incoming MCP requests
    MCP_RESPONSE: true,   // Outgoing MCP responses 
    API_CALL: true,       // API calls to backend
    API_RESPONSE: true,   // API responses from backend
    ERROR: true,          // Errors
    STARTUP: true         // Startup/initialization
};

// Helper: Completely silent for MCP protocol compliance
function writeToLogFile(level, ...args) {
    // MCP requires absolute silence - no logging whatsoever
    // All logging is completely disabled
}

// Console methods already silenced above to prevent service import pollution

// Adapter state
let adapterState = {
    initialized: false,
    lastActivity: 0,
    backendAvailable: false
};

// Initialize tools service with a stub module registry containing the module definitions
// Capabilities aligned with scopeToTools in routes.cjs for permission filtering
const stubModuleRegistry = {
    getAllModules: () => [
        { id: 'search', name: 'search', capabilities: ['search'] },
        { id: 'mail', name: 'mail', capabilities: ['getInbox', 'readMail', 'readMailDetails', 'sendEmail', 'sendMail', 'flagEmail', 'flagMail', 'getEmailDetails', 'markAsRead', 'markEmailRead', 'getMailAttachments', 'addMailAttachment', 'removeMailAttachment'] },
        { id: 'calendar', name: 'calendar', capabilities: ['getEvents', 'getCalendars', 'getRooms', 'getAvailability', 'createEvent', 'updateEvent', 'cancelEvent', 'acceptEvent', 'tentativelyAcceptEvent', 'declineEvent', 'findMeetingTimes', 'addAttachment', 'removeAttachment'] },
        { id: 'files', name: 'files', capabilities: ['listFiles', 'downloadFile', 'uploadFile', 'getFileMetadata', 'getFileContent', 'setFileContent', 'updateFileContent', 'createSharingLink', 'getSharingLinks', 'removeSharingPermission'] },
        { id: 'people', name: 'people', capabilities: ['findPeople', 'getRelevantPeople', 'getPersonById'] },
        { id: 'teams', name: 'teams', capabilities: ['listChats', 'getChatMessages', 'sendChatMessage', 'listJoinedTeams', 'listTeamChannels', 'getChannelMessages', 'sendChannelMessage', 'replyToMessage', 'createOnlineMeeting', 'getOnlineMeeting', 'getMeetingByJoinUrl', 'listOnlineMeetings', 'getMeetingTranscripts', 'getMeetingTranscriptContent'] },
        { id: 'todo', name: 'todo', capabilities: ['listTaskLists', 'getTaskList', 'createTaskList', 'updateTaskList', 'deleteTaskList', 'listTasks', 'getTask', 'createTask', 'updateTask', 'deleteTask', 'completeTask'] },
        { id: 'contacts', name: 'contacts', capabilities: ['listContacts', 'getContact', 'createContact', 'updateContact', 'deleteContact', 'searchContacts'] },
        { id: 'groups', name: 'groups', capabilities: ['listGroups', 'getGroup', 'listGroupMembers', 'listMyGroups'] },
        { id: 'query', name: 'query', capabilities: ['query'] }
    ],
    getModule: (moduleName) => {
        const modules = {
            'search': { id: 'search', capabilities: ['search'] },
            'mail': { id: 'mail', capabilities: ['getInbox', 'readMail', 'readMailDetails', 'sendEmail', 'sendMail', 'flagEmail', 'flagMail', 'getEmailDetails', 'markAsRead', 'markEmailRead', 'getMailAttachments', 'addMailAttachment', 'removeMailAttachment'] },
            'calendar': { id: 'calendar', capabilities: ['getEvents', 'getCalendars', 'getRooms', 'getAvailability', 'createEvent', 'updateEvent', 'cancelEvent', 'acceptEvent', 'tentativelyAcceptEvent', 'declineEvent', 'findMeetingTimes', 'addAttachment', 'removeAttachment'] },
            'files': { id: 'files', capabilities: ['listFiles', 'downloadFile', 'uploadFile', 'getFileMetadata', 'getFileContent', 'setFileContent', 'updateFileContent', 'createSharingLink', 'getSharingLinks', 'removeSharingPermission'] },
            'people': { id: 'people', capabilities: ['findPeople', 'getRelevantPeople', 'getPersonById'] },
            'teams': { id: 'teams', capabilities: ['listChats', 'getChatMessages', 'sendChatMessage', 'listJoinedTeams', 'listTeamChannels', 'getChannelMessages', 'sendChannelMessage', 'replyToMessage', 'createOnlineMeeting', 'getOnlineMeeting', 'getMeetingByJoinUrl', 'listOnlineMeetings', 'getMeetingTranscripts', 'getMeetingTranscriptContent'] },
            'todo': { id: 'todo', capabilities: ['listTaskLists', 'getTaskList', 'createTaskList', 'updateTaskList', 'deleteTaskList', 'listTasks', 'getTask', 'createTask', 'updateTask', 'deleteTask', 'completeTask'] },
            'contacts': { id: 'contacts', capabilities: ['listContacts', 'getContact', 'createContact', 'updateContact', 'deleteContact', 'searchContacts'] },
            'groups': { id: 'groups', capabilities: ['listGroups', 'getGroup', 'listGroupMembers', 'listMyGroups'] },
            'query': { id: 'query', capabilities: ['query'] }
        };
        return modules[moduleName] || null;
    }
};
const toolsService = createToolsService({
    moduleRegistry: stubModuleRegistry,
    logger: {
        debug: (...args) => {/* Silent */},
        info: (...args) => {/* Silent */},
        warn: (...args) => {/* Silent */},
        error: (...args) => {/* Silent */}
    }
});

// Permission-based tool filtering cache
let permissionsCache = null;
let permissionsCacheTime = 0;
const PERMISSIONS_CACHE_TTL = 5 * 60 * 1000; // 5 minutes

/**
 * Fetch available permissions from server and cache them
 * @returns {Promise<Set<string>>} Set of available tool names
 */
async function getAvailableTools() {
    const now = Date.now();

    // Return cached permissions if still valid
    if (permissionsCache && (now - permissionsCacheTime) < PERMISSIONS_CACHE_TTL) {
        return permissionsCache;
    }

    try {
        const result = await callApi('GET', '/v1/permissions');
        if (result && result.availableTools) {
            permissionsCache = new Set(result.availableTools);
            permissionsCacheTime = now;
            logDebug(`[Permissions] Loaded ${result.availableTools.length} available tools from ${result.scopeCount} scopes`);
            return permissionsCache;
        }
    } catch (error) {
        logDebug(`[Permissions] Failed to fetch permissions: ${error.message}`);
    }

    // Return null if permissions check failed - don't filter tools
    return null;
}

/**
 * Filter tools list based on available permissions
 * @param {Array} tools - Full list of tools
 * @returns {Promise<Array>} Filtered list of tools user has permission for
 */
async function filterToolsByPermissions(tools) {
    const availableTools = await getAvailableTools();

    // If permissions check failed, return all tools (fail-open)
    if (!availableTools) {
        return tools;
    }

    // Filter tools based on available permissions
    const filtered = tools.filter(tool => availableTools.has(tool.name));
    logDebug(`[Permissions] Filtered tools: ${filtered.length}/${tools.length} available`);
    return filtered;
}

/**
 * Call the API server
 * @param {string} method - HTTP method (GET, POST, PUT, PATCH, DELETE)
 * @param {string} path - API path (e.g., '/v1/mail')
 * @param {object} data - Request data for POST requests
 * @param {number} _retryCount - Current retry count (internal use)
 * @returns {Promise<object>} - API response
 */
// TODO: [callApi] Implement circuit breaker pattern (e.g., using 'opossum') (HIGH)
// TODO: [callApi] Support PATCH/PUT; add retry + circuit breaker (HIGH) - Partially Implemented (PUT/PATCH support + Basic Retry added)
// DONE: [callApi] Refine retry logic with exponential backoff and jitter

/**
 * Makes an HTTP request to the API server with exponential backoff retry logic
 * @param {string} method - HTTP method (GET, POST, PUT, PATCH, DELETE)
 * @param {string} path - API path to call
 * @param {Object} data - Optional data to send in the request body
 * @param {number} _retryCount - Current retry count (internal use)
 * @returns {Promise<Object>} - Parsed response data
 */
async function callApi(method, path, data = null, _retryCount = 0) {
    // Prevent infinite authentication retry loops
    const MAX_AUTH_RETRIES = 1;
    
    // Ensure we have a valid token before making API calls (except for health checks)
    if (path !== '/health' && !_retryCount) {
        try {
            const hasValidToken = await ensureValidToken();
            if (!hasValidToken) {
                // Try to authenticate if we don't have a valid token
                const authSuccess = await authenticateDevice();
                if (!authSuccess) {
                    // Continue without authentication for non-critical calls
                    process.stderr.write('⚠️  Warning: No valid authentication token available.\n');
                }
            }
        } catch (authError) {
            // Continue without authentication for non-critical calls
            process.stderr.write('⚠️  Warning: Token validation failed, continuing without authentication.\n');
        }
    }

    // Retry configuration
    const MAX_RETRIES = 3;
    const BASE_DELAY_MS = 100;
    const MAX_DELAY_MS = 3000;
    const JITTER_FACTOR = 0.25; // 25% random jitter to prevent thundering herd
    
    // Track attempts and errors
    let lastError = null;
    let attemptMetrics = [];
    
    // Start time for performance tracking
    const startTime = Date.now();

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
            return await new Promise((resolve, reject) => {
                // Parse the server URL to support both localhost and remote servers
                const serverUrl = new URL(MCP_SERVER_URL);
                const isHttps = serverUrl.protocol === 'https:';
                const httpModule = isHttps ? https : http;
                
                const options = {
                    hostname: serverUrl.hostname,
                    port: serverUrl.port || (isHttps ? 443 : 80),
                    path: `${API_BASE_PATH}${path}`,
                    method: method,
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json',
                        'X-MCP-Internal-Call': 'true'
                    },
                    timeout: API_TIMEOUT
                };

                console.error(`[DEBUG] callApi - method: ${method}, input path: ${path}, API_BASE_PATH: ${API_BASE_PATH}, final path: ${options.path}`);

                // Add authorization header if we have an access token
                if (deviceCredentials.accessToken) {
                    options.headers['Authorization'] = `Bearer ${deviceCredentials.accessToken}`;
                }

                const requestBody = data ? JSON.stringify(data) : null;

                if (requestBody && ['POST', 'PUT', 'PATCH'].includes(method.toUpperCase())) {
                    options.headers['Content-Length'] = Buffer.byteLength(requestBody);
                }

                const req = httpModule.request(options, (res) => {
                    let responseData = '';
                    res.on('data', (chunk) => { responseData += chunk; });
                    res.on('end', async () => {
                        try {
                            // Handle 401 Unauthorized responses with automatic re-authentication
                            if (res.statusCode === 401 && _retryCount < MAX_AUTH_RETRIES) {
                                try {
                                    // Clear current token and attempt re-authentication
                                    deviceCredentials.accessToken = null;
                                    deviceCredentials.tokenExpiry = null;
                                    
                                    // Try to re-authenticate
                                    const authSuccess = await authenticateDevice();
                                    if (authSuccess) {
                                        // Retry the original request with new token
                                        try {
                                            const retryResult = await callApi(method, path, data, _retryCount + 1);
                                            return resolve(retryResult);
                                        } catch (retryError) {
                                            return reject(retryError);
                                        }
                                    } else {
                                        const authErr = new Error('Authentication failed after 401 response');
                                        authErr.statusCode = 401;
                                        return reject(authErr);
                                    }
                                } catch (authError) {
                                    const err = new Error(`Authentication error after 401: ${authError.message}`);
                                    err.statusCode = 401;
                                    return reject(err);
                                }
                            }

                            if (!responseData) {
                                if (res.statusCode >= 200 && res.statusCode < 300) {
                                    return resolve({});
                                } else {
                                    // Treat 5xx errors as potentially retryable
                                    const err = new Error(`API error: ${res.statusCode} with empty response`);
                                    if (res.statusCode >= 500) err.isRetryable = true;
                                    return reject(err);
                                }
                            }

                            const parsedData = JSON.parse(responseData);
                            if (res.statusCode >= 400) {
                                const err = new Error(`API error: ${res.statusCode} ${parsedData.error || responseData}`);
                                // Treat 5xx errors as potentially retryable
                                if (res.statusCode >= 500) err.isRetryable = true;
                                reject(err);
                            } else {
                                resolve(parsedData);
                            }
                        } catch (parseError) {
                            reject(new Error(`Failed to parse API response: ${parseError.message}. Raw: ${responseData.slice(0, 100)}...`));
                        }
                    });
                });

                req.on('error', (error) => {
                    // Network errors are retryable
                    error.isRetryable = true;
                    reject(new Error(`API request network error: ${error.message}`));
                });

                req.on('timeout', () => {
                    req.destroy();
                    const timeoutError = new Error(`API request timed out after ${API_TIMEOUT}ms`);
                    timeoutError.isRetryable = true; // Timeouts are retryable
                    reject(timeoutError);
                });

                if (requestBody && ['POST', 'PUT', 'PATCH'].includes(method.toUpperCase())) {
                    req.write(requestBody);
                }
                req.end();
            });
        } catch (error) {
            // Record attempt metrics for logging
            const attemptDuration = Date.now() - startTime;
            attemptMetrics.push({
                attempt,
                duration: attemptDuration,
                error: error.message,
                statusCode: error.statusCode || 'unknown'
            });
            
            // Save the error for potential retries or final throw
            lastError = error;
            
            // Only retry if the error is marked as retryable and we haven't exceeded max retries
            if (error.isRetryable && attempt < MAX_RETRIES) {
                // Calculate exponential backoff with jitter
                // Formula: min(maxDelay, baseDelay * 2^attempt) * (1 ± jitterFactor)
                const exponentialDelay = Math.min(
                    MAX_DELAY_MS,
                    BASE_DELAY_MS * Math.pow(2, attempt - 1)
                );
                
                // Add jitter to prevent thundering herd problem
                const jitter = 1 - JITTER_FACTOR + (Math.random() * JITTER_FACTOR * 2);
                const delay = Math.floor(exponentialDelay * jitter);
                
                // Wait for the calculated delay before retrying
                await new Promise(resolve => setTimeout(resolve, delay));
            } else {
                // Throw the last encountered error
                throw lastError;
            }
        }
    }
    // This line should technically not be reached if MAX_RETRIES >= 1
    throw lastError || new Error(`callApi failed after ${MAX_RETRIES} retries: ${path}`);
}

/**
 * Discover OAuth 2.0 endpoints from the server
 * @returns {Promise<boolean>} - True if discovery was successful
 */
async function discoverOAuthEndpoints() {
    try {
        // Check if discovery is still valid (cache for 1 hour)
        const now = Date.now();
        if (oauthDiscovery.discovered && oauthDiscovery.lastDiscovery && 
            (now - oauthDiscovery.lastDiscovery) < 3600000) {
            return true;
        }

        // Parse the server URL for discovery
        const serverUrl = new URL(MCP_SERVER_URL);
        const isHttps = serverUrl.protocol === 'https:';
        const httpModule = isHttps ? https : http;
        
        const discoveryOptions = {
            hostname: serverUrl.hostname,
            port: serverUrl.port || (isHttps ? 443 : 80),
            path: OAUTH_DISCOVERY_PATH,
            method: 'GET',
            headers: {
                'Accept': 'application/json'
            },
            timeout: 10000 // 10 second timeout for discovery
        };

        const discoveryData = await new Promise((resolve, reject) => {
            const req = httpModule.request(discoveryOptions, (res) => {
                let responseData = '';
                res.on('data', (chunk) => { responseData += chunk; });
                res.on('end', () => {
                    try {
                        if (res.statusCode >= 200 && res.statusCode < 300) {
                            const parsedData = JSON.parse(responseData);
                            resolve(parsedData);
                        } else {
                            reject(new Error(`Discovery failed: ${res.statusCode}`));
                        }
                    } catch (parseError) {
                        reject(new Error(`Failed to parse discovery response: ${parseError.message}`));
                    }
                });
            });

            req.on('error', (error) => {
                reject(new Error(`Discovery request error: ${error.message}`));
            });

            req.on('timeout', () => {
                req.destroy();
                reject(new Error('Discovery request timed out'));
            });

            req.end();
        });

        // Update discovery cache
        oauthDiscovery.discovered = true;
        oauthDiscovery.authorizationEndpoint = discoveryData.authorization_endpoint;
        oauthDiscovery.tokenEndpoint = discoveryData.token_endpoint;
        oauthDiscovery.deviceAuthorizationEndpoint = discoveryData.device_authorization_endpoint;
        oauthDiscovery.lastDiscovery = now;

        return true;
    } catch (error) {
        // Discovery failed - reset cache
        oauthDiscovery.discovered = false;
        oauthDiscovery.lastDiscovery = null;
        return false;
    }
}

// Helper: Completely silent for MCP protocol compliance
function logDebug(message, ...args) {
    // MCP requires absolute silence - no logging at all
    // All logging is completely disabled
}

// Helper: Completely silent for MCP protocol compliance
function logToFile(prefix, method, data) {
    // MCP requires absolute silence - no logging at all
    // All logging is completely disabled
}

// Helper: send JSON-RPC response
function sendResponse(id, result, error) {
    try {
        // Build JSON-RPC response
        const response = {
            jsonrpc: '2.0',
            id
        };
        
        if (error) {
            response.error = error;
        } else {
            response.result = result || null;
        }
        
        // Convert to JSON
        const jsonResponse = JSON.stringify(response);
        
        // Send to stdout (the only place we should write to stdout)
        // CRITICAL: Use process.stdout.write directly instead of console.log
        // to avoid any potential interference from console monkey-patching
        process.stdout.write(jsonResponse + '\n');
    } catch (err) {
        // Silent - no error logging for MCP
        
        // Try to send a simplified error response
        try {
            const fallbackResponse = {
                jsonrpc: '2.0',
                id,
                error: {
                    code: -32603,
                    message: 'Internal error while sending response'
                }
            };
            // Use process.stdout.write directly
            process.stdout.write(JSON.stringify(fallbackResponse) + '\n');
        } catch (fallbackErr) {
            // Last resort: write directly to stdout
            process.stdout.write(JSON.stringify({
                jsonrpc: '2.0',
                id,
                error: { code: -32603, message: 'Critical error' }
            }) + '\n');
        }
    }
}

// Helper: Initialize the adapter and check backend availability
let healthCheckInterval = null; // Keep track of the interval timer
const HEALTH_CHECK_INTERVAL_MS = 60 * 1000; // Check every 60 seconds

async function initializeAdapter() {
    if (adapterState.initialized) {
        return adapterState.backendAvailable;
    }

    try {
        // Check if backend API is available by calling the health endpoint
        const healthResponse = await callApi('GET', '/health');

        if (!healthResponse || healthResponse.status !== 'ok') {
            adapterState.backendAvailable = false;
        } else {
            adapterState.backendAvailable = true;
            adapterState.initialized = true;

            // Perform OAuth 2.0 discovery
            await discoverOAuthEndpoints();

            // Initialize token management system
            initializeTokenManagement();

            // Perform device authentication
            try {
                const authSuccess = await authenticateDevice();
                if (authSuccess) {
                    // Start token refresh scheduler after successful authentication
                    startTokenRefreshScheduler();
                    process.stderr.write('✅ Authentication completed and token management started\n');
                } else {
                    process.stderr.write('⚠️  Warning: Authentication failed. Some features may not be available.\n');
                }
            } catch (authError) {
                process.stderr.write('⚠️  Warning: Authentication error occurred. Continuing without authentication.\n');
                // Continue adapter initialization even if authentication fails
            }

            // Start periodic health checks only if initial check was successful
            if (!healthCheckInterval) {
                healthCheckInterval = setInterval(async () => {
                    try {
                        const healthResponse = await callApi('GET', '/health');
                        if (healthResponse && healthResponse.status === 'ok') {
                            if (!adapterState.backendAvailable) {
                                adapterState.backendAvailable = true;
                                // TODO: Reset consecutive failure count here
                            }
                        } else {
                            throw new Error('Backend health check failed or returned invalid status');
                        }
                    } catch (error) {
                        if (adapterState.backendAvailable) {
                            adapterState.backendAvailable = false;
                            // TODO: Increment consecutive failure count here
                            // TODO: Implement fail-fast: if consecutive failures > threshold, clearInterval(healthCheckInterval) and log permanent failure.
                        }
                        // else: It was already unavailable, do nothing extra
                    }
                }, HEALTH_CHECK_INTERVAL_MS);

                // TODO: Ensure healthCheckInterval is cleared on adapter shutdown/process exit.
            }
        }
        return adapterState.backendAvailable;
    } catch (error) {
        adapterState.initialized = true;
        adapterState.backendAvailable = false;
        return false;
    }
}

/**
 * Check if the backend API is available and if tools/modules are loaded.
 * Calls the `/v1/tools` endpoint.
 */
async function checkModuleAccess() {
    try {
        if (!adapterState.initialized) {
            await initializeAdapter();
        }

        if (!adapterState.backendAvailable) {
            return false;
        }

        // Check if tools are accessible via the API
        const toolsResponse = await callApi('GET', '/v1/tools');

        // Check for a successful response and a non-empty tools array
        if (toolsResponse && Array.isArray(toolsResponse.tools) && toolsResponse.tools.length > 0) {
            return true;
        } else {
            return false;
        }
    } catch (error) {
        return false;
    }
}

// Helper: Execute a module method with error handling via API
async function executeModuleMethod(moduleName, methodName, params = {}) {
    try {
        // Make sure adapter is initialized
        if (!adapterState.initialized) {
            await initializeAdapter();
        }

        // Update last activity timestamp
        adapterState.lastActivity = Date.now();

        // --- Parameter Transformation ---
        // Instead of validating parameters, we'll transform them as needed
        // to match the format expected by the backend API
        
        
        // Transform parameters based on the module and method
        // This ensures compatibility with the backend API expectations
        let transformedParams = { ...params }; // Start with a copy of the original params
        // Map module.method to API endpoints
        let apiPath = '';
        let apiMethod = 'GET';
        let apiData = null;

        // Determine API path and method based on module.method
        const moduleMethod = `${moduleName}.${methodName}`;
        
        // Log the module method being executed for debugging
        
        
        // Helper functions for parameter transformation
        
        // Helper function to format meeting duration in ISO 8601 format
        const formatMeetingDuration = (duration) => {
            // If already in ISO 8601 format (starts with PT), return as is
            if (typeof duration === 'string' && duration.startsWith('PT')) {
                return duration;
            }
            
            // If it's a string but not in ISO format, try to parse it as minutes
            if (typeof duration === 'string') {
                // Check if it's a number in string format (e.g., '60')
                const minutes = parseInt(duration, 10);
                if (!isNaN(minutes)) {
                    return `PT${minutes}M`;
                }
                
                // Check if it's in hours format (e.g., '1h' or '1 hour')
                if (duration.toLowerCase().includes('h')) {
                    const hours = parseInt(duration, 10);
                    if (!isNaN(hours)) {
                        return `PT${hours}H`;
                    }
                }
                
                // Default to 30 minutes if we can't parse it
                return 'PT30M';
            }
            
            // If it's a number, assume it's minutes
            if (typeof duration === 'number') {
                return `PT${duration}M`;
            }
            
            // Default to 30 minutes
            return 'PT30M';
        };
        
        // Helper function to transform attendees from string or array to proper format
        const transformAttendees = (attendees) => {
            if (!attendees) return undefined;
            
            // If attendees is a string (comma-separated), convert to array
            if (typeof attendees === 'string') {
                return attendees.split(',').map(email => email.trim());
            }
            
            // If already an array, return as is
            return attendees;
        };
        
        // Helper function to transform date/time to proper format
        const transformDateTime = (dateTime, timeZone = 'UTC') => {
            if (!dateTime) return undefined;
            
            // If already an object with dateTime, return as is
            if (typeof dateTime === 'object' && dateTime.dateTime) {
                return dateTime;
            }
            
            // If a string, convert to object format
            if (typeof dateTime === 'string') {
                return {
                    dateTime: dateTime,
                    timeZone: timeZone
                };
            }
            
            return dateTime;
        };
        
        // Helper function to safely handle null/undefined values in objects
        const safeObjectAssign = (target, source) => {
            if (!source) return target;
            
            // Only copy properties that are not null or undefined
            Object.keys(source).forEach(key => {
                if (source[key] !== null && source[key] !== undefined) {
                    target[key] = source[key];
                }
            });
            
            return target;
        };
        
        switch (moduleMethod) {
            // Query module endpoints
            case 'query.processQuery':
            case 'query.getQuery': // Add alias for backward compatibility
                apiPath = '/v1/query';
                apiMethod = 'POST';
                // Ensure query parameter is properly named for the API
                apiData = { 
                    query: transformedParams.query,
                    context: transformedParams.context
                };
                break;
            
            // Email details endpoint
            case 'mail.getEmailDetails':
            case 'mail.readMailDetails':
            case 'outlook mail.readMailDetails':
                if (!transformedParams.id) {
                    const errorMessage = 'Email ID is required for getting email details. Please provide an ID parameter with the email ID.';
                    
                    throw new Error(errorMessage);
                }
                // Properly format the path with the ID parameter
                apiPath = `/v1/mail/${transformedParams.id}`;
                apiMethod = 'GET';
                
                // Add query parameters for additional options
                const mailDetailsParams = [];
                
                // Option to include body content
                if (transformedParams.includeBody !== undefined) {
                    mailDetailsParams.push(`includeBody=${transformedParams.includeBody}`);
                }
                
                // Option to include attachments info
                if (transformedParams.includeAttachments !== undefined) {
                    mailDetailsParams.push(`includeAttachments=${transformedParams.includeAttachments}`);
                }
                
                // Add query parameters to the API path
                if (mailDetailsParams.length > 0) {
                    apiPath += `?${mailDetailsParams.join('&')}`;
                }
                
                // Log the API call for debugging
                
                break;
            case 'mail.getInbox':
            case 'mail.readMail':
            case 'outlook mail.readMail':
                apiPath = '/v1/mail';
                apiMethod = 'GET';
                
                // Add query parameters for filtering emails
                const mailParams = [];
                
                // Option to filter by folder
                if (transformedParams.folder) {
                    mailParams.push(`folder=${encodeURIComponent(transformedParams.folder)}`);
                }
                
                // Option to limit the number of emails
                if (transformedParams.top || transformedParams.limit) {
                    mailParams.push(`$top=${transformedParams.top || transformedParams.limit || 25}`);
                }
                
                // Option to filter by read status
                if (transformedParams.isRead !== undefined) {
                    mailParams.push(`isRead=${transformedParams.isRead}`);
                }
                
                // Option to filter by date range
                if (transformedParams.since) {
                    mailParams.push(`since=${encodeURIComponent(transformedParams.since)}`);
                }
                
                // Add query parameters to the API path
                if (mailParams.length > 0) {
                    apiPath += `?${mailParams.join('&')}`;
                }
                
                // Log the API call for debugging
                
                break;
                
            case 'mail.sendEmail':
            case 'mail.sendMail':
            case 'mail.send':
                const toolDefinition = toolsService.getToolByName('mail.sendEmail');
                if (toolDefinition) {
                    apiPath = toolDefinition.endpoint;
                    apiMethod = toolDefinition.method;
                } else {
                    apiPath = '/v1/mail/send';
                    apiMethod = 'POST';
                }
                // Basic structuring - assumes controller expects these directly
                // Infer contentType
                const looksLikeHtml = /<[a-z][\s\S]*>/i.test(params.body || '');
                apiData = {
                    to: params.to,
                    cc: params.cc, // Pass CC if provided
                    bcc: params.bcc, // Pass BCC if provided
                    subject: params.subject,
                    body: params.body,
                    contentType: looksLikeHtml ? 'HTML' : 'Text', // Set content type
                    attachments: params.attachments // Pass attachments if provided
                };
                
                // Log if attachments are being sent
                if (params.attachments && Array.isArray(params.attachments) && params.attachments.length > 0) {
                    
                }
                
                break;
            case 'mail.searchMail':
            case 'mail.searchEmails':
            case 'mail.search': // Added this alias to handle Claude's tool call
                apiPath = '/v1/mail/search';
                apiMethod = 'GET';
                // Map 'query' parameter to 'q' as expected by the controller
                if (params.query && !params.q) {
                    params.q = params.query;
                    // Remove the original query param to avoid confusion
                    delete params.query;
                }
                break;
            case 'mail.flagMail':
            case 'mail.flagEmail':
                if (!transformedParams.id) {
                    throw new Error('Email ID is required for flagging');
                }
                apiPath = '/v1/mail/flag';
                apiMethod = 'POST';
                // Include email ID and flag value in the request body
                apiData = {
                    id: transformedParams.id,
                    flag: transformedParams.flag !== false // Default to true if not explicitly set to false
                };
                // Log the API call for debugging
                
                break;
            case 'mail.markEmailRead':
            case 'mail.markAsRead':
            case 'outlook mail.markEmailRead':
                if (!transformedParams.id) {
                    const errorMessage = 'Email ID is required for marking as read/unread. Please provide an ID parameter with the email ID.';
                    
                    throw new Error(errorMessage);
                }
                // Properly format the path with the ID parameter
                apiPath = `/v1/mail/${transformedParams.id}/read`;
                apiMethod = 'PATCH';
                // Include isRead in the request body
                apiData = {
                    isRead: transformedParams.isRead !== false // Default to true if not explicitly set to false
                };
                // Log the API call for debugging
                
                break;
            case 'mail.getAttachments':
            case 'mail.getMailAttachments':
            case 'outlook mail.getMailAttachments':
                if (!transformedParams.id) {
                    const errorMessage = 'Email ID is required for getting attachments. Please provide an ID parameter with the email ID.';
                    
                    throw new Error(errorMessage);
                }
                
                // Use the correct endpoint path that matches the route in routes.cjs
                apiPath = `/v1/mail/attachments`;
                apiMethod = 'GET';
                
                // Add the email ID as a query parameter - the controller expects 'id' parameter
                const attachmentParams = [`id=${encodeURIComponent(transformedParams.id)}`];
                
                // Also transform the parameter for the handleIntent method which expects 'mailId'
                transformedParams.mailId = transformedParams.id;
                
                // Option to filter by attachment ID
                if (transformedParams.attachmentId) {
                    attachmentParams.push(`attachmentId=${transformedParams.attachmentId}`);
                }
                
                // Option to include content
                if (transformedParams.includeContent !== undefined) {
                    attachmentParams.push(`includeContent=${transformedParams.includeContent}`);
                }
                
                // Add query parameters to the API path
                if (attachmentParams.length > 0) {
                    apiPath += `?${attachmentParams.join('&')}`;
                }
                
                // Log the API call for debugging
                
                break;

            case 'mail.addMailAttachment':
            case 'outlook mail.addMailAttachment':
                if (!transformedParams.id) {
                    const errorMessage = 'Email ID is required for adding attachment. Please provide an ID parameter with the email ID.';
                    throw new Error(errorMessage);
                }
                
                apiPath = `/v1/mail/${transformedParams.id}/attachments`;
                apiMethod = 'POST';
                
                // The API expects attachment fields directly in the request body, not nested in an attachment object
                // Keep the individual fields as they are - the controller expects them directly
                apiData = {
                    name: transformedParams.name,
                    contentBytes: transformedParams.contentBytes,
                    contentType: transformedParams.contentType || 'application/octet-stream',
                    isInline: transformedParams.isInline || false
                };
                break;

            case 'mail.removeMailAttachment':
            case 'outlook mail.removeMailAttachment':
                if (!transformedParams.id) {
                    const errorMessage = 'Email ID is required for removing attachment. Please provide an ID parameter with the email ID.';
                    throw new Error(errorMessage);
                }
                if (!transformedParams.attachmentId) {
                    const errorMessage = 'Attachment ID is required for removing attachment. Please provide an attachmentId parameter.';
                    throw new Error(errorMessage);
                }
                
                apiPath = `/v1/mail/${transformedParams.id}/attachments/${transformedParams.attachmentId}`;
                apiMethod = 'DELETE';
                break;

            // Calendar module endpoints
            case 'calendar.getEvents':
            case 'calendar.getCalendar':
                apiPath = '/v1/calendar';
                apiMethod = 'GET';
                break;
            case 'calendar.createEvent':
            case 'calendar.create':
                apiPath = '/v1/calendar/events';
                apiMethod = 'POST';
                // Map of Microsoft time zones to IANA time zones
                // The server will use this mapping to convert between formats as needed
                const TIMEZONE_MAPPING = {
                    'Pacific Standard Time': 'America/Los_Angeles',
                    'Eastern Standard Time': 'America/New_York',
                    'Central Standard Time': 'America/Chicago',
                    'Mountain Standard Time': 'America/Denver',
                    'Alaskan Standard Time': 'America/Anchorage',
                    'Hawaiian Standard Time': 'Pacific/Honolulu',
                    'W. Europe Standard Time': 'Europe/Berlin',
                    'GMT Standard Time': 'Europe/London',
                    'Romance Standard Time': 'Europe/Paris',
                    'E. Europe Standard Time': 'Europe/Bucharest',
                    'Singapore Standard Time': 'Asia/Singapore',
                    'Tokyo Standard Time': 'Asia/Tokyo',
                    'China Standard Time': 'Asia/Shanghai',
                    'India Standard Time': 'Asia/Kolkata',
                    'UTC': 'UTC'
                };
                
                // For createEvent, we want to preserve the original Microsoft time zone format
                // because the Graph API expects this format in the Prefer header
                // The server's calendar-service.cjs will handle the mapping to IANA format
                
                // Make copies of start/end to avoid mutating the original
                const start = { ...transformedParams.start };
                const end = { ...transformedParams.end };
                
                // Silent - no timezone logging for MCP protocol compliance
                const timeZone = transformedParams.timeZone;
                
                // Transform parameters to match the controller's expectations
                // Normalize and filter input fields for createEvent
                const eventData = {};
                if (transformedParams.subject) eventData.subject = transformedParams.subject;
                if (start && typeof start === 'object') eventData.start = transformDateTime(start, timeZone);
                if (end && typeof end === 'object') eventData.end = transformDateTime(end, timeZone);
                if (transformedParams.isOnlineMeeting !== undefined) eventData.isOnlineMeeting = transformedParams.isOnlineMeeting;
                // Normalize location
                if (transformedParams.location && typeof transformedParams.location === 'object' && transformedParams.location.displayName) {
                    eventData.location = { displayName: transformedParams.location.displayName };
                }
                // Normalize body
                if (transformedParams.body && typeof transformedParams.body === 'object' && transformedParams.body.content) {
                    eventData.body = {
                        content: transformedParams.body.content,
                        contentType: transformedParams.body.contentType || 'text'
                    };
                }
                // Normalize attendees
                if (Array.isArray(transformedParams.attendees)) {
                    const formattedAttendees = [];
                    for (const attendee of transformedParams.attendees) {
                        if (attendee && attendee.emailAddress && attendee.emailAddress.address) {
                            formattedAttendees.push(attendee);
                        }
                    }
                    if (formattedAttendees.length > 0) eventData.attendees = formattedAttendees;
                }
                // Remove any null/undefined fields
                apiData = Object.fromEntries(Object.entries(eventData).filter(([_, v]) => v !== null && v !== undefined));
                
                // Special handling for attendees based on format
                if (transformedParams.attendees) {
                    // Handle different potential formats of attendees
                    if (Array.isArray(transformedParams.attendees)) {
                        const formattedAttendees = [];
                        
                        for (const attendee of transformedParams.attendees) {
                            // Case 1: It's already in the right format with emailAddress.address
                            if (attendee && attendee.emailAddress && attendee.emailAddress.address) {
                                formattedAttendees.push(attendee);
                            }
                            // Case 2: It's an object with email property
                            else if (attendee && attendee.email) {
                                formattedAttendees.push({
                                    emailAddress: { address: attendee.email },
                                    type: attendee.type || 'required'
                                });
                            }
                            // Case 3: It's a simple string (email)
                            else if (typeof attendee === 'string') {
                                formattedAttendees.push({
                                    emailAddress: { address: attendee },
                                    type: 'required'
                                });
                            }
                        }
                        
                        if (formattedAttendees.length > 0) {
                            apiData.attendees = formattedAttendees;
                        }
                    }
                }
                // Log the API call for debugging
                
                break;
            case 'calendar.events': // Handle the events endpoint for fetching events
                apiPath = '/v1/calendar';  // Changed from /v1/calendar/events to /v1/calendar
                apiMethod = 'GET';
                // For GET requests, parameters go in the query string
                apiData = null;
                apiParams = {
                    limit: transformedParams.limit || 20,
                    filter: transformedParams.filter,
                    timeMin: transformedParams.timeMin,
                    timeMax: transformedParams.timeMax
                };
                
                // Log the API call for debugging
                
                break;
            case 'calendar.updateEvent':
            case 'calendar.update':
                // Ensure we have a valid event ID
                if (!transformedParams.id) {
                    throw new Error('Event ID is required for calendar update');
                }
                apiPath = `/v1/calendar/events/${transformedParams.id}`;
                apiMethod = 'PUT';
                
                try {
                    // Create an object with only the provided parameters
                    // This avoids sending undefined values that could overwrite existing data
                    const updateData = {};
                    const defaultTimeZone = transformedParams.timeZone || 'UTC';
                    
                    // Only add properties that are explicitly provided and not null
                    if (transformedParams.subject !== undefined && transformedParams.subject !== null) {
                        updateData.subject = transformedParams.subject;
                    }
                    
                    // Handle start time if provided
                    if (transformedParams.start !== undefined && transformedParams.start !== null) {
                        updateData.start = transformDateTime(transformedParams.start, defaultTimeZone);
                    }
                    
                    // Handle end time if provided
                    if (transformedParams.end !== undefined && transformedParams.end !== null) {
                        updateData.end = transformDateTime(transformedParams.end, defaultTimeZone);
                    }
                    
                    // Handle attendees if provided
                    if (transformedParams.attendees !== undefined && transformedParams.attendees !== null) {
                        const attendees = transformAttendees(transformedParams.attendees);
                        if (attendees && attendees.length > 0) {
                            updateData.attendees = attendees;
                        }
                    }
                    
                    // Handle location if provided
                    if (transformedParams.location !== undefined && transformedParams.location !== null) {
                        // For location, we need to handle both string and object formats
                        if (typeof transformedParams.location === 'string') {
                            updateData.location = { displayName: transformedParams.location };
                        } else {
                            updateData.location = transformedParams.location;
                        }
                    }
                    
                    // Handle body content if provided
                    if (transformedParams.body !== undefined && transformedParams.body !== null) {
                        // For body, we need to handle both string and object formats
                        if (typeof transformedParams.body === 'string') {
                            // Determine if the body looks like HTML
                            const looksLikeHtml = /<[a-z][\s\S]*>/i.test(transformedParams.body);
                            updateData.body = {
                                content: transformedParams.body,
                                contentType: looksLikeHtml ? 'html' : 'text'
                            };
                        } else {
                            updateData.body = transformedParams.body;
                        }
                    }
                    
                    // Handle online meeting flag if provided
                    if (transformedParams.isOnlineMeeting !== undefined) {
                        updateData.isOnlineMeeting = transformedParams.isOnlineMeeting;
                    }
                    
                    // Normalize and filter input fields for updateEvent
                    // Only send fields that are present and valid, never null
                    const safeUpdateData = {};
                    if (updateData.subject) safeUpdateData.subject = updateData.subject;
                    if (updateData.start && typeof updateData.start === 'object') safeUpdateData.start = updateData.start;
                    if (updateData.end && typeof updateData.end === 'object') safeUpdateData.end = updateData.end;
                    if (updateData.isOnlineMeeting !== undefined) safeUpdateData.isOnlineMeeting = updateData.isOnlineMeeting;
                    if (updateData.isAllDay !== undefined) safeUpdateData.isAllDay = updateData.isAllDay;
                    // Normalize location
                    if (updateData.location && typeof updateData.location === 'object' && updateData.location.displayName) {
                        safeUpdateData.location = { displayName: updateData.location.displayName };
                    }
                    // Normalize body
                    if (updateData.body && typeof updateData.body === 'object' && updateData.body.content) {
                        safeUpdateData.body = {
                            content: updateData.body.content,
                            contentType: updateData.body.contentType || 'text'
                        };
                    }
                    // Normalize attendees
                    if (Array.isArray(updateData.attendees)) {
                        const formattedAttendees = [];
                        for (const attendee of updateData.attendees) {
                            if (attendee && attendee.emailAddress && attendee.emailAddress.address) {
                                formattedAttendees.push(attendee);
                            }
                        }
                        if (formattedAttendees.length > 0) safeUpdateData.attendees = formattedAttendees;
                    }
                    // Remove any null/undefined fields
                    apiData = Object.fromEntries(Object.entries(safeUpdateData).filter(([_, value]) => value !== null && value !== undefined));
                } catch (error) {
                    
                    throw new Error(`Failed to transform updateEvent parameters: ${error.message}`);
                }
                
                // Log the update data for debugging
                
                break;
            case 'calendar.getAvailability':
            case 'calendar.availability':
                apiPath = '/v1/calendar/availability';
                apiMethod = 'POST';
                
                // ENHANCED LOGGING: Log the raw input parameters for debugging
                
                
                // Use the tools service for parameter transformation, following separation of concerns
                try {
                    // Get the transformed parameters from the tools service
                    const { params: transformedAvailabilityParams } = toolsService.transformToolParameters('calendar.getAvailability', transformedParams);
                    
                    // Use the transformed parameters
                    apiData = transformedAvailabilityParams;
                    
                    // ENHANCED LOGGING: Log the transformed API data for debugging
                    
                } catch (error) {
                    // Log the error and rethrow
                    
                    throw error;
                }
                
                // Log the transformed data for debugging
                
                break;
            case 'calendar.acceptEvent':
                const acceptEventId = params.eventId || params.id;
                console.error(`[DEBUG] acceptEvent - params:`, JSON.stringify(params, null, 2));
                console.error(`[DEBUG] acceptEvent - acceptEventId:`, acceptEventId);
                if (!acceptEventId) {
                    throw new Error('Event ID is required for accepting an event');
                }
                
                apiPath = `/v1/calendar/events/${acceptEventId}/accept`;
                apiMethod = 'POST';
                console.error(`[DEBUG] acceptEvent - apiPath:`, apiPath);
                
                // Optional comment for the response
                if (params.comment) {
                    apiData = { comment: params.comment };
                } else {
                    apiData = {};
                }
                
                // Log the API call for debugging
                
                break;
            case 'calendar.tentativelyAcceptEvent':
                const tentativeEventId = params.eventId || params.id;
                console.error(`[DEBUG] tentativelyAcceptEvent - params:`, JSON.stringify(params, null, 2));
                console.error(`[DEBUG] tentativelyAcceptEvent - tentativeEventId:`, tentativeEventId);
                if (!tentativeEventId) {
                    throw new Error('Event ID is required for tentatively accepting an event')
                }
                
                apiPath = `/v1/calendar/events/${tentativeEventId}/tentativelyAccept`;
                apiMethod = 'POST';
                console.error(`[DEBUG] tentativelyAcceptEvent - apiPath:`, apiPath);
                
                // Optional comment for the response
                if (params.comment) {
                    apiData = { comment: params.comment };
                } else {
                    apiData = {};
                }
                
                // Log the API call for debugging
                
                break;
            case 'calendar.declineEvent':
                const declineEventId = params.eventId || params.id;
                if (!declineEventId) {
                    throw new Error('Event ID is required for declining an event')
                }
                
                apiPath = `/v1/calendar/events/${declineEventId}/decline`;
                apiMethod = 'POST';
                
                // Optional comment for the response
                if (params.comment) {
                    apiData = { comment: params.comment };
                } else {
                    apiData = {};
                }
                
                // Log the API call for debugging
                
                break;
            case 'calendar.cancelEvent':
                // Handle both eventId and id parameter names
                if (!transformedParams.eventId && !transformedParams.id) {
                    throw new Error('Event ID is required for canceling an event')
                }
                
                const cancelEventId = transformedParams.eventId || transformedParams.id;
                apiPath = `/v1/calendar/events/${cancelEventId}/cancel`;
                apiMethod = 'POST';
                
                // Optional comment for the cancellation
                if (transformedParams.comment) {
                    apiData = { comment: transformedParams.comment };
                } else {
                    apiData = {};
                }
                
                // Log the cancel request
                
                break;
            case 'calendar.findMeetingTimes':
            case 'calendar.findtimes': // Add this case to handle the findtimes endpoint
                apiPath = '/v1/calendar/findMeetingTimes';
                apiMethod = 'POST';
                
                try {
                    // Transform parameters to match what the calendar service expects
                    const defaultTimeZone = transformedParams.timeZone || 'UTC';
                    
                    // Extract time constraints from different possible input formats
                    let timeConstraints = transformedParams.timeConstraints;
                    let timeSlots = [];
                    
                    // Extract meetingDuration if it's incorrectly nested in timeConstraints
                    // This handles the case where the client sends meetingDuration inside timeConstraints
                    // but the Graph API expects it at the top level
                    if (timeConstraints && timeConstraints.meetingDuration) {
                        transformedParams.meetingDuration = timeConstraints.meetingDuration;
                        // Remove it from timeConstraints to avoid confusion
                        delete timeConstraints.meetingDuration;
                    }
                    
                    // Handle different formats of time constraints
                    if (timeConstraints && timeConstraints.timeSlots && Array.isArray(timeConstraints.timeSlots)) {
                        // Format 1: Explicit timeSlots array in timeConstraints
                        timeSlots = timeConstraints.timeSlots.map(slot => {
                            return {
                                start: transformDateTime(slot.start, slot.start?.timeZone || defaultTimeZone),
                                end: transformDateTime(slot.end, slot.end?.timeZone || defaultTimeZone)
                            };
                        });
                    } else if (timeConstraints && (timeConstraints.start || timeConstraints.startTime)) {
                        // Format 2: Single time range in timeConstraints object
                        timeSlots = [{
                            start: transformDateTime(timeConstraints.start || timeConstraints.startTime, timeConstraints.timeZone || defaultTimeZone),
                            end: transformDateTime(timeConstraints.end || timeConstraints.endTime, timeConstraints.timeZone || defaultTimeZone)
                        }];
                    } else if (transformedParams.startTime || transformedParams.start) {
                        // Format 3: Direct start/end parameters at the top level
                        timeSlots = [{
                            start: transformDateTime(transformedParams.startTime || transformedParams.start, defaultTimeZone),
                            end: transformDateTime(transformedParams.endTime || transformedParams.end, defaultTimeZone)
                        }];
                    } else {
                        // Format 4: Default to current time + 7 days if no time constraints provided
                        const now = new Date();
                        const oneWeekLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
                        
                        timeSlots = [{
                            start: {
                                dateTime: now.toISOString(),
                                timeZone: defaultTimeZone
                            },
                            end: {
                                dateTime: oneWeekLater.toISOString(),
                                timeZone: defaultTimeZone
                            }
                        }];
                    }
                    
                    // Ensure we have at least one time slot
                    if (timeSlots.length === 0) {
                        throw new Error('No valid time slots could be determined from the provided parameters');
                    }
                    
                    // Validate that each time slot has start and end
                    timeSlots.forEach((slot, index) => {
                        if (!slot.start || !slot.end) {
                            throw new Error(`Time slot at index ${index} is missing start or end time`);
                        }
                    });
                    
                    // Create the API data with proper structure following Microsoft Graph API requirements
                    apiData = {
                        // Transform attendees to array format - ensure it's never null/undefined
                        attendees: (transformAttendees(transformedParams.attendees) || []).map(attendee => {
                            // Ensure each attendee has the correct format
                            if (typeof attendee === 'string') {
                                return {
                                    emailAddress: { address: attendee },
                                    type: 'required'
                                };
                            } else if (typeof attendee === 'object') {
                                return {
                                    emailAddress: { address: attendee.email || attendee.address },
                                    type: attendee.type || 'required'
                                };
                            }
                        }),
                        
                        // Time constraints with properly formatted timeSlots array (capital 'S') for controller validation
                        timeConstraint: {
                            activityDomain: timeConstraints?.activityDomain || transformedParams.activityDomain || 'work',
                            // Only use timeSlots (capital 'S') for controller validation
                            timeSlots: timeSlots.map(slot => ({
                                start: {
                                    // Format date to match Microsoft Graph API expectations (without milliseconds and Z)
                                    dateTime: slot.start.dateTime.replace(/\.\d{3}Z$/, ''),
                                    timeZone: slot.start.timeZone || 'UTC'
                                },
                                end: {
                                    // Format date to match Microsoft Graph API expectations (without milliseconds and Z)
                                    dateTime: slot.end.dateTime.replace(/\.\d{3}Z$/, ''),
                                    timeZone: slot.end.timeZone || 'UTC'
                                }
                            }))
                        },
                        
                        // Duration in ISO 8601 format (e.g., 'PT1H' for 1 hour)
                        meetingDuration: formatMeetingDuration(transformedParams.meetingDuration || transformedParams.duration || 60),
                        
                        // Additional parameters with reasonable defaults
                        maxCandidates: parseInt(transformedParams.maxCandidates || 10, 10),
                        minimumAttendeePercentage: parseInt(transformedParams.minimumAttendeePercentage || 100, 10)
                    };
                    
                    // Ensure we have at least one attendee
                    if (!apiData.attendees || apiData.attendees.length === 0) {
                        // Add the current user as an attendee if none provided
                        apiData.attendees = [{
                            emailAddress: { address: 'me@example.com' }, // This will be replaced by the API
                            type: 'required'
                        }];
                    }
                    
                    // Log using monitoring service only - avoid console logging in MCP adapter
                    if (monitoringService) {
                        monitoringService.debug('MCP ADAPTER - findMeetingTimes - FINAL API DATA', {
                            apiData: JSON.stringify(apiData, null, 2),
                            originalParams: JSON.stringify(transformedParams, null, 2),
                            apiPath,
                            apiMethod,
                            timestamp: new Date().toISOString()
                        }, 'mcp-adapter');
                    }
                    
                    // Log the transformed parameters for debugging
                    
                } catch (error) {
                    
                    throw new Error(`Failed to transform findMeetingTimes parameters: ${error.message}`);
                }
                break;
                
            case 'calendar.getRooms':
            case 'calendar.rooms':
                apiPath = '/v1/calendar/rooms';
                apiMethod = 'GET';
                
                // Transform parameters for room filtering
                if (transformedParams) {
                    const queryParams = [];
                    
                    // Add optional filters
                    if (transformedParams.building) {
                        queryParams.push(`building=${encodeURIComponent(transformedParams.building)}`);
                    }
                    
                    if (transformedParams.floor) {
                        queryParams.push(`floor=${encodeURIComponent(transformedParams.floor)}`);
                    }
                    
                    if (transformedParams.capacity) {
                        queryParams.push(`capacity=${transformedParams.capacity}`);
                    }
                    
                    // Add pagination parameters
                    if (transformedParams.top || transformedParams.limit) {
                        queryParams.push(`$top=${transformedParams.top || transformedParams.limit || 25}`);
                    }
                    
                    if (transformedParams.skip) {
                        queryParams.push(`$skip=${transformedParams.skip}`);
                    }
                    
                    // Add query parameters to the API path
                    if (queryParams.length > 0) {
                        apiPath += `?${queryParams.join('&')}`;
                    }
                }
                
                // Log the API call for debugging
                
                break;
                
            case 'calendar.getCalendars':
            case 'calendar.calendars':
            case 'microsoft calendar.getCalendars':
                apiPath = '/v1/calendar/calendars';
                apiMethod = 'GET';
                
                // Transform parameters for calendar options
                if (transformedParams) {
                    const queryParams = [];
                    
                    // Add optional filters
                    if (transformedParams.includeDelegated !== undefined) {
                        queryParams.push(`includeDelegated=${transformedParams.includeDelegated}`);
                    }
                    
                    if (transformedParams.includeShared !== undefined) {
                        queryParams.push(`includeShared=${transformedParams.includeShared}`);
                    }
                    
                    // Add query parameters to the API path
                    if (queryParams.length > 0) {
                        apiPath += `?${queryParams.join('&')}`;
                    }
                }
                
                // Log the API call for debugging
                
                break;
            case 'calendar.addAttachment':
                if (!transformedParams.id) {
                    throw new Error('Event ID is required for adding attachment')
                }
                apiPath = `/v1/calendar/events/${transformedParams.id}/attachments`;
                apiMethod = 'POST';
                apiData = {
                    name: transformedParams.name,
                    contentBytes: transformedParams.contentBytes,
                    contentType: transformedParams.contentType || 'application/octet-stream',
                    isInline: transformedParams.isInline || false
                };
                break;
            case 'calendar.removeAttachment':
                if (!transformedParams.eventId) {
                    throw new Error('Event ID is required for removing attachment')
                }
                if (!transformedParams.attachmentId) {
                    throw new Error('Attachment ID is required for removing attachment')
                }
                apiPath = `/v1/calendar/events/${transformedParams.eventId}/attachments/${transformedParams.attachmentId}`;
                apiMethod = 'DELETE';
                break;

            // Files module endpoints
            case 'files.listFiles':
                apiPath = '/v1/files';
                apiMethod = 'GET';
                break;
            case 'files.searchFiles':
                apiPath = '/v1/files/search';
                apiMethod = 'GET';
                // Map 'query' parameter to 'q' as expected by the controller
                if (params.query && !params.q) {
                    params.q = params.query;
                    delete params.query;
                }
                break;
            case 'files.downloadFile':
                apiPath = '/v1/files/download';
                apiMethod = 'GET';
                break;
            case 'files.uploadFile':
                apiPath = '/v1/files/upload';
                apiMethod = 'POST';
                apiData = params;
                break;
            case 'files.getFileMetadata':
                apiPath = '/v1/files/metadata';
                apiMethod = 'GET';
                break;
            case 'files.createSharingLink':
                apiPath = '/v1/files/share';
                apiMethod = 'POST';
                apiData = params;
                break;
            case 'files.getSharingLinks':
                apiPath = '/v1/files/sharing';
                apiMethod = 'GET';
                break;
            case 'files.removeSharingPermission':
                apiPath = '/v1/files/sharing/remove';
                apiMethod = 'POST';
                apiData = params;
                break;
            case 'files.getFileContent':
                apiPath = '/v1/files/content';
                apiMethod = 'GET';
                // Map 'fileId' parameter to 'id' as expected by the controller
                if (params.fileId && !params.id) {
                    params.id = params.fileId;
                    delete params.fileId;
                }
                break;
            case 'files.setFileContent':
                apiPath = '/v1/files/content';
                apiMethod = 'POST';
                apiData = params;
                break;
            case 'files.updateFileContent':
                apiPath = '/v1/files/content/update';
                apiMethod = 'POST';
                apiData = params;
                break;

            // People module endpoints
            case 'people.getRelevantPeople':
                apiPath = '/v1/people';
                apiMethod = 'GET';
                // Map any limit parameter correctly
                if (params.limit) {
                    params.limit = parseInt(params.limit, 10);
                }
                break;
            case 'people.searchPeople':
            case 'people.search':
                apiPath = '/v1/people/search';
                apiMethod = 'GET';
                // Make sure query parameter is preserved
                if (!params.query && params.q) {
                    params.query = params.q;
                    // Remove the q param to avoid confusion
                    delete params.q;
                }
                break;
            case 'people.findPeople':
            case 'people.find':
                apiPath = '/v1/people/find';
                apiMethod = 'GET';
                // Make sure query parameter is preserved
                if (!params.query && params.q) {
                    params.query = params.q;
                    // Remove the q param to avoid confusion
                    delete params.q;
                }
                // Ensure limit is parsed as integer
                if (params.limit) {
                    params.limit = parseInt(params.limit, 10);
                }
                break;
            case 'people.getPersonById':
                if (!params.id) {
                    throw new Error('Person ID is required for getPersonById');
                }
                apiPath = `/v1/people/${params.id}`;
                apiMethod = 'GET';
                break;

            // Teams module endpoints
            case 'teams.listChats':
                apiPath = '/v1/teams/chats';
                apiMethod = 'GET';
                break;
            case 'teams.getChat':
                if (!params.chatId) {
                    throw new Error('Chat ID is required for getChat');
                }
                apiPath = `/v1/teams/chats/${params.chatId}`;
                apiMethod = 'GET';
                break;
            case 'teams.listChatMessages':
            case 'teams.getChatMessages':
                if (!params.chatId) {
                    throw new Error('Chat ID is required for listChatMessages');
                }
                apiPath = `/v1/teams/chats/${params.chatId}/messages`;
                apiMethod = 'GET';
                break;
            case 'teams.sendChatMessage':
                if (!params.chatId) {
                    throw new Error('Chat ID is required for sendChatMessage');
                }
                apiPath = `/v1/teams/chats/${params.chatId}/messages`;
                apiMethod = 'POST';
                apiData = {
                    content: params.content,
                    contentType: params.contentType || 'text'
                };
                break;
            case 'teams.listTeams':
            case 'teams.listJoinedTeams':
                apiPath = '/v1/teams';
                apiMethod = 'GET';
                break;
            case 'teams.getTeam':
                if (!params.teamId) {
                    throw new Error('Team ID is required for getTeam');
                }
                apiPath = `/v1/teams/${params.teamId}`;
                apiMethod = 'GET';
                break;
            case 'teams.listChannels':
            case 'teams.listTeamChannels':
                if (!params.teamId) {
                    throw new Error('Team ID is required for listChannels');
                }
                apiPath = `/v1/teams/${params.teamId}/channels`;
                apiMethod = 'GET';
                break;
            case 'teams.getChannel':
                if (!params.teamId || !params.channelId) {
                    throw new Error('Team ID and Channel ID are required for getChannel');
                }
                apiPath = `/v1/teams/${params.teamId}/channels/${params.channelId}`;
                apiMethod = 'GET';
                break;
            case 'teams.listChannelMessages':
            case 'teams.getChannelMessages':
                if (!params.teamId || !params.channelId) {
                    throw new Error('Team ID and Channel ID are required for listChannelMessages');
                }
                apiPath = `/v1/teams/${params.teamId}/channels/${params.channelId}/messages`;
                apiMethod = 'GET';
                break;
            case 'teams.sendChannelMessage':
                if (!params.teamId || !params.channelId) {
                    throw new Error('Team ID and Channel ID are required for sendChannelMessage');
                }
                apiPath = `/v1/teams/${params.teamId}/channels/${params.channelId}/messages`;
                apiMethod = 'POST';
                apiData = {
                    content: params.content,
                    contentType: params.contentType || 'text'
                };
                break;
            case 'teams.replyToMessage':
                if (!params.teamId || !params.channelId || !params.messageId) {
                    throw new Error('Team ID, Channel ID, and Message ID are required for replyToMessage');
                }
                apiPath = `/v1/teams/${params.teamId}/channels/${params.channelId}/messages/${params.messageId}/replies`;
                apiMethod = 'POST';
                apiData = {
                    content: params.content,
                    contentType: params.contentType || 'text'
                };
                break;
            case 'teams.listOnlineMeetings':
                apiPath = '/v1/teams/meetings';
                apiMethod = 'GET';
                break;
            case 'teams.createOnlineMeeting':
                apiPath = '/v1/teams/meetings';
                apiMethod = 'POST';
                apiData = {
                    subject: params.subject,
                    startDateTime: params.startDateTime,
                    endDateTime: params.endDateTime,
                    participants: params.participants
                };
                break;
            case 'teams.getOnlineMeeting':
                if (!params.meetingId) {
                    throw new Error('Meeting ID is required for getOnlineMeeting');
                }
                apiPath = `/v1/teams/meetings/${params.meetingId}`;
                apiMethod = 'GET';
                break;
            case 'teams.getMeetingByJoinUrl':
                apiPath = '/v1/teams/meetings/findByJoinUrl';
                apiMethod = 'GET';
                break;
            case 'teams.getMeetingTranscripts':
                if (!params.meetingId) {
                    throw new Error('Meeting ID is required for getMeetingTranscripts');
                }
                apiPath = `/v1/teams/meetings/${params.meetingId}/transcripts`;
                apiMethod = 'GET';
                break;
            case 'teams.getMeetingTranscriptContent':
                if (!params.meetingId) {
                    throw new Error('Meeting ID is required for getMeetingTranscriptContent');
                }
                if (!params.transcriptId) {
                    throw new Error('Transcript ID is required for getMeetingTranscriptContent');
                }
                apiPath = `/v1/teams/meetings/${params.meetingId}/transcripts/${params.transcriptId}`;
                apiMethod = 'GET';
                break;

            // Search module endpoints
            case 'search.search':
                apiPath = '/v1/search';
                apiMethod = 'POST';
                apiData = {
                    query: params.query,
                    entityTypes: params.entityTypes,
                    from: params.from,
                    size: params.size
                };
                break;

            // To-Do module endpoints
            case 'todo.listTaskLists':
                apiPath = '/v1/todo/lists';
                apiMethod = 'GET';
                break;
            case 'todo.getTaskList':
                if (!params.listId) {
                    throw new Error('List ID is required for getTaskList');
                }
                apiPath = `/v1/todo/lists/${params.listId}`;
                apiMethod = 'GET';
                break;
            case 'todo.createTaskList':
                apiPath = '/v1/todo/lists';
                apiMethod = 'POST';
                apiData = {
                    displayName: params.displayName
                };
                break;
            case 'todo.updateTaskList':
                if (!params.listId) {
                    throw new Error('List ID is required for updateTaskList');
                }
                apiPath = `/v1/todo/lists/${params.listId}`;
                apiMethod = 'PATCH';
                apiData = {
                    displayName: params.displayName
                };
                break;
            case 'todo.deleteTaskList':
                if (!params.listId) {
                    throw new Error('List ID is required for deleteTaskList');
                }
                apiPath = `/v1/todo/lists/${params.listId}`;
                apiMethod = 'DELETE';
                break;
            case 'todo.listTasks':
                // If no listId provided, get the default task list first
                if (!params.listId) {
                    const listsResult = await callApi('GET', '/v1/todo/lists');
                    const lists = listsResult?.data || listsResult?.taskLists || [];
                    if (lists.length > 0) {
                        // Find default list or use first one
                        const defaultList = lists.find(l => l.wellknownListName === 'defaultList') || lists[0];
                        params.listId = defaultList.id;
                    } else {
                        throw new Error('No task lists found. Create a task list first.');
                    }
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks`;
                apiMethod = 'GET';
                break;
            case 'todo.getTask':
                if (!params.listId || !params.taskId) {
                    throw new Error('List ID and Task ID are required for getTask');
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks/${params.taskId}`;
                apiMethod = 'GET';
                break;
            case 'todo.createTask':
                // If no listId provided, get the default task list first
                if (!params.listId) {
                    const listsResult = await callApi('GET', '/v1/todo/lists');
                    const lists = listsResult?.data || listsResult?.taskLists || [];
                    if (lists.length > 0) {
                        const defaultList = lists.find(l => l.wellknownListName === 'defaultList') || lists[0];
                        params.listId = defaultList.id;
                    } else {
                        throw new Error('No task lists found. Create a task list first.');
                    }
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks`;
                apiMethod = 'POST';
                apiData = {
                    title: params.title,
                    body: params.body,
                    importance: params.importance,
                    dueDateTime: params.dueDateTime,
                    reminderDateTime: params.reminderDateTime,
                    isReminderOn: params.isReminderOn
                };
                break;
            case 'todo.updateTask':
                if (!params.listId || !params.taskId) {
                    throw new Error('List ID and Task ID are required for updateTask');
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks/${params.taskId}`;
                apiMethod = 'PATCH';
                apiData = {
                    title: params.title,
                    body: params.body,
                    importance: params.importance,
                    status: params.status,
                    dueDateTime: params.dueDateTime,
                    reminderDateTime: params.reminderDateTime,
                    isReminderOn: params.isReminderOn
                };
                break;
            case 'todo.deleteTask':
                if (!params.listId || !params.taskId) {
                    throw new Error('List ID and Task ID are required for deleteTask');
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks/${params.taskId}`;
                apiMethod = 'DELETE';
                break;
            case 'todo.completeTask':
                if (!params.listId || !params.taskId) {
                    throw new Error('List ID and Task ID are required for completeTask');
                }
                apiPath = `/v1/todo/lists/${params.listId}/tasks/${params.taskId}/complete`;
                apiMethod = 'POST';
                break;

            // Contacts module endpoints
            case 'contacts.listContacts':
                apiPath = '/v1/contacts';
                apiMethod = 'GET';
                break;
            case 'contacts.getContact':
                if (!params.contactId) {
                    throw new Error('Contact ID is required for getContact');
                }
                apiPath = `/v1/contacts/${params.contactId}`;
                apiMethod = 'GET';
                break;
            case 'contacts.createContact':
                apiPath = '/v1/contacts';
                apiMethod = 'POST';
                apiData = {
                    givenName: params.givenName,
                    surname: params.surname,
                    displayName: params.displayName,
                    emailAddresses: params.emailAddresses,
                    businessPhones: params.businessPhones,
                    mobilePhone: params.mobilePhone,
                    companyName: params.companyName,
                    jobTitle: params.jobTitle
                };
                break;
            case 'contacts.updateContact':
                if (!params.contactId) {
                    throw new Error('Contact ID is required for updateContact');
                }
                apiPath = `/v1/contacts/${params.contactId}`;
                apiMethod = 'PATCH';
                apiData = {
                    givenName: params.givenName,
                    surname: params.surname,
                    displayName: params.displayName,
                    emailAddresses: params.emailAddresses,
                    businessPhones: params.businessPhones,
                    mobilePhone: params.mobilePhone,
                    companyName: params.companyName,
                    jobTitle: params.jobTitle
                };
                break;
            case 'contacts.deleteContact':
                if (!params.contactId) {
                    throw new Error('Contact ID is required for deleteContact');
                }
                apiPath = `/v1/contacts/${params.contactId}`;
                apiMethod = 'DELETE';
                break;
            case 'contacts.searchContacts':
                apiPath = '/v1/contacts/search';
                apiMethod = 'GET';
                break;

            // Groups module endpoints
            case 'groups.listGroups':
                apiPath = '/v1/groups';
                apiMethod = 'GET';
                break;
            case 'groups.getGroup':
                if (!params.groupId) {
                    throw new Error('Group ID is required for getGroup');
                }
                apiPath = `/v1/groups/${params.groupId}`;
                apiMethod = 'GET';
                break;
            case 'groups.listGroupMembers':
                if (!params.groupId) {
                    throw new Error('Group ID is required for listGroupMembers');
                }
                apiPath = `/v1/groups/${params.groupId}/members`;
                apiMethod = 'GET';
                break;
            case 'groups.listMyGroups':
                apiPath = '/v1/groups/my';
                apiMethod = 'GET';
                break;

            // System endpoints
            case 'system.getToolDefinitions':
                apiPath = '/tools';
                apiMethod = 'GET';
                break;
            case 'system.getManifest':
                apiPath = '/tools';  // Same as getToolDefinitions but we'll transform the response
                apiMethod = 'GET';
                break;
            case 'system.getResources':
                // Return empty resources list as this isn't implemented yet
                return { resources: [] };
            case 'system.getPrompts':
                // Return empty prompts list as this isn't implemented yet
                return { prompts: [] };
            default:
                throw new Error(`Unknown API endpoint for ${moduleName}.${methodName}`);
        }

        // For GET requests with params, add them as query parameters
        if (apiMethod === 'GET' && Object.keys(params).length > 0) {
            const queryParams = new URLSearchParams();
            for (const [key, value] of Object.entries(params)) {
                // Special handling for complex objects and arrays to prevent [object Object] in URLs
                if (typeof value === 'object' && value !== null) {
                    continue; // Skip adding complex objects to query params
                } else {
                    queryParams.append(key, value);
                }
            }
            
            if (queryParams.toString()) {
                apiPath += `?${queryParams.toString()}`;
            }
        }

        // Add enhanced logging for calendar controller API calls
        if (moduleName === 'calendar' || apiPath.startsWith('/v1/calendar')) {
            logDebug(`[MCP Adapter] Making calendar API call: ${apiMethod} ${apiPath}`, {
                method: apiMethod,
                path: apiPath,
                data: apiData,
                timestamp: new Date().toISOString(),
                moduleMethod: `${moduleName}.${methodName}`
            });
        }
        
        // Make the API call
        const result = await callApi(apiMethod, apiPath, apiData);
        
        // Log the result for calendar calls
        if (moduleName === 'calendar' || apiPath.startsWith('/v1/calendar')) {
            logDebug(`[MCP Adapter] Calendar API call result: ${apiMethod} ${apiPath}`, {
                success: true,
                resultType: typeof result,
                hasData: result != null,
                timestamp: new Date().toISOString(),
                moduleMethod: `${moduleName}.${methodName}`
            });
        }

        // Special case for system.getManifest to transform tools list to manifest format
        if (`${moduleName}.${methodName}` === 'system.getManifest' && result && result.tools) {
            return {
                protocolVersion: params.protocolVersion || "2024-11-05",
                capabilities: {
                    toolInvocation: true,
                    manifest: true
                },
                serverInfo: {
                    name: "Microsoft 365 MCP Gateway",
                    version: "1.0.0"
                }
            };
        }

        // TODO: [Streaming] Handle streaming responses if API supports it (e.g., for file downloads) (MEDIUM)

        return result;
    } catch (error) {
        
        if (error.message.includes('authentication') || error.message.includes('unauthorized') || error.message.includes('401')) {
            return { 
                error: 'Authentication required. Please log in via the web UI.',
                errorType: 'auth_required'
            };
        }
        return { 
            error: `Error executing ${moduleName}.${methodName}: ${error.message}`,
            errorType: 'module_error'
        };
    }
}

// Helper function to handle calendar timeframe conversion
function processCalendarTimeframe(toolArgs) {
    // Process date ranges for calendar
    const timeframe = toolArgs.timeframe || 'today';
    let start, end;

    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];

    switch (timeframe) {
        case 'today':
            start = todayStr;
            end = todayStr;
            break;
        case 'week':
            // Start of current week (Sunday)
            const startOfWeek = new Date(today);
            startOfWeek.setDate(today.getDate() - today.getDay());
            start = startOfWeek.toISOString().split('T')[0];

            // End of current week (Saturday)
            const endOfWeek = new Date(today);
            endOfWeek.setDate(today.getDate() + (6 - today.getDay()));
            end = endOfWeek.toISOString().split('T')[0];
            break;
        case 'month':
            // Start of current month
            start = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-01`;

            // End of current month
            const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0);
            end = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(lastDay.getDate()).padStart(2, '0')}`;
            break;
        default:
            // Default to today if timeframe is not recognized
            start = todayStr;
            end = todayStr;
    }

    

    // Add date range to the arguments
    return {
        ...toolArgs,
        start,
        end
    };
}

// Helper: Map MCP tool calls to module methods
async function handleToolCall(toolName, toolArgs) {
    try {
        // Make sure adapter is initialized
        if (!adapterState.initialized) {
            await initializeAdapter();
        }

        // Update last activity timestamp
        adapterState.lastActivity = Date.now();
        
        // Special handling for calendar date ranges
        if ((toolName === 'getEvents' || toolName === 'getCalendar') && toolArgs.timeframe) {
            toolArgs = processCalendarTimeframe(toolArgs);
        }
        
        // Special handling for file operations
        // Map common file operation tool names directly to their module methods
        const fileOperations = {
            'searchFiles': 'files.searchFiles',
            'getFileContent': 'files.getFileContent',
            'createSharingLink': 'files.createSharingLink',
            'updateFileContent': 'files.updateFileContent',
            'getSharingLinks': 'files.getSharingLinks',
            'removeSharingPermission': 'files.removeSharingPermission',
            'setFileContent': 'files.setFileContent',
            'downloadFile': 'files.downloadFile',
            'uploadFile': 'files.uploadFile',
            'getFileMetadata': 'files.getFileMetadata',
            'listFiles': 'files.listFiles'
        };
        
        // If the tool is a file operation, map it directly
        if (fileOperations[toolName]) {
            const [moduleName, methodName] = fileOperations[toolName].split('.');
            return await executeModuleMethod(moduleName, methodName, toolArgs);
        }
        
        // First try to use the toolsService to map and transform the parameters
        const { mapping, params } = toolsService.transformToolParameters(toolName, toolArgs);
        
        if (mapping) {
            return await executeModuleMethod(mapping.moduleName, mapping.methodName, params);
        }
        
        // If tools service can't map it, try the API approach
        try {
            const toolsResponse = await callApi('GET', '/tools');
            if (toolsResponse && Array.isArray(toolsResponse.tools)) {
                // Find the tool definition
                const toolDef = toolsResponse.tools.find(tool => tool.name === toolName);
                if (toolDef) {
                    // Extract module and method from endpoint
                    // Example: /api/v1/mail/send -> mail.sendEmail
                    const path = toolDef.endpoint.replace('/api/v1/', '');
                    const parts = path.split('/');

                    if (parts.length > 0) {
                        const moduleName = parts[0];
                        let methodName = parts.length > 1 ? parts[1] : 'get' + moduleName.charAt(0).toUpperCase() + moduleName.slice(1);

                        // Special case mappings
                        const methodMappings = {
                            'mail': { '': 'getInbox' },
                            'calendar': { '': 'getEvents' },
                            'files': { '': 'listFiles' },
                            'people': { '': 'getRelevantPeople' }
                        };

                        if (parts.length === 1 && methodMappings[moduleName] && methodMappings[moduleName]['']) {
                            methodName = methodMappings[moduleName][''];
                        }

                        return await executeModuleMethod(moduleName, methodName, toolArgs);
                    }
                }
            }
        } catch (error) {
            // If API call fails, silently continue
        }
        
        // If we get here, we couldn't map the tool
        return { error: `Unknown tool: ${toolName}` };
    } catch (error) {
        return { 
            error: `Error handling tool call: ${error.message}`,
            errorType: 'tool_error'
        };
    }
}

// Helper: handle incoming JSON-RPC requests
async function handleRequest(msg) {
    const { id, method, params } = msg;

    // Silent - no logging for MCP protocol compliance

    try {
        let result;

        // Handle standard MCP methods
        switch (method) {
            case 'initialize': {
                // Standard MCP handshake with protocolVersion, capabilities
                result = {
                    protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                    capabilities: {
                        toolInvocation: true,
                        manifest: true
                    },
                    serverInfo: {
                        name: "MCP Microsoft Office Adapter",
                        version: "1.0.0"
                    }
                };
                sendResponse(id, result, null);
                return;
            }
            case 'tools/call': {
                if (!params || !params.name) {
                    sendResponse(id, null, { code: -32602, message: 'Missing tool name in params' });
                    return;
                }

                const toolName = params.name;
                const toolArgs = params.arguments || {};

                try {
                    // Use the handleToolCall function to process the tool call
                    const toolResult = await handleToolCall(toolName, toolArgs);

                    if (toolResult && toolResult.error) {
                        sendResponse(id, null, { code: -32603, message: `Tool error: ${toolResult.error}` });
                        return;
                    }

                    // Format the result according to MCP spec
                    const formattedResult = {
                        content: [
                            {
                                type: "text",
                                text: JSON.stringify(toolResult, null, 2)
                            }
                        ],
                        isError: false
                    };

                    // Success response
                    sendResponse(id, formattedResult, null);
                    return;
                } catch (err) {
                    sendResponse(id, null, { code: -32603, message: `Tool error: ${err.message}` });
                    return;
                }
                break;
            }
            case 'tools/list':
                // Try to get tools list from the API first, then fallback to tools service if needed
                try {
                    const toolsResult = await callApi('GET', '/tools');

                    // Helper to format tools for MCP response
                    const formatTool = (tool) => ({
                        name: tool.name,
                        description: tool.description,
                        inputSchema: {
                            type: 'object',
                            properties: Object.entries(tool.parameters || {}).reduce((acc, [key, value]) => {
                                acc[key] = {
                                    type: value.type || 'string',
                                    description: value.description || ''
                                };
                                return acc;
                            }, {}),
                            required: Object.entries(tool.parameters || {})
                                .filter(([_, value]) => !value.optional)
                                .map(([key, _]) => key)
                        }
                    });

                    let mcpTools = [];
                    if (!toolsResult || !toolsResult.tools) {
                        // Use the tools service to get the tools list
                        const toolsFromService = toolsService.getAllTools();
                        if (toolsFromService && toolsFromService.length > 0) {
                            mcpTools = toolsFromService.map(formatTool);
                        }
                    } else {
                        mcpTools = toolsResult.tools.map(formatTool);
                    }

                    // Filter tools based on user's permissions
                    const filteredTools = await filterToolsByPermissions(mcpTools);

                    result = {
                        protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                        tools: filteredTools
                    };
                } catch (error) {

                    // Use the tools service as fallback
                    try {
                        const toolsFromService = toolsService.getAllTools();

                        if (toolsFromService && toolsFromService.length > 0) {
                            // Format tools for MCP response
                            const mcpTools = toolsFromService.map(tool => ({
                                name: tool.name,
                                description: tool.description,
                                inputSchema: {
                                    type: 'object',
                                    properties: Object.entries(tool.parameters || {}).reduce((acc, [key, value]) => {
                                        acc[key] = {
                                            type: value.type || 'string',
                                            description: value.description || ''
                                        };
                                        return acc;
                                    }, {}),
                                    required: Object.entries(tool.parameters || {})
                                        .filter(([_, value]) => !value.optional)
                                        .map(([key, _]) => key)
                                }
                            }));

                            // Filter tools based on user's permissions
                            const filteredTools = await filterToolsByPermissions(mcpTools);

                            result = {
                                protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                                tools: filteredTools
                            };
                        } else {
                            // If tools service also fails, return empty list
                            result = {
                                protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                                tools: []
                            };
                        }
                    } catch (serviceError) {
                        // Final fallback - empty list
                        result = {
                            protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                            tools: []
                        };
                    }
                }
                sendResponse(id, result, null);
                return;
            case 'getManifest':
                // Legacy method - same as tools/list for compatibility
                // Create a standard MCP manifest
                result = {
                    protocolVersion: (params && params.protocolVersion) ? params.protocolVersion : "2024-11-05",
                    capabilities: {
                        toolInvocation: true,
                        manifest: true
                    },
                    serverInfo: {
                        name: "Microsoft 365 MCP Gateway",
                        version: "1.0.0"
                    }
                };
                sendResponse(id, result, null);
                return;
            case 'resources/list':
                // Return empty resources list (not implemented yet)
                result = { resources: [] };
                sendResponse(id, result, null);
                return;
            case 'prompts/list':
                // Return empty prompts list (not implemented yet)
                result = { prompts: [] };
                sendResponse(id, result, null);
                return;
            default:
                sendResponse(id, null, { code: -32601, message: `Method not found: ${method}` });
                return;
        }
    } catch (err) {
        sendResponse(id, null, { code: -32603, message: `Internal error: ${err.message}` });
    }
}

// Helper: handle core requests (manifest, listResources)
async function _handleCoreRequest(coreMethodName, requestId, coreParams = {}) {
    
    try {
        const hasAccess = await checkModuleAccess();
        if (!hasAccess) {
            sendResponse(requestId, null, createJsonRpcError(-32001, 'Access Denied', 'Backend or modules unavailable.'));
            return;
        }
        const result = await executeModuleMethod('core', coreMethodName, coreParams);
        sendResponse(requestId, result);
    } catch (error) { 
        
        sendResponse(requestId, null, createJsonRpcError(-32603, 'Internal Error', `Error handling ${coreMethodName}: ${error.message}`));
    }
}

// Initialize: check backend API availability
async function initialize() {
    // Allow skipping initialization for specific test environments
    if (process.env.MCP_SKIP_INIT === 'true') {
        
        adapterState.initialized = true; // Mark as initialized to avoid re-triggering
        adapterState.backendAvailable = true; // Assume available for tests
        return Promise.resolve(true);
    }

    
    const ok = await initializeAdapter();
    if (!ok) {
        
    } else {
        
    }
    return Promise.resolve(ok);
}

// Graceful shutdown
async function shutdown() {
    
    if (healthCheckInterval) {
        clearInterval(healthCheckInterval);
        healthCheckInterval = null;
        
    }
    // TODO: [Cleanup] Add any other necessary cleanup (e.g., close connections) (LOW)
}

// Main: read stdin line by line, handle JSON-RPC
const rl = readline.createInterface({ input: process.stdin });
rl.on('line', (line) => {
    try {
        const msg = JSON.parse(line);
        const { id, method, params } = msg;
        // MCP PATCH: Ignore JSON-RPC notifications (method present, id undefined)
        if (method && typeof id === 'undefined') {
            // Do not respond to notifications, per JSON-RPC spec
            return;
        }
        // Normal request: handle as before
        handleRequest(msg);
    } catch (err) {
        // Silent - invalid JSON-RPC
    }
});

// Handle stdin close
rl.on('close', () => {
    
    process.exit(0);
});

// Handle process signals for clean shutdown
process.on('SIGINT', () => {
    
    shutdown().then(() => process.exit(0));
});

process.on('SIGTERM', () => {
    
    shutdown().then(() => process.exit(0));
});

// Console methods already silenced above - MCP requires absolute silence

// Run initialization (proxy approach)
initialize();

/**
 * Register a new device with the MCP server
 * @returns {Promise<Object>} - Device registration response with user code and verification URI
 */
async function registerDevice() {
    const MAX_REGISTRATION_RETRIES = 3;
    const REGISTRATION_TIMEOUT_MS = 30000; // 30 seconds
    
    for (let attempt = 1; attempt <= MAX_REGISTRATION_RETRIES; attempt++) {
        try {
            // Log registration attempt
            if (attempt > 1) {
                process.stderr.write(`🔄 Device registration attempt ${attempt}/${MAX_REGISTRATION_RETRIES}...\n`);
            }

            // Ensure OAuth discovery has been performed
            if (!oauthDiscovery.discovered) {
                const discoverySuccess = await discoverOAuthEndpoints();
                if (!discoverySuccess) {
                    throw new Error('OAuth discovery failed - cannot register device');
                }
            }

            // Create a timeout promise for the registration request
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => reject(new Error('Device registration timed out')), REGISTRATION_TIMEOUT_MS);
            });

            // Call the device registration endpoint with timeout
            const registrationPromise = callApi('POST', DEVICE_REGISTRATION_PATH, {
                client_info: {
                    name: 'MCP Microsoft Office Adapter',
                    version: '1.0.0',
                    platform: process.platform,
                    arch: process.arch
                }
            });

            const registrationData = await Promise.race([registrationPromise, timeoutPromise]);

            if (!registrationData || !registrationData.device_code) {
                throw new Error('Invalid device registration response - missing device_code');
            }

            if (!registrationData.user_code || !registrationData.verification_uri) {
                throw new Error('Invalid device registration response - missing user_code or verification_uri');
            }

            // Store device credentials
            deviceCredentials.deviceId = registrationData.device_id;
            deviceCredentials.deviceSecret = registrationData.device_secret;
            
            // Log successful registration
            process.stderr.write('✅ Device registration successful\n');
            
            return {
                deviceCode: registrationData.device_code,
                userCode: registrationData.user_code,
                verificationUri: registrationData.verification_uri,
                verificationUriComplete: registrationData.verification_uri_complete,
                expiresIn: registrationData.expires_in,
                interval: registrationData.interval || 5
            };
        } catch (error) {
            const isLastAttempt = attempt === MAX_REGISTRATION_RETRIES;
            
            // Categorize error types for better handling
            if (error.message.includes('timeout') || error.message.includes('TIMEOUT')) {
                process.stderr.write(`⚠️  Registration timeout on attempt ${attempt}\n`);
            } else if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                process.stderr.write(`⚠️  Network connection failed on attempt ${attempt}\n`);
            } else if (error.message.includes('OAuth discovery failed')) {
                process.stderr.write(`⚠️  OAuth discovery failed on attempt ${attempt}\n`);
            } else {
                process.stderr.write(`⚠️  Registration error on attempt ${attempt}: ${error.message}\n`);
            }
            
            if (isLastAttempt) {
                // Final attempt failed - provide detailed error
                if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                    throw new Error('Device registration failed: Unable to connect to MCP server. Please check your network connection and server URL.');
                } else if (error.message.includes('timeout')) {
                    throw new Error('Device registration failed: Request timed out. The MCP server may be overloaded or unreachable.');
                } else {
                    throw new Error(`Device registration failed after ${MAX_REGISTRATION_RETRIES} attempts: ${error.message}`);
                }
            }
            
            // Wait before retrying (exponential backoff)
            const backoffDelay = 1000 * Math.pow(2, attempt - 1); // 1s, 2s
            await new Promise(resolve => setTimeout(resolve, backoffDelay));
        }
    }
}

/**
 * Display user code and verification instructions to the user
 * @param {Object} deviceInfo - Device registration information
 */
function displayUserCode(deviceInfo) {
    // Use stderr for user-facing messages (not stdout which is reserved for JSON-RPC)
    process.stderr.write('\n=== MCP Microsoft Office Authentication Required ===\n');
    process.stderr.write('To authenticate this MCP adapter with the Microsoft Office server:\n\n');
    process.stderr.write(`1. Open your web browser and go to: ${deviceInfo.verificationUri}\n`);
    process.stderr.write(`2. Enter this code: ${deviceInfo.userCode}\n`);
    process.stderr.write(`3. Follow the instructions to sign in with your Microsoft account\n\n`);
    
    if (deviceInfo.verificationUriComplete) {
        process.stderr.write(`Or use this direct link: ${deviceInfo.verificationUriComplete}\n\n`);
    }
    
    process.stderr.write('Waiting for authentication...\n');
    process.stderr.write('(This adapter will automatically continue once you complete authentication)\n\n');
}

/**
 * Poll for access tokens after user authorization
 * @param {string} deviceCode - Device code from registration
 * @param {number} interval - Polling interval in seconds
 * @param {number} expiresIn - Device code expiration time in seconds
 * @returns {Promise<Object>} - Token response with access and refresh tokens
 */
async function pollForTokens(deviceCode, interval = 5, expiresIn = 600) {
    const POLL_TIMEOUT_MS = 15000; // 15 seconds per poll request
    const MAX_CONSECUTIVE_ERRORS = 5;
    
    let consecutiveErrors = 0;
    let pollCount = 0;
    
    const startTime = Date.now();
    const expirationTime = startTime + (expiresIn * 1000);
    
    process.stderr.write('🔄 Waiting for user authorization...\n');
    
    while (Date.now() < expirationTime) {
        pollCount++;
        
        try {
            // Create a timeout promise for each poll request
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => reject(new Error('Poll request timed out')), POLL_TIMEOUT_MS);
            });

            // Poll the token endpoint with timeout
            const tokenPromise = callApi('POST', DEVICE_TOKEN_PATH, {
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                device_code: deviceCode,
                client_id: deviceCredentials.deviceId
            });

            const tokenData = await Promise.race([tokenPromise, timeoutPromise]);

            if (tokenData && tokenData.access_token) {
                // Store tokens
                deviceCredentials.accessToken = tokenData.access_token;
                deviceCredentials.refreshToken = tokenData.refresh_token;
                deviceCredentials.tokenExpiry = Date.now() + (tokenData.expires_in * 1000);
                
                // Update refresh token if provided
                if (tokenData.refresh_token) {
                    deviceCredentials.refreshToken = tokenData.refresh_token;
                }
                
                process.stderr.write('✅ Authentication successful!\n\n');
                return tokenData;
            }
            
            // Reset consecutive error count on successful request
            consecutiveErrors = 0;
            
        } catch (error) {
            consecutiveErrors++;
            
            // Handle specific OAuth errors
            if (error.message.includes('authorization_pending')) {
                // User hasn't completed authorization yet - continue polling
                // Log progress every 10 polls to show we're still waiting
                if (pollCount % 10 === 0) {
                    const remainingTime = Math.ceil((expirationTime - Date.now()) / 1000);
                    process.stderr.write(`⏳ Still waiting for authorization... (${remainingTime}s remaining)\n`);
                }
            } else if (error.message.includes('slow_down')) {
                // Server requested slower polling - increase interval
                const oldInterval = interval;
                interval = Math.min(interval * 2, 30); // Cap at 30 seconds
                process.stderr.write(`⚠️  Server requested slower polling. Increasing interval from ${oldInterval}s to ${interval}s\n`);
            } else if (error.message.includes('expired_token')) {
                throw new Error('Device code expired. Please restart the authentication process.');
            } else if (error.message.includes('access_denied')) {
                throw new Error('Authentication was denied by the user.');
            } else if (error.message.includes('Poll request timed out')) {
                process.stderr.write(`⚠️  Poll request ${pollCount} timed out\n`);
            } else if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                process.stderr.write(`⚠️  Network connection failed during poll ${pollCount}\n`);
            } else {
                // Other error - log it but continue polling
                process.stderr.write(`⚠️  Poll error ${pollCount}: ${error.message}\n`);
            }
            
            // If we have too many consecutive errors, fail fast
            if (consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
                if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                    throw new Error(`Authentication failed: Lost connection to MCP server after ${MAX_CONSECUTIVE_ERRORS} consecutive network errors.`);
                } else if (error.message.includes('timeout')) {
                    throw new Error(`Authentication failed: Server not responding after ${MAX_CONSECUTIVE_ERRORS} consecutive timeouts.`);
                } else {
                    throw new Error(`Authentication failed: Too many consecutive errors (${MAX_CONSECUTIVE_ERRORS}). Last error: ${error.message}`);
                }
            }
            
            // Wait for the specified interval before next poll
            await new Promise(resolve => setTimeout(resolve, interval * 1000));
        }
    }

    throw new Error('Authentication timed out. Please restart the authentication process.');
}

/**
 * Refresh the access token using the refresh token
 * @returns {Promise<boolean>} - True if refresh was successful
 */
async function refreshAccessToken() {
    const REFRESH_TIMEOUT_MS = 15000; // 15 seconds
    const MAX_REFRESH_RETRIES = 2;
    
    for (let attempt = 1; attempt <= MAX_REFRESH_RETRIES; attempt++) {
        try {
            if (!deviceCredentials.refreshToken) {
                process.stderr.write('⚠️  No refresh token available for token refresh\n');
                return false;
            }

            if (attempt > 1) {
                process.stderr.write(`🔄 Token refresh attempt ${attempt}/${MAX_REFRESH_RETRIES}...\n`);
            }

            // Create a timeout promise for the refresh request
            const timeoutPromise = new Promise((_, reject) => {
                setTimeout(() => reject(new Error('Token refresh timed out')), REFRESH_TIMEOUT_MS);
            });

            const refreshPromise = callApi('POST', DEVICE_TOKEN_PATH, {
                grant_type: 'refresh_token',
                refresh_token: deviceCredentials.refreshToken,
                client_id: deviceCredentials.deviceId
            });

            const tokenData = await Promise.race([refreshPromise, timeoutPromise]);

            if (tokenData && tokenData.access_token) {
                deviceCredentials.accessToken = tokenData.access_token;
                deviceCredentials.tokenExpiry = Date.now() + (tokenData.expires_in * 1000);
                
                // Update refresh token if provided
                if (tokenData.refresh_token) {
                    deviceCredentials.refreshToken = tokenData.refresh_token;
                }
                
                process.stderr.write('✅ Token refresh successful\n');
                return true;
            } else {
                process.stderr.write('⚠️  Token refresh returned invalid response\n');
            }
            
        } catch (error) {
            const isLastAttempt = attempt === MAX_REFRESH_RETRIES;
            
            // Categorize error types for better handling
            if (error.message.includes('invalid_grant') || error.message.includes('invalid_token')) {
                // Refresh token is invalid or expired - clear all tokens
                process.stderr.write('⚠️  Refresh token expired or invalid - will require re-authentication\n');
                deviceCredentials.accessToken = null;
                deviceCredentials.refreshToken = null;
                deviceCredentials.tokenExpiry = null;
                return false;
            } else if (error.message.includes('timeout') || error.message.includes('TIMEOUT')) {
                process.stderr.write(`⚠️  Token refresh timeout on attempt ${attempt}\n`);
            } else if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                process.stderr.write(`⚠️  Network connection failed during token refresh attempt ${attempt}\n`);
            } else {
                process.stderr.write(`⚠️  Token refresh error on attempt ${attempt}: ${error.message}\n`);
            }
            
            if (isLastAttempt) {
                // Final attempt failed - clear tokens to force re-authentication
                process.stderr.write('❌ Token refresh failed - clearing tokens to force re-authentication\n');
                deviceCredentials.accessToken = null;
                deviceCredentials.refreshToken = null;
                deviceCredentials.tokenExpiry = null;
                return false;
            }
            
            // Wait before retrying
            const backoffDelay = 2000 * attempt; // 2s, 4s
            await new Promise(resolve => setTimeout(resolve, backoffDelay));
        }
    }
    
    return false;
}

/**
 * Check if the current access token is valid and refresh if needed
 * @returns {Promise<boolean>} - True if we have a valid access token
 */
async function ensureValidToken() {
    const TOKEN_REFRESH_BUFFER_MS = 10 * 60 * 1000; // 10 minutes before expiry
    const TOKEN_MINIMUM_VALIDITY_MS = 2 * 60 * 1000; // 2 minutes minimum validity
    
    try {
        // Check if we have basic token information
        if (!deviceCredentials.accessToken || !deviceCredentials.tokenExpiry) {
            // Clear any partial token data to ensure clean state
            clearTokenCredentials();
            return false;
        }
        
        const now = Date.now();
        const timeUntilExpiry = deviceCredentials.tokenExpiry - now;
        
        // Check if token is already expired
        if (timeUntilExpiry <= 0) {
            process.stderr.write('⚠️  Access token has expired\n');
            
            // Try to refresh if we have a refresh token
            if (deviceCredentials.refreshToken) {
                const refreshSuccess = await refreshAccessToken();
                if (refreshSuccess) {
                    return true;
                } else {
                    // Refresh failed - clear all tokens
                    clearTokenCredentials();
                    return false;
                }
            } else {
                // No refresh token - clear credentials and require re-authentication
                clearTokenCredentials();
                return false;
            }
        }
        
        // Check if token needs proactive refresh (expires within buffer time)
        if (timeUntilExpiry <= TOKEN_REFRESH_BUFFER_MS) {
            process.stderr.write(`🔄 Access token expires in ${Math.ceil(timeUntilExpiry / 60000)} minutes, attempting refresh...\n`);
            
            // Try to refresh the token proactively
            if (deviceCredentials.refreshToken) {
                const refreshSuccess = await refreshAccessToken();
                if (refreshSuccess) {
                    return true;
                } else {
                    // Refresh failed but token is still valid for a bit
                    if (timeUntilExpiry > TOKEN_MINIMUM_VALIDITY_MS) {
                        process.stderr.write('⚠️  Token refresh failed, but current token is still valid\n');
                        return true; // Use current token for now
                    } else {
                        // Token expires too soon and refresh failed
                        clearTokenCredentials();
                        return false;
                    }
                }
            } else {
                // No refresh token but current token is still valid
                if (timeUntilExpiry > TOKEN_MINIMUM_VALIDITY_MS) {
                    process.stderr.write('⚠️  No refresh token available, using current token\n');
                    return true;
                } else {
                    clearTokenCredentials();
                    return false;
                }
            }
        }
        
        // Token is valid and doesn't need refresh yet
        return true;
        
    } catch (error) {
        process.stderr.write(`⚠️  Error during token validation: ${error.message}\n`);
        
        // On any error, clear tokens to ensure clean state
        clearTokenCredentials();
        return false;
    }
}

/**
 * Securely clear all token credentials from memory
 */
function clearTokenCredentials() {
    // Overwrite sensitive data before clearing (basic security measure)
    if (deviceCredentials.accessToken) {
        deviceCredentials.accessToken = null;
    }
    if (deviceCredentials.refreshToken) {
        deviceCredentials.refreshToken = null;
    }
    
    // Clear all token-related fields
    deviceCredentials.accessToken = null;
    deviceCredentials.refreshToken = null;
    deviceCredentials.tokenExpiry = null;
    
    // Note: Keep deviceId and deviceSecret as they're needed for re-authentication
    // deviceCredentials.deviceId and deviceCredentials.deviceSecret remain intact
}

/**
 * Get token status information for debugging/monitoring
 * @returns {Object} - Token status information (no sensitive data)
 */
function getTokenStatus() {
    const now = Date.now();
    const hasAccessToken = !!deviceCredentials.accessToken;
    const hasRefreshToken = !!deviceCredentials.refreshToken;
    const hasDeviceCredentials = !!(deviceCredentials.deviceId && deviceCredentials.deviceSecret);
    
    let timeUntilExpiry = null;
    let isExpired = null;
    let needsRefresh = null;
    
    if (deviceCredentials.tokenExpiry) {
        timeUntilExpiry = deviceCredentials.tokenExpiry - now;
        isExpired = timeUntilExpiry <= 0;
        needsRefresh = timeUntilExpiry <= (10 * 60 * 1000); // 10 minutes
    }
    
    return {
        hasAccessToken,
        hasRefreshToken,
        hasDeviceCredentials,
        isExpired,
        needsRefresh,
        timeUntilExpiryMs: timeUntilExpiry,
        timeUntilExpiryMinutes: timeUntilExpiry ? Math.ceil(timeUntilExpiry / 60000) : null
    };
}

/**
 * Token refresh scheduler to proactively refresh tokens
 */
let tokenRefreshInterval = null;

/**
 * Start the token refresh scheduler
 */
function startTokenRefreshScheduler() {
    // Clear any existing interval
    if (tokenRefreshInterval) {
        clearInterval(tokenRefreshInterval);
    }
    
    // Check token status every 5 minutes
    const CHECK_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
    
    tokenRefreshInterval = setInterval(async () => {
        try {
            const tokenStatus = getTokenStatus();
            
            // Only attempt refresh if we have tokens and they need refreshing
            if (tokenStatus.hasAccessToken && tokenStatus.needsRefresh && !tokenStatus.isExpired) {
                process.stderr.write('🔄 Scheduled token refresh triggered\n');
                
                await ensureValidToken();
            }
            
            // Clear expired tokens that couldn't be refreshed
            if (tokenStatus.isExpired && tokenStatus.hasAccessToken) {
                process.stderr.write('⚠️  Clearing expired tokens from scheduler\n');
                clearTokenCredentials();
            }
            
        } catch (error) {
            // Silent error handling for scheduler - don't spam logs
            // Only log if it's a critical issue
            if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                process.stderr.write('⚠️  Token refresh scheduler: Connection to server lost\n');
            }
        }
    }, CHECK_INTERVAL_MS);
}

/**
 * Stop the token refresh scheduler
 */
function stopTokenRefreshScheduler() {
    if (tokenRefreshInterval) {
        clearInterval(tokenRefreshInterval);
        tokenRefreshInterval = null;
    }
}

/**
 * Perform complete device authentication flow
 * @returns {Promise<boolean>} - True if authentication was successful
 */
async function authenticateDevice() {
    // In simple mode, we already have the bearer token - no OAuth flow needed
    if (SIMPLE_MODE) {
        process.stderr.write('✅ Using bearer token authentication (simple mode)\n');
        return true;
    }
    
    const MAX_AUTH_ATTEMPTS = 2;
    
    for (let attempt = 1; attempt <= MAX_AUTH_ATTEMPTS; attempt++) {
        try {
            // Check if we already have a valid token
            const hasValidToken = await ensureValidToken();
            if (hasValidToken) {
                process.stderr.write('✅ Using existing valid authentication token\n');
                return true;
            }

            if (attempt > 1) {
                process.stderr.write(`\n🔄 Authentication attempt ${attempt}/${MAX_AUTH_ATTEMPTS}...\n`);
            }

            // Register the device
            const deviceInfo = await registerDevice();
            
            // Display user code to user
            displayUserCode(deviceInfo);
            
            // Poll for tokens
            await pollForTokens(deviceInfo.deviceCode, deviceInfo.interval, deviceInfo.expiresIn);
            
            process.stderr.write('🎉 Device authentication completed successfully!\n\n');
            return true;
            
        } catch (error) {
            const isLastAttempt = attempt === MAX_AUTH_ATTEMPTS;
            
            // Categorize authentication errors for better user feedback
            if (error.message.includes('Unable to connect to MCP server') || 
                error.message.includes('ECONNREFUSED') || 
                error.message.includes('ENOTFOUND')) {
                
                process.stderr.write(`❌ Authentication failed (attempt ${attempt}): Cannot connect to MCP server\n`);
                process.stderr.write(`   Please check your network connection and server URL: ${MCP_SERVER_URL}\n`);
                
                if (isLastAttempt) {
                    process.stderr.write('   The adapter will continue without authentication (limited functionality)\n\n');
                }
                
            } else if (error.message.includes('OAuth discovery failed')) {
                process.stderr.write(`❌ Authentication failed (attempt ${attempt}): Server does not support OAuth authentication\n`);
                
                if (isLastAttempt) {
                    process.stderr.write('   The adapter will continue without authentication (limited functionality)\n\n');
                }
                
            } else if (error.message.includes('timed out') || error.message.includes('timeout')) {
                process.stderr.write(`❌ Authentication failed (attempt ${attempt}): Request timed out\n`);
                process.stderr.write('   The MCP server may be overloaded or unreachable\n');
                
                if (isLastAttempt) {
                    process.stderr.write('   The adapter will continue without authentication (limited functionality)\n\n');
                }
                
            } else if (error.message.includes('Authentication was denied by the user')) {
                process.stderr.write(`❌ Authentication failed: User denied authorization\n`);
                process.stderr.write('   The adapter will continue without authentication (limited functionality)\n\n');
                return false; // Don't retry if user explicitly denied
                
            } else if (error.message.includes('Device code expired')) {
                process.stderr.write(`❌ Authentication failed (attempt ${attempt}): Device code expired\n`);
                process.stderr.write('   Please complete the authentication process more quickly\n');
                
                if (!isLastAttempt) {
                    process.stderr.write('   Retrying with a new device code...\n');
                }
                
            } else {
                process.stderr.write(`❌ Authentication failed (attempt ${attempt}): ${error.message}\n`);
            }
            
            if (isLastAttempt) {
                // Final attempt failed - provide detailed error
                if (error.message.includes('ECONNREFUSED') || error.message.includes('ENOTFOUND')) {
                    throw new Error('Device authentication failed: Unable to connect to MCP server. Please check your network connection and server URL.');
                } else if (error.message.includes('timeout')) {
                    throw new Error('Device authentication failed: Request timed out. The MCP server may be overloaded or unreachable.');
                } else {
                    throw new Error(`Device authentication failed after ${MAX_AUTH_ATTEMPTS} attempts: ${error.message}`);
                }
            }
            
            // Wait before retrying (progressive backoff)
            const backoffDelay = 2000 * attempt; // 2s, 4s
            process.stderr.write(`   Retrying in ${backoffDelay / 1000} seconds...\n`);
            await new Promise(resolve => setTimeout(resolve, backoffDelay));
        }
    }
    
    return false;
}

/**
 * Initialize token management system
 */
function initializeTokenManagement() {
    // Validate existing token storage
    validateTokenStorage();
    
    // Start the token refresh scheduler if we have tokens
    const tokenStatus = getTokenStatus();
    if (tokenStatus.hasAccessToken) {
        startTokenRefreshScheduler();
        process.stderr.write('✅ Token refresh scheduler started\n');
    }
}

/**
 * Cleanup token management system
 */
function cleanupTokenManagement() {
    stopTokenRefreshScheduler();
    clearTokenCredentials();
}

/**
 * Enhanced token validation with automatic re-authentication fallback
 * @param {boolean} allowReauth - Whether to attempt re-authentication if tokens are invalid
 * @returns {Promise<boolean>} - True if we have a valid access token
 */
async function ensureValidTokenWithReauth(allowReauth = true) {
    const hasValidToken = await ensureValidToken();
    
    if (!hasValidToken && allowReauth) {
        // Attempt re-authentication if we don't have valid tokens
        process.stderr.write('🔄 No valid tokens available, attempting re-authentication...\n');
        
        const authSuccess = await authenticateDevice();
        if (authSuccess) {
            // Restart the token refresh scheduler after successful authentication
            startTokenRefreshScheduler();
            return true;
        } else {
            process.stderr.write('❌ Re-authentication failed\n');
            return false;
        }
    }
    
    return hasValidToken;
}

/**
 * Secure token storage validation
 * @returns {boolean} - True if token storage is in a valid state
 */
function validateTokenStorage() {
    try {
        // Check for token storage consistency
        const hasAccessToken = !!deviceCredentials.accessToken;
        const hasTokenExpiry = !!deviceCredentials.tokenExpiry;
        const hasRefreshToken = !!deviceCredentials.refreshToken;
        
        // If we have an access token, we should have an expiry time
        if (hasAccessToken && !hasTokenExpiry) {
            process.stderr.write('⚠️  Invalid token storage: access token without expiry\n');
            clearTokenCredentials();
            return false;
        }
        
        // If we have an expiry time, we should have an access token
        if (hasTokenExpiry && !hasAccessToken) {
            process.stderr.write('⚠️  Invalid token storage: expiry without access token\n');
            clearTokenCredentials();
            return false;
        }
        
        // Validate expiry time format
        if (hasTokenExpiry && (typeof deviceCredentials.tokenExpiry !== 'number' || deviceCredentials.tokenExpiry <= 0)) {
            process.stderr.write('⚠️  Invalid token storage: malformed expiry time\n');
            clearTokenCredentials();
            return false;
        }
        
        return true;
        
    } catch (error) {
        process.stderr.write(`⚠️  Token storage validation error: ${error.message}\n`);
        clearTokenCredentials();
        return false;
    }
}

// Graceful shutdown handling for token management
process.on('SIGINT', () => {
    process.stderr.write('\n🔄 Gracefully shutting down token management...\n');
    cleanupTokenManagement();
    process.exit(0);
});

process.on('SIGTERM', () => {
    process.stderr.write('\n🔄 Gracefully shutting down token management...\n');
    cleanupTokenManagement();
    process.exit(0);
});

process.on('exit', () => {
    // Final cleanup on any exit
    cleanupTokenManagement();
});
