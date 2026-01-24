/**
 * @fileoverview MCP Transport Controller
 *
 * Implements MCP (Model Context Protocol) transport over HTTP/SSE.
 * This allows Claude Desktop to connect directly to the server without
 * needing a local adapter file.
 *
 * Endpoints:
 * - GET /api/mcp/sse - SSE connection for receiving server messages
 * - POST /api/mcp/message - Send client messages to server
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const apiContext = require('../api-context.cjs');
const msalService = require('../../auth/msal-service.cjs');

// MCP Protocol version
const MCP_PROTOCOL_VERSION = '2024-11-05';

// Active SSE connections (for multi-client support)
const activeConnections = new Map();

/**
 * Generate a unique session ID
 */
function generateSessionId() {
    return `mcp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Format tool definitions for MCP protocol
 */
function formatToolsForMCP(tools) {
    return tools.map(tool => ({
        name: tool.name,
        description: tool.description || `${tool.name} operation`,
        inputSchema: {
            type: 'object',
            properties: Object.entries(tool.parameters || {}).reduce((acc, [key, param]) => {
                acc[key] = {
                    type: param.type || 'string',
                    description: param.description || ''
                };
                if (param.enum) acc[key].enum = param.enum;
                if (param.default !== undefined) acc[key].default = param.default;
                return acc;
            }, {}),
            required: Object.entries(tool.parameters || {})
                .filter(([_, param]) => param.required)
                .map(([key]) => key)
        }
    }));
}

/**
 * Execute a tool call via the existing API infrastructure
 */
async function executeTool(toolName, args, req) {
    const startTime = Date.now();

    // Parse tool name (format: module.method or just method)
    let moduleName, methodName;
    if (toolName.includes('.')) {
        [moduleName, methodName] = toolName.split('.');
    } else {
        // Try to determine module from tool name
        const toolAliases = {
            // Mail tools
            getMail: { moduleName: 'mail', methodName: 'getInbox' },
            getInbox: { moduleName: 'mail', methodName: 'getInbox' },
            sendEmail: { moduleName: 'mail', methodName: 'sendEmail' },
            sendMail: { moduleName: 'mail', methodName: 'sendEmail' },
            searchEmails: { moduleName: 'mail', methodName: 'searchEmails' },
            searchMail: { moduleName: 'mail', methodName: 'searchEmails' },
            flagEmail: { moduleName: 'mail', methodName: 'flagEmail' },
            getEmailDetails: { moduleName: 'mail', methodName: 'getEmailDetails' },
            markAsRead: { moduleName: 'mail', methodName: 'markAsRead' },

            // Calendar tools
            getEvents: { moduleName: 'calendar', methodName: 'getEvents' },
            getCalendar: { moduleName: 'calendar', methodName: 'getEvents' },
            createEvent: { moduleName: 'calendar', methodName: 'create' },
            updateEvent: { moduleName: 'calendar', methodName: 'update' },
            cancelEvent: { moduleName: 'calendar', methodName: 'cancelEvent' },
            getAvailability: { moduleName: 'calendar', methodName: 'getAvailability' },
            findMeetingTimes: { moduleName: 'calendar', methodName: 'findMeetingTimes' },

            // Files tools
            listFiles: { moduleName: 'files', methodName: 'listFiles' },
            searchFiles: { moduleName: 'files', methodName: 'searchFiles' },
            downloadFile: { moduleName: 'files', methodName: 'downloadFile' },
            uploadFile: { moduleName: 'files', methodName: 'uploadFile' },
            getFileMetadata: { moduleName: 'files', methodName: 'getFileMetadata' },

            // People tools
            findPeople: { moduleName: 'people', methodName: 'find' },
            getRelevantPeople: { moduleName: 'people', methodName: 'getRelevantPeople' }
        };

        const alias = toolAliases[toolName];
        if (alias) {
            moduleName = alias.moduleName;
            methodName = alias.methodName;
        } else {
            throw new Error(`Unknown tool: ${toolName}`);
        }
    }

    MonitoringService.info('Executing MCP tool via SSE transport', {
        toolName,
        moduleName,
        methodName,
        timestamp: new Date().toISOString()
    }, 'mcp-transport');

    // Get the module from registry
    const module = apiContext.moduleRegistry.getModule(moduleName);
    if (!module) {
        throw new Error(`Module not found: ${moduleName}`);
    }

    // Get Microsoft Graph access token
    // For SSE connections with Bearer token, we need to fetch the Graph token
    // using the user's identity from the JWT
    let accessToken = null;

    // Try session-based token first (browser requests)
    if (req.session?.msUser?.accessToken) {
        accessToken = req.session.msUser.accessToken;
    }

    // If no session token, get it via MSAL service using the user identity
    if (!accessToken && req.user?.userId) {
        try {
            // The MSAL service will retrieve the token for the authenticated user
            accessToken = await msalService.getAccessToken(req);
        } catch (tokenError) {
            MonitoringService.error('Failed to get access token for MCP tool execution', {
                userId: req.user.userId,
                error: tokenError.message,
                timestamp: new Date().toISOString()
            }, 'mcp-transport');
        }
    }

    if (!accessToken) {
        throw new Error('No valid Microsoft Graph access token available. Please re-authenticate.');
    }

    // Execute the module method
    if (typeof module[methodName] !== 'function') {
        throw new Error(`Method not found: ${moduleName}.${methodName}`);
    }

    // Call the handler with args, access token, and request object
    // Module methods expect (options, req) signature for proper authentication context
    // IMPORTANT: Use .call() to preserve 'this' binding for module methods
    const result = await module[methodName].call(module, { ...args, accessToken }, req);

    MonitoringService.info('MCP tool executed successfully', {
        toolName,
        duration: Date.now() - startTime,
        timestamp: new Date().toISOString()
    }, 'mcp-transport');

    return result;
}

/**
 * Handle MCP JSON-RPC message
 */
async function handleMCPMessage(message, req, sessionId) {
    const { jsonrpc, id, method, params } = message;

    // Validate JSON-RPC format
    if (jsonrpc !== '2.0') {
        return {
            jsonrpc: '2.0',
            id,
            error: {
                code: -32600,
                message: 'Invalid Request: must be JSON-RPC 2.0'
            }
        };
    }

    MonitoringService.debug('Processing MCP message', {
        method,
        sessionId,
        timestamp: new Date().toISOString()
    }, 'mcp-transport');

    try {
        switch (method) {
            case 'initialize': {
                // Initialize handshake
                return {
                    jsonrpc: '2.0',
                    id,
                    result: {
                        protocolVersion: MCP_PROTOCOL_VERSION,
                        capabilities: {
                            tools: { listChanged: false }
                        },
                        serverInfo: {
                            name: 'mcp-microsoft-365',
                            version: '2.5.0'
                        }
                    }
                };
            }

            case 'initialized': {
                // Client acknowledges initialization
                return {
                    jsonrpc: '2.0',
                    id,
                    result: {}
                };
            }

            case 'tools/list': {
                // Return available tools
                const tools = apiContext.toolsService.getAllTools();
                const formattedTools = formatToolsForMCP(tools);

                return {
                    jsonrpc: '2.0',
                    id,
                    result: {
                        tools: formattedTools
                    }
                };
            }

            case 'tools/call': {
                // Execute a tool
                const { name, arguments: toolArgs } = params || {};

                if (!name) {
                    return {
                        jsonrpc: '2.0',
                        id,
                        error: {
                            code: -32602,
                            message: 'Invalid params: tool name is required'
                        }
                    };
                }

                try {
                    const result = await executeTool(name, toolArgs || {}, req);

                    return {
                        jsonrpc: '2.0',
                        id,
                        result: {
                            content: [
                                {
                                    type: 'text',
                                    text: typeof result === 'string' ? result : JSON.stringify(result, null, 2)
                                }
                            ],
                            isError: false
                        }
                    };
                } catch (toolError) {
                    MonitoringService.error('MCP tool execution failed', {
                        toolName: name,
                        error: toolError.message,
                        timestamp: new Date().toISOString()
                    }, 'mcp-transport');

                    return {
                        jsonrpc: '2.0',
                        id,
                        result: {
                            content: [
                                {
                                    type: 'text',
                                    text: `Error: ${toolError.message}`
                                }
                            ],
                            isError: true
                        }
                    };
                }
            }

            case 'ping': {
                return {
                    jsonrpc: '2.0',
                    id,
                    result: {}
                };
            }

            default: {
                return {
                    jsonrpc: '2.0',
                    id,
                    error: {
                        code: -32601,
                        message: `Method not found: ${method}`
                    }
                };
            }
        }
    } catch (error) {
        MonitoringService.error('MCP message processing error', {
            method,
            error: error.message,
            stack: error.stack,
            timestamp: new Date().toISOString()
        }, 'mcp-transport');

        return {
            jsonrpc: '2.0',
            id,
            error: {
                code: -32603,
                message: `Internal error: ${error.message}`
            }
        };
    }
}

/**
 * SSE endpoint for receiving server messages
 * GET /api/mcp/sse
 *
 * Supports authentication via:
 * - Authorization header (Bearer token)
 * - Query parameter: ?token=xxx
 */
async function sseConnect(req, res) {
    const sessionId = generateSessionId();

    MonitoringService.info('MCP SSE connection established', {
        sessionId,
        userAgent: req.get('User-Agent'),
        timestamp: new Date().toISOString()
    }, 'mcp-transport');

    // Set SSE headers
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.setHeader('X-Accel-Buffering', 'no'); // Disable nginx buffering

    // Send initial connection message with session ID
    const endpoint = `${req.protocol}://${req.get('host')}/api/mcp/message?sessionId=${sessionId}`;
    res.write(`event: endpoint\ndata: ${endpoint}\n\n`);

    // Store connection
    activeConnections.set(sessionId, { res, req, createdAt: Date.now() });

    // Send keepalive ping every 30 seconds
    const keepaliveInterval = setInterval(() => {
        if (!res.writableEnded) {
            res.write(`: keepalive\n\n`);
        }
    }, 30000);

    // Handle client disconnect
    req.on('close', () => {
        clearInterval(keepaliveInterval);
        activeConnections.delete(sessionId);

        MonitoringService.info('MCP SSE connection closed', {
            sessionId,
            duration: Date.now() - activeConnections.get(sessionId)?.createdAt || 0,
            timestamp: new Date().toISOString()
        }, 'mcp-transport');
    });
}

/**
 * Message endpoint for receiving client messages
 * POST /api/mcp/message
 */
async function handleMessage(req, res) {
    const sessionId = req.query.sessionId || req.body.sessionId;

    try {
        const message = req.body;

        if (!message || typeof message !== 'object') {
            return res.status(400).json({
                jsonrpc: '2.0',
                id: null,
                error: {
                    code: -32700,
                    message: 'Parse error: invalid JSON'
                }
            });
        }

        // Handle the MCP message
        const response = await handleMCPMessage(message, req, sessionId);

        // If there's an active SSE connection, also send via SSE
        const connection = activeConnections.get(sessionId);
        if (connection && !connection.res.writableEnded) {
            connection.res.write(`event: message\ndata: ${JSON.stringify(response)}\n\n`);
        }

        // Always return the response via HTTP too
        res.json(response);

    } catch (error) {
        MonitoringService.error('MCP message handler error', {
            sessionId,
            error: error.message,
            stack: error.stack,
            timestamp: new Date().toISOString()
        }, 'mcp-transport');

        res.status(500).json({
            jsonrpc: '2.0',
            id: req.body?.id || null,
            error: {
                code: -32603,
                message: 'Internal error'
            }
        });
    }
}

/**
 * Simple HTTP endpoint for MCP (alternative to SSE)
 * POST /api/mcp
 *
 * This provides a simpler integration path that doesn't require SSE.
 * Each request/response is a complete MCP transaction.
 */
async function handleSimpleMessage(req, res) {
    try {
        const message = req.body;

        if (!message || typeof message !== 'object') {
            return res.status(400).json({
                jsonrpc: '2.0',
                id: null,
                error: {
                    code: -32700,
                    message: 'Parse error: invalid JSON'
                }
            });
        }

        // Handle the MCP message
        const response = await handleMCPMessage(message, req, 'simple');
        res.json(response);

    } catch (error) {
        MonitoringService.error('MCP simple message handler error', {
            error: error.message,
            stack: error.stack,
            timestamp: new Date().toISOString()
        }, 'mcp-transport');

        res.status(500).json({
            jsonrpc: '2.0',
            id: req.body?.id || null,
            error: {
                code: -32603,
                message: 'Internal error'
            }
        });
    }
}

/**
 * Get server capabilities and info
 * GET /api/mcp/info
 */
async function getInfo(req, res) {
    try {
        const tools = apiContext.toolsService.getAllTools();

        res.json({
            name: 'mcp-microsoft-365',
            version: '2.5.0',
            protocolVersion: MCP_PROTOCOL_VERSION,
            capabilities: {
                tools: {
                    available: true,
                    count: tools.length
                }
            },
            endpoints: {
                sse: '/api/mcp/sse',
                message: '/api/mcp/message',
                simple: '/api/mcp'
            }
        });
    } catch (error) {
        res.status(500).json({
            error: 'INTERNAL_ERROR',
            message: 'Failed to get server info'
        });
    }
}

module.exports = {
    sseConnect,
    handleMessage,
    handleSimpleMessage,
    getInfo,
    MCP_PROTOCOL_VERSION
};
