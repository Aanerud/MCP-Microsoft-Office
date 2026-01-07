/**
 * @fileoverview Handles Microsoft Graph authentication for MCP backend.
 * Provides status checks and simulated login for demo/dev.
 */

const msal = require('@azure/msal-node');
const url = require('url');
const crypto = require('crypto');
const storageService = require('../core/storage-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

// External token support - lazy loaded to avoid circular dependency
let externalTokenController = null;
function getExternalTokenController() {
    if (!externalTokenController) {
        externalTokenController = require('../api/controllers/external-token-controller.cjs');
    }
    return externalTokenController;
}

// Load environment variables
const CLIENT_ID = process.env.MICROSOFT_CLIENT_ID;
const TENANT_ID = process.env.MICROSOFT_TENANT_ID || 'common';
const REDIRECT_URI = process.env.MICROSOFT_REDIRECT_URI || 'http://localhost:3000/api/auth/callback';
// Use scopes that match the permissions granted in Azure AD
const SCOPES = [
    'User.Read',           // Sign in and read user profile
    'openid',
    'profile',
    'email',
    'Calendars.ReadWrite', // Full access to user calendars
    'Mail.ReadWrite',      // Read and write access to user mail
    'Mail.Send',           // Send mail as a user
    'Files.ReadWrite'      // Full access to user files
];

// Debug environment variables (only in development)
if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('MSAL environment variables', {
        clientIdSet: CLIENT_ID ? 'Set' : 'Not set',
        tenantId: TENANT_ID,
        redirectUri: REDIRECT_URI,
        timestamp: new Date().toISOString()
    }, 'auth');
}

// Initialize storage service
storageService.init().catch(err => {
    const mcpError = ErrorService.createError(
        ErrorService.CATEGORIES.SYSTEM,
        `MSAL storage service initialization failed: ${err.message}`,
        ErrorService.SEVERITIES.ERROR,
        { stack: err.stack, timestamp: new Date().toISOString() }
    );
    MonitoringService.logError(mcpError);
});

// Session storage for tokens and accounts - Multi-user support
const userSessions = new Map(); // Map<userId, sessionData>

/**
 * Get user session by user ID
 * @param {string} userId - User ID
 * @returns {Object|null} User session data or null
 */
function getUserSession(userId) {
    if (!userId) return null;
    return userSessions.get(userId) || null;
}

/**
 * Set user session data
 * @param {string} userId - User ID
 * @param {Object} sessionData - Session data
 */
function setUserSession(userId, sessionData) {
    if (!userId) return;
    userSessions.set(userId, sessionData);
}

/**
 * Clear user session
 * @param {string} userId - User ID
 */
function clearUserSession(userId) {
    if (!userId) return;
    userSessions.delete(userId);
}

const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${TENANT_ID}`,
        // NO clientSecret for public client
    },
    system: { 
        loggerOptions: { 
            loggerCallback(level, message) {
                // Map MSAL log levels to appropriate MCP severities
                let severity;
                let logMethod;
                
                switch (level) {
                    case 'Error':
                        severity = ErrorService.SEVERITIES.ERROR;
                        logMethod = 'logError';
                        break;
                    case 'Warning':
                        severity = ErrorService.SEVERITIES.WARNING;
                        logMethod = 'warn';
                        break;
                    case 'Info':
                    default:
                        severity = ErrorService.SEVERITIES.INFO;
                        logMethod = 'info';
                        break;
                }
                
                if (logMethod === 'logError') {
                    // For errors, create an error object and use logError
                    const mcpError = ErrorService.createError(
                        ErrorService.CATEGORIES.AUTH,
                        `MSAL library message: ${message}`,
                        severity,
                        { level, timestamp: new Date().toISOString() }
                    );
                    MonitoringService.logError(mcpError);
                } else {
                    // For info and warnings, use direct logging methods
                    MonitoringService[logMethod](
                        `MSAL library message: ${message}`,
                        { level, timestamp: new Date().toISOString() },
                        'auth'
                    );
                }
            } 
        } 
    }
};

// Verify MSAL config (only in development)
if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('MSAL configuration', {
        authority: msalConfig.auth.authority,
        clientIdSet: !!msalConfig.auth.clientId,
        timestamp: new Date().toISOString()
    }, 'auth');
}

const pca = new msal.PublicClientApplication(msalConfig);

// Generate PKCE code verifier and code challenge
function generatePkceCodes() {
    const codeVerifier = base64URLEncode(crypto.randomBytes(32));
    const codeChallenge = base64URLEncode(sha256(codeVerifier));
    return { codeVerifier, codeChallenge };
}

function sha256(buffer) {
    return crypto.createHash('sha256').update(buffer).digest();
}

function base64URLEncode(buffer) {
    return buffer.toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');
}

/**
 * Extract user context from request for logging purposes
 * @param {Object} req - Express request object
 * @returns {Object} User context with userId and sessionId
 */
function extractUserContext(req) {
    let userId, sessionId;
    
    if (req.user?.userId) {
        userId = req.user.userId;
        sessionId = req.user.sessionId || req.session?.id;
    } else if (req.session?.id) {
        sessionId = req.session.id;
        userId = req.session.msUser?.username ? `ms365:${req.session.msUser.username}` : `session:${sessionId}`;
    }
    
    return { userId, sessionId };
}

/**
 * Get the login URL for Microsoft authentication
 * @param {Object} req - Express request object
 */
async function getLoginUrl(req) {
    const startTime = Date.now();
    const { userId, sessionId } = extractUserContext(req);
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing getLoginUrl request', {
                sessionId,
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString(),
                userId
            }, 'auth');
        }
        
        const { codeVerifier, codeChallenge } = generatePkceCodes();
        
        // Store PKCE verifier in session (simple approach)
        req.session.pkceCodeVerifier = codeVerifier;
        
        // Force session save to ensure code verifier is persisted immediately
        await new Promise((resolve, reject) => {
            req.session.save(err => {
                if (err) {
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'auth',
                        'Failed to save session with PKCE code verifier',
                        'error',
                        {
                            error: err.message,
                            stack: err.stack,
                            sessionId,
                            userId,
                            timestamp: new Date().toISOString()
                        }
                    );
                    MonitoringService.logError(mcpError);
                    
                    // Pattern 4: User Error Tracking
                    if (userId) {
                        MonitoringService.error('Login URL generation failed', {
                            error: err.message,
                            operation: 'getLoginUrl',
                            timestamp: new Date().toISOString()
                        }, 'auth', null, userId);
                    } else if (sessionId) {
                        MonitoringService.error('Login URL generation failed', {
                            sessionId,
                            error: err.message,
                            operation: 'getLoginUrl',
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }
                    
                    reject(err);
                } else {
                    // Only log in development mode
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Session saved with PKCE code verifier', {
                            sessionId,
                            hasCodeVerifier: !!req.session.pkceCodeVerifier,
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }
                    resolve();
                }
            });
        }).catch(err => {
            // Log but continue even if save fails
            MonitoringService.error('Error saving session', { error: err.message }, 'auth');
        });
        
        const authCodeUrlParameters = {
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            codeChallenge,
            codeChallengeMethod: 'S256',
            prompt: 'select_account'
        };
        
        const authUrl = pca.getAuthCodeUrl(authCodeUrlParameters);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Login URL generated successfully', {
                operation: 'getLoginUrl',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Login URL generated with session', {
                sessionId,
                operation: 'getLoginUrl',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        return authUrl;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            'Failed to generate login URL',
            'error',
            {
                error: error.message,
                stack: error.stack,
                operation: 'getLoginUrl',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Login URL generation failed', {
                error: error.message,
                operation: 'getLoginUrl',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Login URL generation failed', {
                sessionId,
                error: error.message,
                operation: 'getLoginUrl',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        throw error;
    }
}

/**
 * Handle login request
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 */
async function login(req, res) {
    const startTime = Date.now();
    const { userId, sessionId } = extractUserContext(req);
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing login request', {
                hasSession: !!req.session,
                sessionId,
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString(),
                userId
            }, 'auth');
        }
        
        // Check if session is available
        if (!req.session) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                'Session middleware not available',
                'error',
                {
                    endpoint: '/api/auth/login',
                    userId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Login failed - session not available', {
                    error: 'Session middleware not available',
                    operation: 'login',
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);
            }
            
            return res.status(500).json({ 
                error: 'session_not_available',
                error_description: 'Session not available' 
            });
        }
        
        const authUrl = await getLoginUrl(req);
        
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Generated auth URL for login', {
                authUrlLength: authUrl.length,
                sessionId,
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Set CORS headers for the redirect
        res.header('Access-Control-Allow-Origin', '*');
        res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Login redirect initiated successfully', {
                operation: 'login',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Login redirect initiated with session', {
                sessionId,
                operation: 'login',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        res.redirect(authUrl);
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `MSAL login error: ${err.message}`,
            'error',
            {
                stack: err.stack,
                operation: 'login',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Login failed', {
                error: err.message,
                operation: 'login',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Login failed', {
                sessionId,
                error: err.message,
                operation: 'login',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        res.status(500).json({
            error: 'login_failed',
            error_description: 'Failed to initiate login process'
        });
    }
}

/**
 * Handle the OAuth callback
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 */
async function handleAuthCallback(req, res) {
    const startTime = Date.now();
    const { userId, sessionId } = extractUserContext(req);
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing auth callback request', {
                clientIdSet: !!CLIENT_ID,
                tenantId: TENANT_ID,
                redirectUri: REDIRECT_URI,
                sessionId,
                hasSession: !!req.session,
                hasCodeVerifier: !!req.session?.pkceCodeVerifier,
                environment: process.env.NODE_ENV || 'development',
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString(),
                userId
            }, 'auth');
        }

        
        // Get code verifier from session
        const codeVerifier = req.session.pkceCodeVerifier;
        if (!codeVerifier) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                'No PKCE codeVerifier found',
                'error',
                {
                    sessionId,
                    userId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Auth callback failed - no code verifier', {
                    error: 'No PKCE codeVerifier found',
                    operation: 'handleAuthCallback',
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Auth callback failed - no code verifier', {
                    sessionId,
                    error: 'No PKCE codeVerifier found',
                    operation: 'handleAuthCallback',
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            return res.status(400).json({
                error: 'authentication_failed',
                error_description: 'Authentication failed: No code verifier found. Please try logging in again.'
            });
        }
        
        const tokenRequest = {
            code: req.query.code,
            scopes: SCOPES,
            redirectUri: REDIRECT_URI,
            codeVerifier: codeVerifier
        };
        
        const response = await pca.acquireTokenByCode(tokenRequest);
        
        // Store user info in session, memory, and SQLite database
        const userInfo = {
            username: response.account.username,
            name: response.account.name,
            homeAccountId: response.account.homeAccountId,
            accessToken: response.accessToken,
            expiresOn: response.expiresOn,
            account: response.account
        };
        
        // Store in session if available
        if (req.session) {
            req.session.msUser = userInfo;
        }

        // Store in user session map using ms365:email as the key
        // This ensures MCP adapter lookups (which use ms365:email) can find the token
        const memoryKey = `ms365:${userInfo.username}`;
        setUserSession(memoryKey, { msUser: userInfo });
        
        // Clean up temporary session
        delete req.session.pkceCodeVerifier;
        
        // Also store in SQLite database for persistence across restarts
        try {
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Storing authentication token in database', {
                    username: userInfo.username,
                    sessionId,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            // MICROSOFT 365-CENTRIC AUTH: Store token using Microsoft 365 email as user identifier
            // CRITICAL FIX: Use ms365:email as the userId for storage, NOT the browser session ID
            // This ensures MCP adapter API calls can retrieve the token using the same userId
            const userKey = `ms365:${userInfo.username}`;
            await storageService.setSecureSetting(`${userKey}:ms-access-token`, userInfo.accessToken, userKey);
            await storageService.setSetting(`${userKey}:ms-user-info`, {
                username: userInfo.username,
                name: userInfo.name,
                homeAccountId: userInfo.homeAccountId,
                expiresOn: userInfo.expiresOn
            }, userKey);
            
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Authentication token stored successfully', {
                    username: userInfo.username,
                    sessionId,
                    userKey: userKey,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
        } catch (dbError) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                `Error storing token in database: ${dbError.message}`,
                'warning',
                {
                    stack: dbError.stack,
                    username: userInfo.username,
                    sessionId,
                    userId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            // Continue even if database storage fails
        }
        
        // Pattern 2: User Activity Logs
        const finalUserId = `ms365:${userInfo.username}`;
        MonitoringService.info('Authentication completed successfully', {
            username: userInfo.username,
            operation: 'handleAuthCallback',
            duration: Date.now() - startTime,
            timestamp: new Date().toISOString()
        }, 'auth', null, finalUserId);
        
        res.redirect('/');
        
    } catch (err) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `Authentication callback error: ${err.message}`,
            'error',
            {
                stack: err.stack,
                operation: 'handleAuthCallback',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Auth callback failed', {
                error: err.message,
                operation: 'handleAuthCallback',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Auth callback failed', {
                sessionId,
                error: err.message,
                operation: 'handleAuthCallback',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        res.status(500).json({
            error: 'authentication_failed',
            error_description: 'Authentication failed during callback processing'
        });
    }
}

/**
 * Check if the user is authenticated
 * @param {Object} req - Express request object
 * @returns {Promise<boolean>} - True if authenticated
 */
async function isAuthenticated(req) {
    const startTime = Date.now();
    
    try {
        // Handle both session-based auth (browser) and device-based auth (MCP adapter)
        let userId, sessionId;
        
        if (req.user?.isApiCall && req.user?.userId) {
            // Device auth flow - use Microsoft 365-based userId from JWT token
            userId = req.user.userId;  // This should be ms365:email@domain.com
            sessionId = req.user.sessionId || req.user.userId;
        } else if (req.session?.id && req.session?.msUser?.username) {
            // Session-based auth flow - use Microsoft 365 email as consistent identifier
            sessionId = req.session.id;
            userId = `ms365:${req.session.msUser.username}`;
        } else if (req.session?.id) {
            // Fallback for sessions without Microsoft 365 auth
            sessionId = req.session.id;
            userId = `session:${sessionId}`;
        }
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing authentication check', {
                hasSession: !!req.session,
                sessionId: sessionId,
                hasUser: !!req.user,
                isApiCall: req.user?.isApiCall,
                userId: userId,
                hasMsUser: !!req.session?.msUser,
                hasAccessToken: !!req.session?.msUser?.accessToken,
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Check if user is authenticated via Express session (primary method for browser)
        if (req.session?.msUser?.accessToken) {
            // Pattern 2: User Activity Logs
            const finalUserId = `ms365:${req.session.msUser.username}`;
            MonitoringService.info('User authenticated via Express session', {
                username: req.session.msUser.username,
                operation: 'isAuthenticated',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, finalUserId);
            return true;
        }
        
        // Check database storage using userId (works for both session and device auth)
        if (userId) {
            try {
                const tokenKey = `${userId}:ms-access-token`;
                const storedToken = await storageService.getSecureSetting(tokenKey, sessionId);
                if (storedToken) {
                    // Pattern 2: User Activity Logs
                    if (userId) {
                        MonitoringService.info('User authenticated via database', {
                            isApiCall: req.user?.isApiCall,
                            operation: 'isAuthenticated',
                            duration: Date.now() - startTime,
                            timestamp: new Date().toISOString()
                        }, 'auth', null, userId);
                    } else if (sessionId) {
                        MonitoringService.info('User authenticated via database with session', {
                            sessionId,
                            isApiCall: req.user?.isApiCall,
                            operation: 'isAuthenticated',
                            duration: Date.now() - startTime,
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }
                    return true;
                }
            } catch (error) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'auth',
                    'Failed to check database authentication',
                    'warning',
                    {
                        userId: userId,
                        sessionId: sessionId,
                        error: error.message,
                        stack: error.stack,
                        operation: 'isAuthenticated',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Database authentication check failed', {
                        userId: userId,
                        sessionId: sessionId,
                        error: error.message,
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }
            }
        }
        
        // Pattern 2: User Activity Logs (for not authenticated case)
        if (userId) {
            MonitoringService.info('User not authenticated', {
                operation: 'isAuthenticated',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.info('User not authenticated with session', {
                sessionId,
                operation: 'isAuthenticated',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        return false;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            'Error during authentication check',
            'error',
            {
                error: error.message,
                stack: error.stack,
                operation: 'isAuthenticated',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        const { userId, sessionId } = extractUserContext(req);
        if (userId) {
            MonitoringService.error('Authentication check failed', {
                error: error.message,
                operation: 'isAuthenticated',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Authentication check failed', {
                sessionId,
                error: error.message,
                operation: 'isAuthenticated',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        return false;
    }
}

/**
 * Get an access token for Microsoft Graph API
 * @param {Object} req - Express request object (optional)
 * @returns {Promise<string>} The access token
 */
async function getAccessToken(req) {
    const startTime = Date.now();
    
    try {
        // Handle both session-based auth (browser) and device-based auth (MCP adapter)
        let userId, sessionId;
        
        if (req.user?.isApiCall && req.user?.userId) {
            // Device auth flow - use userId from JWT token
            userId = req.user.userId;
            sessionId = req.user.sessionId || req.user.userId; // Use userId as sessionId if not provided
        } else if (req.session?.id && req.session?.msUser?.username) {
            // Session-based auth flow with MS365 user - use Microsoft 365 email as consistent identifier
            sessionId = req.session.id;
            userId = `ms365:${req.session.msUser.username}`;
        } else if (req.session?.id) {
            // Session-based auth flow without MS365 user - fallback to session ID
            sessionId = req.session.id;
            userId = `session:${sessionId}`;
        } else {
            // Fallback to query parameter
            userId = req.query?.userId;
            sessionId = userId;
        }
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing getAccessToken request', {
                hasReq: !!req,
                hasUser: !!req?.user,
                hasSession: !!req?.session,
                sessionId,
                extractedUserId: userId,
                isApiCall: req.user?.isApiCall,
                userAgent: req.get ? req.get('User-Agent') : 'N/A',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        if (!userId) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                'No user ID available for token retrieval',
                'error',
                {
                    operation: 'getAccessToken',
                    hasReq: !!req,
                    hasUser: !!req?.user,
                    hasSession: !!req?.session,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            throw new Error('No user ID available for token retrieval');
        }

        // Check for external token first (enterprise tokens from external tools)
        try {
            const extController = getExternalTokenController();
            const externalToken = await extController.getActiveExternalToken(userId);
            if (externalToken) {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Using external enterprise token', {
                        userId: userId.substring(0, 8) + '...',
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }

                // Pattern 2: User Activity Logs
                MonitoringService.info('Access token retrieved from external source', {
                    operation: 'getAccessToken',
                    source: 'external_token',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);

                return externalToken;
            }
        } catch (extError) {
            // External token check failed - continue with normal flow
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('External token check failed, continuing with normal flow', {
                    error: extError.message,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
        }

        const userSession = getUserSession(userId);
        if (userSession?.msUser?.accessToken) {
            // TODO: Check token expiration and refresh if needed
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Using access token from user session', {
                    userId: userId.substring(0, 8) + '...',
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Access token retrieved from session', {
                    operation: 'getAccessToken',
                    source: 'user_session',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);
            }
            
            return userSession.msUser.accessToken;
        }
        
        // If not in memory, try to get from SQLite database using session-based key
        if (userId) {
            try {
                const tokenKey = `${userId}:ms-access-token`;
                const storedToken = await storageService.getSecureSetting(tokenKey, sessionId);
                if (storedToken) {
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Using access token from SQLite database', {
                            userId: userId.substring(0, 8) + '...',
                            tokenKey,
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }
                    
                    // Also load it into memory for future use
                    const userInfoKey = `${userId}:ms-user-info`;
                    const userInfo = await storageService.getSetting(userInfoKey, sessionId) || {};
                    setUserSession(userId, {
                        msUser: {
                            ...userInfo,
                            accessToken: storedToken
                        }
                    });
                    
                    // Pattern 2: User Activity Logs
                    MonitoringService.info('Access token retrieved from database', {
                        operation: 'getAccessToken',
                        source: 'database',
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'auth', null, userId);
                    
                    return storedToken;
                }
            } catch (dbError) {
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'auth',
                    `Error getting token from database: ${dbError.message}`,
                    'warning',
                    {
                        stack: dbError.stack,
                        userId,
                        sessionId,
                        operation: 'getAccessToken',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
            }
        }
        
        // If we have an account, try to get a token silently
        if (userSession?.msUser?.account) {
            const silentRequest = {
                account: userSession.msUser.account,
                scopes: SCOPES
            };
            
            try {
                const response = await pca.acquireTokenSilent(silentRequest);
                if (response && response.accessToken) {
                    // Update the token in session
                    setUserSession(userId, {
                        msUser: {
                            ...userSession.msUser,
                            accessToken: response.accessToken,
                            expiresOn: response.expiresOn
                        }
                    });
                    
                    // Pattern 2: User Activity Logs
                    MonitoringService.info('Access token acquired silently', {
                        operation: 'getAccessToken',
                        source: 'silent_acquisition',
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'auth', null, userId);
                    
                    return response.accessToken;
                }
            } catch (error) {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Silent token acquisition failed', {
                        error: error.message,
                        userId: userId.substring(0, 8) + '...',
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }
                throw error;
            }
        }
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('User not authenticated - no token available', {
                error: 'User not authenticated',
                operation: 'getAccessToken',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('User not authenticated - no token available', {
                sessionId,
                error: 'User not authenticated',
                operation: 'getAccessToken',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        throw new Error('User not authenticated');
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `Failed to get access token: ${error.message}`,
            'error',
            {
                stack: error.stack,
                operation: 'getAccessToken',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Access token retrieval failed', {
                error: error.message,
                operation: 'getAccessToken',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Access token retrieval failed', {
                sessionId,
                error: error.message,
                operation: 'getAccessToken',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        throw error;
    }
}

/**
 * Get detailed status information about the authentication service
 * @param {Object} req - Express request object
 * @returns {Promise<Object>} Status details
 */
async function statusDetails(req) {
    const startTime = Date.now();
    const { userId, sessionId } = extractUserContext(req);
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing statusDetails request', {
                sessionId,
                userId,
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        if (await isAuthenticated(req)) {
            // First try to get user info from Express session
            let userInfo = req.session?.msUser;
            
            // Fallback: get from database using session ID
            if (!userInfo && req.session?.id) {
                try {
                    const userKey = `user:${req.session.id}`;
                    const storedUserInfo = await storageService.getSetting(`${userKey}:ms-user-info`, req.session.id);
                    if (storedUserInfo) {
                        userInfo = storedUserInfo;
                    }
                } catch (error) {
                    // Pattern 3: Infrastructure Error Logging
                    const mcpError = ErrorService.createError(
                        'auth',
                        'Failed to get user info from database',
                        'warning',
                        {
                            sessionId: req.session.id,
                            error: error.message,
                            stack: error.stack,
                            operation: 'statusDetails',
                            timestamp: new Date().toISOString()
                        }
                    );
                    MonitoringService.logError(mcpError);
                    
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Failed to get user info from database', {
                            sessionId: req.session.id,
                            error: error.message,
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }
                }
            }
            
            const statusResult = {
                authenticated: true,
                user: userInfo?.username || 'Unknown User',
                name: userInfo?.name,
                sessionId: req.session?.id,
                message: 'Authenticated',
                logoutUrl: '/api/auth/logout'
            };
            
            // Pattern 2: User Activity Logs
            const finalUserId = userInfo?.username ? `ms365:${userInfo.username}` : userId;
            if (finalUserId) {
                MonitoringService.info('Status details retrieved for authenticated user', {
                    username: userInfo?.username,
                    operation: 'statusDetails',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, finalUserId);
            } else if (sessionId) {
                MonitoringService.info('Status details retrieved with session', {
                    sessionId,
                    operation: 'statusDetails',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            return statusResult;
        } else {
            const statusResult = {
                authenticated: false,
                loginUrl: '/api/auth/login',
                message: 'Not authenticated'
            };
            
            // Pattern 2: User Activity Logs
            if (userId) {
                MonitoringService.info('Status details retrieved for unauthenticated user', {
                    operation: 'statusDetails',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Status details retrieved for unauthenticated session', {
                    sessionId,
                    operation: 'statusDetails',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            return statusResult;
        }
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `Error getting status details: ${error.message}`,
            'error',
            {
                stack: error.stack,
                operation: 'statusDetails',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Status details retrieval failed', {
                error: error.message,
                operation: 'statusDetails',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Status details retrieval failed', {
                sessionId,
                error: error.message,
                operation: 'statusDetails',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Return fallback status on error
        return {
            authenticated: false,
            loginUrl: '/api/auth/login',
            message: 'Error determining authentication status',
            error: true
        };
    }
}

/**
 * Handle logout request
 * @param {Object} req - Express request object
 * @param {Object} res - Express response object
 */
async function logout(req, res) {
    const startTime = Date.now();
    const { userId, sessionId } = extractUserContext(req);
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing logout request', {
                sessionId,
                userId,
                userAgent: req.get('User-Agent'),
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Clear session if available
        if (req.session) {
            req.session.destroy(() => {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('User session destroyed', {
                        sessionId,
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }
            });
        }
        
        // Clear user session
        const logoutUserId = req.session?.userId || req.query.userId || userId;
        clearUserSession(logoutUserId);
        
        // Clear SQLite database storage
        try {
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Clearing authentication token from database', {
                    sessionId,
                    userId: logoutUserId,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
            
            // Clear Microsoft 365-based tokens using the ms365:email format
            // Also clear any legacy session-based tokens for backwards compatibility
            if (req.session?.msUser?.username) {
                const userKey = `ms365:${req.session.msUser.username}`;
                await storageService.setSecureSetting(`${userKey}:ms-access-token`, '', userKey);
                await storageService.setSetting(`${userKey}:ms-user-info`, null, userKey);
            }
            // Also clear legacy session-based tokens if session exists
            if (req.session?.id) {
                const legacyKey = `user:${req.session.id}`;
                await storageService.setSecureSetting(`${legacyKey}:ms-access-token`, '', req.session.id);
                await storageService.setSetting(`${legacyKey}:ms-user-info`, null, req.session.id);
            }
            
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Authentication token cleared from database', {
                    sessionId,
                    userId: logoutUserId,
                    timestamp: new Date().toISOString()
                }, 'auth');
            }
        } catch (dbError) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                `Error clearing token from database: ${dbError.message}`,
                'warning',
                {
                    stack: dbError.stack,
                    operation: 'logout',
                    userId: logoutUserId,
                    sessionId,
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
        }
        
        // Pattern 2: User Activity Logs
        if (logoutUserId) {
            MonitoringService.info('User logout completed successfully', {
                operation: 'logout',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, logoutUserId);
        } else if (sessionId) {
            MonitoringService.info('Session logout completed successfully', {
                sessionId,
                operation: 'logout',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Redirect to home page
        res.redirect('/');
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `Logout error: ${error.message}`,
            'error',
            {
                stack: error.stack,
                operation: 'logout',
                userId,
                sessionId,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Logout failed', {
                error: error.message,
                operation: 'logout',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Logout failed', {
                sessionId,
                error: error.message,
                operation: 'logout',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        res.status(500).json({
            error: 'logout_failed',
            error_description: 'Logout process failed'
        });
    }
}

/**
 * Get the most recently used access token for internal MCP adapter calls.
 * This allows the MCP adapter to leverage existing authentication without handling it directly.
 * @param {string} userId - User ID for multi-user token isolation
 * @returns {Promise<string|null>} The most recent access token, or null if none available
 */
async function getMostRecentToken(userId) {
    const startTime = Date.now();
    
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Processing getMostRecentToken request', {
                userId: userId ? userId.substring(0, 8) + '...' : 'N/A',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // FIRST: Check for external enterprise token (injected via UI)
        // External tokens are the actual Microsoft Graph tokens that can be used directly
        if (userId) {
            try {
                const extController = getExternalTokenController();
                const externalToken = await extController.getActiveExternalToken(userId);
                if (externalToken) {
                    if (process.env.NODE_ENV === 'development') {
                        MonitoringService.debug('Found valid external enterprise token', {
                            userId: userId.substring(0, 8) + '...',
                            timestamp: new Date().toISOString()
                        }, 'auth');
                    }

                    MonitoringService.info('Most recent token retrieved from external source', {
                        operation: 'getMostRecentToken',
                        source: 'external_token',
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'auth', null, userId);

                    return externalToken;
                }
            } catch (extError) {
                // External token check failed - continue with normal flow
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('External token check failed, continuing', {
                        error: extError.message,
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }
            }
        }

        // Try to get token from user sessions (in-memory cache)
        if (userId) {
            const userSession = getUserSession(userId);
            if (userSession?.msUser?.accessToken) {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Found valid token in user session', {
                        userId: userId.substring(0, 8) + '...',
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }

                // Pattern 2: User Activity Logs
                MonitoringService.info('Most recent token retrieved from user session', {
                    operation: 'getMostRecentToken',
                    source: 'user_session',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, userId);

                return userSession.msUser.accessToken;
            }
        }

        // If userId not provided or no session found, try any available user session
        for (const [sessionUserId, userSession] of userSessions.entries()) {
            if (userSession.msUser?.accessToken) {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Found valid token in fallback user session', {
                        userId: sessionUserId.substring(0, 8) + '...',
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }
                
                // Pattern 2: User Activity Logs
                MonitoringService.info('Most recent token retrieved from fallback session', {
                    operation: 'getMostRecentToken',
                    source: 'fallback_session',
                    duration: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'auth', null, sessionUserId);
                
                return userSession.msUser.accessToken;
            }
        }
        
        // If not in memory, try to get from SQLite database with user-specific key
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Trying to get token from SQLite database', {
                userId: userId ? userId.substring(0, 8) + '...' : 'N/A',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        try {
            // Use consistent key format: ms365:email:ms-access-token (matches OAuth callback storage)
            const tokenKey = userId ? `${userId}:ms-access-token` : 'ms-access-token';
            const storedToken = await storageService.getSecureSetting(tokenKey, userId);
            if (storedToken) {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Found valid token in SQLite database', {
                        userId: userId ? userId.substring(0, 8) + '...' : 'N/A',
                        tokenKey,
                        timestamp: new Date().toISOString()
                    }, 'auth');
                }

                // Also load it into memory for future use
                const userInfoKey = userId ? `${userId}:ms-user-info` : 'ms-user-info';
                const userInfo = await storageService.getSetting(userInfoKey, userId) || {};
                if (userId) {
                    setUserSession(userId, {
                        msUser: {
                            ...userInfo,
                            accessToken: storedToken
                        }
                    });
                }
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Most recent token retrieved from database', {
                        operation: 'getMostRecentToken',
                        source: 'database',
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'auth', null, userId);
                }
                
                return storedToken;
            }
        } catch (dbError) {
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'auth',
                `Error getting token from database: ${dbError.message}`,
                'warning',
                {
                    userId,
                    stack: dbError.stack,
                    operation: 'getMostRecentToken',
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
        }
        
        // If no token found, we have no authenticated user
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('No authenticated user found for internal MCP call', {
                userId: userId ? userId.substring(0, 8) + '...' : 'N/A',
                timestamp: new Date().toISOString()
            }, 'auth');
        }
        
        // Pattern 2: User Activity Logs (for no token case)
        if (userId) {
            MonitoringService.info('No recent token available', {
                operation: 'getMostRecentToken',
                result: 'no_token',
                duration: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        }
        
        return null;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'auth',
            `Error getting most recent token: ${error.message}`,
            'error',
            {
                userId,
                stack: error.stack,
                operation: 'getMostRecentToken',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Most recent token retrieval failed', {
                error: error.message,
                operation: 'getMostRecentToken',
                timestamp: new Date().toISOString()
            }, 'auth', null, userId);
        }
        
        return null;
    }
}

module.exports = { 
    isAuthenticated, 
    statusDetails, 
    login, 
    handleAuthCallback, 
    logout, 
    getAccessToken,
    getMostRecentToken
};
