/**
 * @fileoverview Graph Token Exchange Controller
 *
 * Exchanges Microsoft Graph access tokens for MCP JWT bearer tokens.
 * This enables Synthetic Employees (and other ROPC clients) to authenticate
 * with the MCP server using their Graph tokens.
 *
 * Flow:
 * 1. Client obtains MS Graph access token (via MSAL ROPC)
 * 2. Client calls POST /api/auth/graph-token-exchange with the Graph token
 * 3. Server validates the Graph token (decode + verify with /me endpoint)
 * 4. Server generates and returns an MCP JWT bearer token (24h validity)
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const StorageService = require('../../core/storage-service.cjs');
const DeviceJwtService = require('../../auth/device-jwt.cjs');
const ExternalTokenValidator = require('../../auth/external-token-validator.cjs');
const crypto = require('crypto');

// Storage keys (matching external-token-controller)
const STORAGE_KEYS = {
    TOKEN: 'external-graph-token',
    METADATA: 'external-token-metadata',
    SOURCE: 'token-source'
};

/**
 * Get storage key with user prefix
 */
function getStorageKey(userId, key) {
    return `${userId}:${key}`;
}

// Validation schemas
const exchangeTokenSchema = Joi.object({
    graph_access_token: Joi.string().required().description('Microsoft Graph access token to exchange')
});

/**
 * Helper function to validate request against schema
 */
function validateRequest(req, schema, endpoint) {
    const { error, value } = schema.validate(req.body, {
        abortEarly: false,
        stripUnknown: true
    });

    if (error) {
        const validationError = ErrorService.createError(
            'validation',
            `Invalid request data for ${endpoint}`,
            'warning',
            {
                endpoint,
                validationErrors: error.details.map(detail => ({
                    field: detail.path.join('.'),
                    message: detail.message
                }))
            }
        );
        MonitoringService.logError(validationError);
    }

    return { error, value };
}

/**
 * Verify Graph token by calling Microsoft Graph /me endpoint
 * @param {string} graphToken - Microsoft Graph access token
 * @returns {Promise<Object>} User profile from Graph API
 */
async function verifyGraphTokenWithApi(graphToken) {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
        headers: {
            'Authorization': `Bearer ${graphToken}`,
            'Content-Type': 'application/json'
        }
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Graph API validation failed: ${response.status} - ${errorText}`);
    }

    return await response.json();
}

/**
 * Exchange a Microsoft Graph access token for an MCP JWT bearer token
 * POST /api/auth/graph-token-exchange
 *
 * This endpoint does NOT require prior authentication - it validates
 * the provided Graph token and issues an MCP JWT if valid.
 */
async function exchange(req, res) {
    const startTime = Date.now();

    try {
        // Validate request body
        const { error, value } = validateRequest(req, exchangeTokenSchema, '/api/auth/graph-token-exchange');
        if (error) {
            return res.status(400).json({
                error: 'INVALID_REQUEST',
                error_description: error.details.map(d => d.message).join(', ')
            });
        }

        const { graph_access_token } = value;

        MonitoringService.info('Graph token exchange attempt', {
            tokenPrefix: graph_access_token.substring(0, 20) + '...',
            ip: req.ip || req.connection?.remoteAddress,
            userAgent: req.get('User-Agent')?.substring(0, 50),
            timestamp: new Date().toISOString()
        }, 'auth');

        // Step 1: Quick validate the token structure and claims
        let tokenValidation;
        try {
            tokenValidation = ExternalTokenValidator.quickValidate(graph_access_token);
            if (!tokenValidation.valid) {
                MonitoringService.warn('Graph token validation failed', {
                    error: tokenValidation.error,
                    timestamp: new Date().toISOString()
                }, 'auth');

                return res.status(401).json({
                    error: 'INVALID_TOKEN',
                    error_description: tokenValidation.error || 'Token validation failed'
                });
            }
        } catch (validationError) {
            MonitoringService.warn('Graph token decode failed', {
                error: validationError.message,
                timestamp: new Date().toISOString()
            }, 'auth');

            return res.status(401).json({
                error: 'INVALID_TOKEN',
                error_description: 'Failed to decode token'
            });
        }

        // Step 2: Verify token with Microsoft Graph API /me endpoint
        let graphUser;
        try {
            graphUser = await verifyGraphTokenWithApi(graph_access_token);
        } catch (graphError) {
            MonitoringService.warn('Graph API verification failed', {
                error: graphError.message,
                timestamp: new Date().toISOString()
            }, 'auth');

            return res.status(401).json({
                error: 'TOKEN_VERIFICATION_FAILED',
                error_description: 'Could not verify token with Microsoft Graph API'
            });
        }

        // Extract user information
        const userEmail = graphUser.mail || graphUser.userPrincipalName;
        const userName = graphUser.displayName;
        const userId = graphUser.id;

        if (!userEmail || !userId) {
            MonitoringService.warn('Graph user info incomplete', {
                hasEmail: !!userEmail,
                hasUserId: !!userId,
                timestamp: new Date().toISOString()
            }, 'auth');

            return res.status(401).json({
                error: 'INVALID_USER_INFO',
                error_description: 'Could not extract user information from token'
            });
        }

        // Step 3: Generate MCP JWT using DeviceJwtService
        // Use a consistent device ID based on user email for token management
        const deviceId = `synthetic-employee-${crypto.createHash('sha256').update(userEmail).digest('hex').substring(0, 16)}`;

        // CRITICAL: userId must be ms365:email format to match token storage key format
        // The getAccessToken function looks up tokens using ${userId}:ms-access-token
        const mcpUserId = `ms365:${userEmail}`;

        const mcpToken = DeviceJwtService.generateLongLivedAccessToken(
            deviceId,
            mcpUserId,  // Use ms365:email format, not Graph user GUID
            {
                email: userEmail,
                name: userName,
                graphUserId: userId,  // Store the original Graph user ID in metadata
                source: 'graph-token-exchange',
                exchangedAt: new Date().toISOString()
            }
        );

        // Calculate expiration (24 hours from now)
        const expiresAt = new Date(Date.now() + 24 * 60 * 60 * 1000);
        const expiresIn = 24 * 60 * 60; // 24 hours in seconds

        // Step 4: Store the Graph token for MCP server to use when calling Graph API
        // This is critical - without this, the MCP server can't make Graph API calls
        const storageUserId = `ms365:${userEmail}`;

        // Store the Graph access token
        const tokenKey = getStorageKey(storageUserId, STORAGE_KEYS.TOKEN);
        await StorageService.setSecureSetting(tokenKey, graph_access_token, storageUserId);

        // Also store under ms-access-token for compatibility with isAuthenticated middleware
        const msAccessTokenKey = `${storageUserId}:ms-access-token`;
        await StorageService.setSecureSetting(msAccessTokenKey, graph_access_token, storageUserId);

        // Store metadata
        const metadata = {
            user: { id: userId, email: userEmail, name: userName },
            expires_at: tokenValidation.metadata?.expires_at || expiresAt.toISOString(),
            scopes: tokenValidation.metadata?.scopes || [],
            source: 'graph-token-exchange'
        };
        const metadataKey = getStorageKey(storageUserId, STORAGE_KEYS.METADATA);
        await StorageService.setSecureSetting(metadataKey, JSON.stringify(metadata), storageUserId);

        // Set token source to external
        const sourceKey = getStorageKey(storageUserId, STORAGE_KEYS.SOURCE);
        await StorageService.setSecureSetting(sourceKey, 'external', storageUserId);

        MonitoringService.info('Graph token exchange successful', {
            userEmail,
            mcpUserId,
            graphUserId: userId.substring(0, 8) + '...',
            deviceId,
            graphTokenStored: true,
            duration: Date.now() - startTime,
            timestamp: new Date().toISOString()
        }, 'auth', null, mcpUserId);

        // Return the MCP JWT
        res.json({
            access_token: mcpToken,
            token_type: 'Bearer',
            expires_in: expiresIn,
            expires_at: expiresAt.toISOString(),
            user: {
                id: userId,
                email: userEmail,
                name: userName
            }
        });

    } catch (error) {
        MonitoringService.error('Graph token exchange error', {
            error: error.message,
            stack: error.stack,
            duration: Date.now() - startTime,
            timestamp: new Date().toISOString()
        }, 'auth');

        res.status(500).json({
            error: 'EXCHANGE_FAILED',
            error_description: 'Failed to exchange token'
        });
    }
}

module.exports = {
    exchange
};
