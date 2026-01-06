/**
 * @fileoverview External Token Controller
 *
 * Handles API endpoints for injecting and managing externally-obtained
 * Microsoft Graph access tokens (e.g., from enterprise Windows tools).
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const StorageService = require('../../core/storage-service.cjs');
const ExternalTokenValidator = require('../../auth/external-token-validator.cjs');

// Storage keys
const STORAGE_KEYS = {
  TOKEN: 'external-graph-token',
  METADATA: 'external-token-metadata',
  SOURCE: 'token-source'
};

// Validation schemas
const injectTokenSchema = Joi.object({
  access_token: Joi.string().required().description('Microsoft Graph access token')
});

const switchSourceSchema = Joi.object({
  source: Joi.string().valid('oauth', 'external').required().description('Token source to use')
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
 * Get user ID from request
 */
function getUserId(req) {
  return req.user?.userId ||
         (req.session?.msUser?.username ? `ms365:${req.session.msUser.username}` : null);
}

/**
 * Get storage key with user prefix
 */
function getStorageKey(userId, key) {
  return `${userId}:${key}`;
}

/**
 * Inject an external Microsoft Graph token
 * POST /api/auth/external-token
 */
async function inject(req, res) {
  try {
    const userId = getUserId(req);
    if (!userId) {
      return res.status(401).json({
        error: 'UNAUTHORIZED',
        message: 'Authentication required'
      });
    }

    // Validate request
    const { error, value } = validateRequest(req, injectTokenSchema, '/api/auth/external-token');
    if (error) {
      return res.status(400).json({
        error: 'INVALID_REQUEST',
        message: error.details.map(d => d.message).join(', ')
      });
    }

    const { access_token } = value;

    MonitoringService.info('External token injection attempt', {
      userId,
      tokenPrefix: ExternalTokenValidator.redactToken(access_token),
      timestamp: new Date().toISOString()
    }, 'auth');

    // Validate the token
    let validationResult;
    try {
      validationResult = await ExternalTokenValidator.validateExternalToken(access_token);
    } catch (validationError) {
      MonitoringService.warn('External token validation failed', {
        userId,
        errorCode: validationError.code,
        errorMessage: validationError.message,
        timestamp: new Date().toISOString()
      }, 'auth');

      return res.status(400).json({
        error: validationError.code || 'VALIDATION_FAILED',
        message: validationError.message
      });
    }

    // Store the token (encrypted)
    const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
    await StorageService.setSecureSetting(tokenKey, validationResult.token, userId);

    // Store metadata (JSON)
    const metadataKey = getStorageKey(userId, STORAGE_KEYS.METADATA);
    await StorageService.setSecureSetting(
      metadataKey,
      JSON.stringify(validationResult.metadata),
      userId
    );

    // Set token source to external
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    await StorageService.setSecureSetting(sourceKey, 'external', userId);

    MonitoringService.info('External token injected successfully', {
      userId,
      userEmail: validationResult.metadata.user.email,
      scopeCount: validationResult.metadata.scopes.length,
      expiresAt: validationResult.metadata.expires_at,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.json({
      success: true,
      metadata: validationResult.metadata
    });

  } catch (error) {
    MonitoringService.error('External token injection error', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.status(500).json({
      error: 'INTERNAL_ERROR',
      message: 'Failed to inject external token'
    });
  }
}

/**
 * Get external token status
 * GET /api/auth/external-token/status
 */
async function status(req, res) {
  try {
    const userId = getUserId(req);
    if (!userId) {
      return res.status(401).json({
        error: 'UNAUTHORIZED',
        message: 'Authentication required'
      });
    }

    // Get token source
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    const tokenSource = await StorageService.getSecureSetting(sourceKey, userId);

    // Get stored token
    const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
    const storedToken = await StorageService.getSecureSetting(tokenKey, userId);

    if (!storedToken) {
      return res.json({
        has_external_token: false,
        is_active: false,
        token_source: tokenSource || 'oauth'
      });
    }

    // Quick validate the stored token
    const validation = ExternalTokenValidator.quickValidate(storedToken);

    if (!validation.valid) {
      // Token is invalid/expired - clear it
      await StorageService.deleteSecureSetting(tokenKey, userId);
      await StorageService.deleteSecureSetting(
        getStorageKey(userId, STORAGE_KEYS.METADATA),
        userId
      );

      return res.json({
        has_external_token: false,
        is_active: false,
        token_source: tokenSource || 'oauth',
        expired_reason: validation.error
      });
    }

    res.json({
      has_external_token: true,
      is_active: tokenSource === 'external',
      token_source: tokenSource || 'oauth',
      metadata: validation.metadata
    });

  } catch (error) {
    MonitoringService.error('External token status error', {
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.status(500).json({
      error: 'INTERNAL_ERROR',
      message: 'Failed to get token status'
    });
  }
}

/**
 * Clear external token
 * DELETE /api/auth/external-token
 */
async function clear(req, res) {
  try {
    const userId = getUserId(req);
    if (!userId) {
      return res.status(401).json({
        error: 'UNAUTHORIZED',
        message: 'Authentication required'
      });
    }

    // Delete token and metadata
    const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
    const metadataKey = getStorageKey(userId, STORAGE_KEYS.METADATA);

    await StorageService.deleteSecureSetting(tokenKey, userId);
    await StorageService.deleteSecureSetting(metadataKey, userId);

    // Reset token source to oauth
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    await StorageService.setSecureSetting(sourceKey, 'oauth', userId);

    MonitoringService.info('External token cleared', {
      userId,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.json({
      success: true,
      message: 'External token cleared successfully'
    });

  } catch (error) {
    MonitoringService.error('External token clear error', {
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.status(500).json({
      error: 'INTERNAL_ERROR',
      message: 'Failed to clear external token'
    });
  }
}

/**
 * Switch token source between oauth and external
 * POST /api/auth/external-token/switch
 */
async function switchSource(req, res) {
  try {
    const userId = getUserId(req);
    if (!userId) {
      return res.status(401).json({
        error: 'UNAUTHORIZED',
        message: 'Authentication required'
      });
    }

    // Validate request
    const { error, value } = validateRequest(req, switchSourceSchema, '/api/auth/external-token/switch');
    if (error) {
      return res.status(400).json({
        error: 'INVALID_REQUEST',
        message: error.details.map(d => d.message).join(', ')
      });
    }

    const { source } = value;

    // If switching to external, verify we have a valid external token
    if (source === 'external') {
      const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
      const storedToken = await StorageService.getSecureSetting(tokenKey, userId);

      if (!storedToken) {
        return res.status(400).json({
          error: 'NO_EXTERNAL_TOKEN',
          message: 'No external token available. Please inject a token first.'
        });
      }

      const validation = ExternalTokenValidator.quickValidate(storedToken);
      if (!validation.valid) {
        return res.status(400).json({
          error: 'EXTERNAL_TOKEN_INVALID',
          message: `External token is invalid: ${validation.message}`
        });
      }
    }

    // Update token source
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    await StorageService.setSecureSetting(sourceKey, source, userId);

    MonitoringService.info('Token source switched', {
      userId,
      newSource: source,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.json({
      success: true,
      active_source: source
    });

  } catch (error) {
    MonitoringService.error('Token source switch error', {
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.status(500).json({
      error: 'INTERNAL_ERROR',
      message: 'Failed to switch token source'
    });
  }
}

/**
 * Login with external token (no auth required)
 * POST /api/auth/external-token/login
 *
 * This endpoint allows users to authenticate using an external enterprise token
 * without requiring prior OAuth authentication.
 */
async function loginWithToken(req, res) {
  try {
    // Validate request
    const { error, value } = validateRequest(req, injectTokenSchema, '/api/auth/external-token/login');
    if (error) {
      return res.status(400).json({
        error: 'INVALID_REQUEST',
        message: error.details.map(d => d.message).join(', ')
      });
    }

    const { access_token } = value;

    MonitoringService.info('External token login attempt', {
      tokenPrefix: ExternalTokenValidator.redactToken(access_token),
      timestamp: new Date().toISOString()
    }, 'auth');

    // Validate the token
    let validationResult;
    try {
      validationResult = await ExternalTokenValidator.validateExternalToken(access_token);
    } catch (validationError) {
      MonitoringService.warn('External token login validation failed', {
        errorCode: validationError.code,
        errorMessage: validationError.message,
        timestamp: new Date().toISOString()
      }, 'auth');

      return res.status(400).json({
        error: validationError.code || 'VALIDATION_FAILED',
        message: validationError.message
      });
    }

    // Create user ID from token metadata
    const userEmail = validationResult.metadata.user.email;
    const userId = userEmail ? `ms365:${userEmail}` : `ms365:${validationResult.metadata.user.id}`;

    // Store user info in session
    if (req.session) {
      req.session.msUser = {
        username: userEmail || validationResult.metadata.user.id,
        name: validationResult.metadata.user.name || 'External User',
        homeAccountId: validationResult.metadata.user.id,
        accessToken: validationResult.token,
        expiresOn: new Date(validationResult.metadata.expires_at),
        authMethod: 'external_token'
      };

      // Force session save
      await new Promise((resolve, reject) => {
        req.session.save((err) => {
          if (err) reject(err);
          else resolve();
        });
      });
    }

    // Store the token (encrypted)
    const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
    await StorageService.setSecureSetting(tokenKey, validationResult.token, userId);

    // Store metadata (JSON)
    const metadataKey = getStorageKey(userId, STORAGE_KEYS.METADATA);
    await StorageService.setSecureSetting(
      metadataKey,
      JSON.stringify(validationResult.metadata),
      userId
    );

    // Set token source to external
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    await StorageService.setSecureSetting(sourceKey, 'external', userId);

    MonitoringService.info('External token login successful', {
      userId,
      userEmail,
      scopeCount: validationResult.metadata.scopes.length,
      expiresAt: validationResult.metadata.expires_at,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.json({
      success: true,
      authenticated: true,
      user: {
        name: validationResult.metadata.user.name,
        email: userEmail,
        id: validationResult.metadata.user.id
      },
      metadata: validationResult.metadata
    });

  } catch (error) {
    MonitoringService.error('External token login error', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    }, 'auth');

    res.status(500).json({
      error: 'INTERNAL_ERROR',
      message: 'Failed to login with external token'
    });
  }
}

/**
 * Get the current active external token (for internal use by MSAL service)
 * Returns null if no valid external token or source is not external
 */
async function getActiveExternalToken(userId) {
  try {
    // Check token source
    const sourceKey = getStorageKey(userId, STORAGE_KEYS.SOURCE);
    const tokenSource = await StorageService.getSecureSetting(sourceKey, userId);

    if (tokenSource !== 'external') {
      return null;
    }

    // Get token
    const tokenKey = getStorageKey(userId, STORAGE_KEYS.TOKEN);
    const storedToken = await StorageService.getSecureSetting(tokenKey, userId);

    if (!storedToken) {
      return null;
    }

    // Validate token is still valid
    const validation = ExternalTokenValidator.quickValidate(storedToken);
    if (!validation.valid) {
      return null;
    }

    return storedToken;

  } catch (error) {
    MonitoringService.error('Get active external token error', {
      userId,
      error: error.message,
      timestamp: new Date().toISOString()
    }, 'auth');
    return null;
  }
}

module.exports = {
  inject,
  status,
  clear,
  switchSource,
  loginWithToken,
  getActiveExternalToken,
  STORAGE_KEYS
};
