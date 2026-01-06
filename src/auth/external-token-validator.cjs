/**
 * @fileoverview External Token Validator Service
 *
 * Validates externally-obtained Microsoft Graph access tokens.
 * These tokens come from enterprise tools (e.g., Windows auth tools)
 * and cannot be refreshed - they must be re-injected when expired.
 */

const https = require('https');

/**
 * Error codes for token validation failures
 */
const ERROR_CODES = {
  INVALID_FORMAT: 'INVALID_FORMAT',
  TOKEN_EXPIRED: 'TOKEN_EXPIRED',
  INVALID_AUDIENCE: 'INVALID_AUDIENCE',
  VALIDATION_FAILED: 'VALIDATION_FAILED',
  MISSING_CLAIMS: 'MISSING_CLAIMS'
};

/**
 * Expected audience for Microsoft Graph tokens
 */
const GRAPH_AUDIENCE = 'https://graph.microsoft.com';

/**
 * Validation error class
 */
class TokenValidationError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'TokenValidationError';
    this.code = code;
  }
}

/**
 * Decode a JWT token without signature verification
 * Microsoft signs their own tokens - we just need to read the claims
 *
 * @param {string} token - The JWT token string
 * @returns {object} Decoded token with header and payload
 * @throws {TokenValidationError} If token format is invalid
 */
function decodeToken(token) {
  if (!token || typeof token !== 'string') {
    throw new TokenValidationError(
      ERROR_CODES.INVALID_FORMAT,
      'Token must be a non-empty string'
    );
  }

  // Clean the token (remove Bearer prefix if present)
  const cleanToken = token.replace(/^Bearer\s+/i, '').trim();

  // JWT must have exactly 3 parts separated by dots
  const parts = cleanToken.split('.');
  if (parts.length !== 3) {
    throw new TokenValidationError(
      ERROR_CODES.INVALID_FORMAT,
      'Invalid JWT format: expected 3 parts separated by dots'
    );
  }

  try {
    // Decode header and payload (base64url)
    const header = JSON.parse(Buffer.from(parts[0], 'base64url').toString('utf8'));
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));

    return {
      header,
      payload,
      raw: cleanToken
    };
  } catch (error) {
    throw new TokenValidationError(
      ERROR_CODES.INVALID_FORMAT,
      `Failed to decode JWT: ${error.message}`
    );
  }
}

/**
 * Validate the token has required structure and claims
 *
 * @param {object} decoded - Decoded token from decodeToken()
 * @throws {TokenValidationError} If required claims are missing
 */
function validateTokenStructure(decoded) {
  const { payload } = decoded;

  const requiredClaims = ['aud', 'exp', 'iat'];
  const missingClaims = requiredClaims.filter(claim => !(claim in payload));

  if (missingClaims.length > 0) {
    throw new TokenValidationError(
      ERROR_CODES.MISSING_CLAIMS,
      `Missing required claims: ${missingClaims.join(', ')}`
    );
  }
}

/**
 * Validate the token audience is Microsoft Graph
 *
 * @param {object} decoded - Decoded token from decodeToken()
 * @throws {TokenValidationError} If audience is not Microsoft Graph
 */
function validateAudience(decoded) {
  const { payload } = decoded;

  // Audience can be a string or array
  const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];

  if (!audiences.includes(GRAPH_AUDIENCE)) {
    throw new TokenValidationError(
      ERROR_CODES.INVALID_AUDIENCE,
      `Token audience must be ${GRAPH_AUDIENCE}, got: ${payload.aud}`
    );
  }
}

/**
 * Validate the token has not expired
 *
 * @param {object} decoded - Decoded token from decodeToken()
 * @throws {TokenValidationError} If token has expired
 */
function validateExpiration(decoded) {
  const { payload } = decoded;
  const now = Math.floor(Date.now() / 1000);

  if (payload.exp <= now) {
    const expiredAt = new Date(payload.exp * 1000).toISOString();
    throw new TokenValidationError(
      ERROR_CODES.TOKEN_EXPIRED,
      `Token expired at ${expiredAt}`
    );
  }
}

/**
 * Test the token by making a call to Microsoft Graph /me endpoint
 *
 * @param {string} token - The raw JWT token
 * @returns {Promise<object>} User profile data from Graph API
 * @throws {TokenValidationError} If the API call fails
 */
async function testWithGraphAPI(token) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com',
      path: '/v1.0/me',
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          try {
            resolve(JSON.parse(data));
          } catch {
            resolve({ raw: data });
          }
        } else {
          let errorMessage = `Graph API returned ${res.statusCode}`;
          try {
            const errorData = JSON.parse(data);
            errorMessage = errorData.error?.message || errorMessage;
          } catch {
            // Use default message
          }
          reject(new TokenValidationError(
            ERROR_CODES.VALIDATION_FAILED,
            `Token validation failed: ${errorMessage}`
          ));
        }
      });
    });

    req.on('error', (error) => {
      reject(new TokenValidationError(
        ERROR_CODES.VALIDATION_FAILED,
        `Network error during validation: ${error.message}`
      ));
    });

    req.setTimeout(10000, () => {
      req.destroy();
      reject(new TokenValidationError(
        ERROR_CODES.VALIDATION_FAILED,
        'Token validation request timed out'
      ));
    });

    req.end();
  });
}

/**
 * Extract metadata from a decoded token
 *
 * @param {object} decoded - Decoded token from decodeToken()
 * @param {object} graphProfile - Optional profile from Graph API test
 * @returns {object} Token metadata
 */
function extractMetadata(decoded, graphProfile = null) {
  const { payload } = decoded;
  const now = Math.floor(Date.now() / 1000);

  // Parse scopes from 'scp' claim (space-separated string)
  const scopes = payload.scp ? payload.scp.split(' ').sort() : [];

  // Calculate time remaining
  const expiresInSeconds = Math.max(0, payload.exp - now);
  const expiresAt = new Date(payload.exp * 1000).toISOString();

  // User info from token or Graph API response
  const user = {
    id: payload.oid || payload.sub || null,
    name: graphProfile?.displayName || payload.name || null,
    email: graphProfile?.mail || graphProfile?.userPrincipalName ||
           payload.upn || payload.unique_name || payload.preferred_username || null,
    tenant: payload.tid || null
  };

  // App info
  const app = {
    id: payload.appid || payload.azp || null,
    name: payload.app_displayname || null
  };

  return {
    user,
    app,
    scopes,
    expires_at: expiresAt,
    expires_in_seconds: expiresInSeconds,
    is_expiring_soon: expiresInSeconds < 600, // < 10 minutes
    issued_at: new Date(payload.iat * 1000).toISOString()
  };
}

/**
 * Fully validate an external token and return metadata
 *
 * @param {string} token - The JWT token string
 * @param {object} options - Validation options
 * @param {boolean} options.skipGraphTest - Skip Graph API test (for faster validation)
 * @returns {Promise<object>} Validation result with metadata
 */
async function validateExternalToken(token, options = {}) {
  const { skipGraphTest = false } = options;

  // Step 1: Decode the token
  const decoded = decodeToken(token);

  // Step 2: Validate structure
  validateTokenStructure(decoded);

  // Step 3: Validate audience
  validateAudience(decoded);

  // Step 4: Validate expiration
  validateExpiration(decoded);

  // Step 5: Test with Graph API (unless skipped)
  let graphProfile = null;
  if (!skipGraphTest) {
    graphProfile = await testWithGraphAPI(decoded.raw);
  }

  // Step 6: Extract and return metadata
  const metadata = extractMetadata(decoded, graphProfile);

  return {
    valid: true,
    token: decoded.raw,
    metadata
  };
}

/**
 * Quick validation without Graph API test
 * Useful for checking if a stored token is still valid
 *
 * @param {string} token - The JWT token string
 * @returns {object} Validation result
 */
function quickValidate(token) {
  try {
    const decoded = decodeToken(token);
    validateTokenStructure(decoded);
    validateAudience(decoded);
    validateExpiration(decoded);
    const metadata = extractMetadata(decoded);
    return { valid: true, metadata };
  } catch (error) {
    return {
      valid: false,
      error: error.code || 'UNKNOWN_ERROR',
      message: error.message
    };
  }
}

/**
 * Redact token for logging (show only first 8 characters)
 *
 * @param {string} token - The token to redact
 * @returns {string} Redacted token
 */
function redactToken(token) {
  if (!token || token.length < 8) return '[REDACTED]';
  return token.substring(0, 8) + '...[REDACTED]';
}

module.exports = {
  ERROR_CODES,
  TokenValidationError,
  decodeToken,
  validateTokenStructure,
  validateAudience,
  validateExpiration,
  testWithGraphAPI,
  extractMetadata,
  validateExternalToken,
  quickValidate,
  redactToken,
  GRAPH_AUDIENCE
};
