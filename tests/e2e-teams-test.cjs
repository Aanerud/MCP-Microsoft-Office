#!/usr/bin/env node
/**
 * @fileoverview E2E test for Teams module functionality
 *
 * Tests the Teams API endpoints against a running server.
 * Requires a valid Microsoft Graph token with Teams permissions.
 *
 * Usage:
 *   node tests/e2e-teams-test.cjs
 */

const http = require('http');
const fs = require('fs');
const path = require('path');

// ANSI colors
const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  cyan: '\x1b[36m',
  dim: '\x1b[2m',
  bold: '\x1b[1m'
};

const API_BASE = process.env.API_BASE || 'http://localhost:3000';
const TOKEN_FILE = path.join(__dirname, '..', 'token.txt');

// Test state
let sessionCookies = [];
let testResults = { passed: 0, failed: 0, skipped: 0, tests: [] };

/**
 * Load token from file
 */
function loadToken() {
  if (!fs.existsSync(TOKEN_FILE)) {
    throw new Error('token.txt not found. Please create it with a valid Microsoft Graph token.');
  }
  return fs.readFileSync(TOKEN_FILE, 'utf8').trim();
}

/**
 * Make HTTP request with cookie support
 */
function makeRequest(method, endpoint, body = null, headers = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(endpoint.startsWith('/') ? `${API_BASE}${endpoint}` : endpoint);

    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || 80,
      path: parsedUrl.pathname + parsedUrl.search,
      method,
      headers: {
        'Content-Type': 'application/json',
        ...headers
      }
    };

    // Add session cookies
    if (sessionCookies.length > 0) {
      options.headers['Cookie'] = sessionCookies.join('; ');
    }

    if (body) {
      const bodyStr = JSON.stringify(body);
      options.headers['Content-Length'] = Buffer.byteLength(bodyStr);
    }

    const req = http.request(options, (res) => {
      // Store set-cookie headers
      const setCookies = res.headers['set-cookie'];
      if (setCookies) {
        setCookies.forEach(cookie => {
          const cookieName = cookie.split('=')[0];
          sessionCookies = sessionCookies.filter(c => !c.startsWith(cookieName + '='));
          sessionCookies.push(cookie.split(';')[0]);
        });
      }

      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve({
            status: res.statusCode,
            data: data ? JSON.parse(data) : {},
            headers: res.headers,
            success: res.statusCode >= 200 && res.statusCode < 300
          });
        } catch {
          resolve({
            status: res.statusCode,
            data: data,
            headers: res.headers,
            success: res.statusCode >= 200 && res.statusCode < 300
          });
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(15000, () => {
      req.destroy();
      reject(new Error('Request timeout'));
    });

    if (body) {
      req.write(JSON.stringify(body));
    }
    req.end();
  });
}

/**
 * Log test result
 */
function logTest(name, passed, details = '', skipped = false) {
  let status;
  if (skipped) {
    status = `${colors.yellow}⊘ SKIP${colors.reset}`;
    testResults.skipped++;
  } else if (passed) {
    status = `${colors.green}✓ PASS${colors.reset}`;
    testResults.passed++;
  } else {
    status = `${colors.red}✗ FAIL${colors.reset}`;
    testResults.failed++;
  }

  console.log(`  ${name.padEnd(50)} ${status}`);
  if (details && !skipped) {
    console.log(`    ${colors.dim}${details}${colors.reset}`);
  }
  testResults.tests.push({ name, passed, details, skipped });
}

/**
 * Main test suite
 */
async function main() {
  console.log(`\n${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                 TEAMS MODULE E2E TESTS${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}\n`);

  // Load token
  let graphToken;
  try {
    graphToken = loadToken();
    console.log(`${colors.green}✓${colors.reset} Token loaded from token.txt\n`);
  } catch (error) {
    console.log(`${colors.red}✗${colors.reset} ${error.message}\n`);
    process.exit(1);
  }

  // ═══════════════════════════════════════════════════════════════
  // SETUP: Login with external token
  // ═══════════════════════════════════════════════════════════════
  console.log(`${colors.bold}Setup: Authentication${colors.reset}`);
  console.log(`${colors.dim}─────────────────────────────────────────────────────────────────${colors.reset}`);

  try {
    const loginRes = await makeRequest('POST', '/api/auth/external-token/login', {
      access_token: graphToken
    });

    if (loginRes.success && loginRes.data.success) {
      logTest('Login with Graph token', true, `Authenticated as ${loginRes.data.user?.email || 'user'}`);
    } else {
      logTest('Login with Graph token', false, loginRes.data?.error || 'Unknown error');
      console.log(`\n${colors.red}Cannot continue without authentication${colors.reset}\n`);
      process.exit(1);
    }
  } catch (error) {
    logTest('Login with Graph token', false, error.message);
    console.log(`\n${colors.red}Cannot continue without authentication${colors.reset}\n`);
    process.exit(1);
  }

  // Store IDs for chained tests
  let teamId = null;
  let channelId = null;
  let chatId = null;

  // ═══════════════════════════════════════════════════════════════
  // TEAMS OPERATIONS
  // ═══════════════════════════════════════════════════════════════
  console.log(`\n${colors.bold}Teams Operations${colors.reset}`);
  console.log(`${colors.dim}─────────────────────────────────────────────────────────────────${colors.reset}`);

  // List joined teams
  try {
    const result = await makeRequest('GET', '/api/v1/teams');
    const teams = result.data?.data || [];

    if (result.success && Array.isArray(teams)) {
      if (teams.length > 0) {
        teamId = teams[0].id;
        logTest('List joined teams', true, `Found ${teams.length} teams`);
      } else {
        logTest('List joined teams', true, 'No teams found (0 teams)');
      }
    } else {
      const isPermError = result.status === 403 || result.data?.error?.code === 'Forbidden';
      if (isPermError) {
        logTest('List joined teams', false, 'Missing Team.ReadBasic.All permission', true);
      } else {
        logTest('List joined teams', false, result.data?.error?.message || `Status ${result.status}`);
      }
    }
  } catch (error) {
    logTest('List joined teams', false, error.message);
  }

  // List team channels
  try {
    if (!teamId) {
      logTest('List team channels', false, 'Skipped - no team ID available', true);
    } else {
      const result = await makeRequest('GET', `/api/v1/teams/${teamId}/channels`);
      const channels = result.data?.data || [];

      if (result.success && Array.isArray(channels)) {
        if (channels.length > 0) {
          channelId = channels[0].id;
        }
        logTest('List team channels', true, `Found ${channels.length} channels`);
      } else {
        const isPermError = result.status === 403;
        if (isPermError) {
          logTest('List team channels', false, 'Missing permission', true);
        } else {
          logTest('List team channels', false, result.data?.error?.message || `Status ${result.status}`);
        }
      }
    }
  } catch (error) {
    logTest('List team channels', false, error.message);
  }

  // Get channel messages
  try {
    if (!teamId || !channelId) {
      logTest('Get channel messages', false, 'Skipped - no team/channel ID available', true);
    } else {
      const result = await makeRequest('GET', `/api/v1/teams/${teamId}/channels/${channelId}/messages?limit=5`);

      if (result.success) {
        const messages = result.data?.data || [];
        logTest('Get channel messages', true, `Found ${messages.length} messages`);
      } else {
        const isPermError = result.status === 403;
        if (isPermError) {
          logTest('Get channel messages', false, 'Missing ChannelMessage.Read.All permission', true);
        } else {
          logTest('Get channel messages', false, result.data?.error?.message || `Status ${result.status}`);
        }
      }
    }
  } catch (error) {
    logTest('Get channel messages', false, error.message);
  }

  // ═══════════════════════════════════════════════════════════════
  // CHAT OPERATIONS
  // ═══════════════════════════════════════════════════════════════
  console.log(`\n${colors.bold}Chat Operations${colors.reset}`);
  console.log(`${colors.dim}─────────────────────────────────────────────────────────────────${colors.reset}`);

  // List chats
  try {
    const result = await makeRequest('GET', '/api/v1/teams/chats');

    if (result.success) {
      const chats = result.data?.data || [];
      if (chats.length > 0) {
        chatId = chats[0].id;
      }
      logTest('List chats', true, `Found ${chats.length} chats`);
    } else {
      // Check for permission errors (403 or TEAMS_OPERATION_FAILED which usually indicates permission issues)
      const errorStr = JSON.stringify(result.data);
      const isPermError = result.status === 403 ||
        result.data?.error?.code === 'Forbidden' ||
        errorStr.includes('403') ||
        errorStr.includes('Forbidden') ||
        result.data?.error === 'TEAMS_OPERATION_FAILED'; // Chat.Read permission required
      if (isPermError) {
        logTest('List chats', false, 'Missing Chat.Read permission', true);
      } else {
        logTest('List chats', false, result.data?.error?.message || result.data?.error || `Status ${result.status}`);
      }
    }
  } catch (error) {
    logTest('List chats', false, error.message);
  }

  // Get chat messages
  try {
    if (!chatId) {
      logTest('Get chat messages', false, 'Skipped - no chat ID available', true);
    } else {
      const result = await makeRequest('GET', `/api/v1/teams/chats/${chatId}/messages?limit=5`);

      if (result.success) {
        const messages = result.data?.data || [];
        logTest('Get chat messages', true, `Found ${messages.length} messages`);
      } else {
        const isPermError = result.status === 403;
        if (isPermError) {
          logTest('Get chat messages', false, 'Missing Chat.Read permission', true);
        } else {
          logTest('Get chat messages', false, result.data?.error?.message || `Status ${result.status}`);
        }
      }
    }
  } catch (error) {
    logTest('Get chat messages', false, error.message);
  }

  // ═══════════════════════════════════════════════════════════════
  // MEETING OPERATIONS
  // ═══════════════════════════════════════════════════════════════
  console.log(`\n${colors.bold}Meeting Operations${colors.reset}`);
  console.log(`${colors.dim}─────────────────────────────────────────────────────────────────${colors.reset}`);

  // List online meetings
  try {
    const result = await makeRequest('GET', '/api/v1/teams/meetings');

    if (result.success) {
      const meetings = result.data?.data || [];
      logTest('List online meetings', true, `Found ${meetings.length} meetings`);
    } else {
      // Check for permission errors (403 or TEAMS_OPERATION_FAILED which usually indicates permission issues)
      const errorStr = JSON.stringify(result.data);
      const isPermError = result.status === 403 ||
        result.data?.error?.code === 'Forbidden' ||
        errorStr.includes('403') ||
        errorStr.includes('Forbidden') ||
        result.data?.error === 'TEAMS_OPERATION_FAILED'; // OnlineMeetings.Read permission required
      if (isPermError) {
        logTest('List online meetings', false, 'Missing OnlineMeetings.Read permission', true);
      } else {
        logTest('List online meetings', false, result.data?.error?.message || result.data?.error || `Status ${result.status}`);
      }
    }
  } catch (error) {
    logTest('List online meetings', false, error.message);
  }

  // ═══════════════════════════════════════════════════════════════
  // SUMMARY
  // ═══════════════════════════════════════════════════════════════
  console.log(`\n${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                         SUMMARY${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}\n`);

  console.log(`${colors.green}Passed:${colors.reset}  ${testResults.passed}`);
  console.log(`${colors.red}Failed:${colors.reset}  ${testResults.failed}`);
  console.log(`${colors.yellow}Skipped:${colors.reset} ${testResults.skipped}`);

  if (testResults.failed > 0) {
    console.log(`\n${colors.bold}Failed tests:${colors.reset}`);
    testResults.tests
      .filter(t => !t.passed && !t.skipped)
      .forEach(t => console.log(`  ${colors.red}✗${colors.reset} ${t.name}: ${t.details}`));
  }

  if (testResults.skipped > 0) {
    console.log(`\n${colors.bold}Skipped tests (missing permissions or dependencies):${colors.reset}`);
    testResults.tests
      .filter(t => t.skipped)
      .forEach(t => console.log(`  ${colors.yellow}⊘${colors.reset} ${t.name}`));
    console.log(`\n${colors.dim}To enable skipped tests, add these permissions to your Azure AD app:${colors.reset}`);
    console.log(`${colors.dim}  - Team.ReadBasic.All (list teams)${colors.reset}`);
    console.log(`${colors.dim}  - Chat.Read / Chat.ReadWrite (chat operations)${colors.reset}`);
    console.log(`${colors.dim}  - ChannelMessage.Read.All / ChannelMessage.Send (channel messages)${colors.reset}`);
    console.log(`${colors.dim}  - OnlineMeetings.Read / OnlineMeetings.ReadWrite (meetings)${colors.reset}`);
  }

  console.log();

  // Exit with error if any tests failed (not skipped)
  process.exit(testResults.failed > 0 ? 1 : 0);
}

// Run tests
main().catch(error => {
  console.error(`${colors.red}Fatal error: ${error.message}${colors.reset}`);
  process.exit(1);
});
