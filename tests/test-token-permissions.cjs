#!/usr/bin/env node
/**
 * @fileoverview Test script for validating Microsoft Graph token permissions
 *
 * Usage:
 *   node tests/test-token-permissions.cjs <token>
 *   node tests/test-token-permissions.cjs  # prompts for token
 *
 * This script:
 *   1. Decodes the JWT token to show granted scopes
 *   2. Tests various Microsoft Graph API endpoints
 *   3. Reports which permissions are working
 */

const https = require('https');
const readline = require('readline');

// ANSI colors for output
const colors = {
  reset: '\x1b[0m',
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  cyan: '\x1b[36m',
  dim: '\x1b[2m',
  bold: '\x1b[1m'
};

/**
 * Decode a JWT token (without verification - just for inspection)
 */
function decodeJwt(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) {
      throw new Error('Invalid JWT format');
    }

    const header = JSON.parse(Buffer.from(parts[0], 'base64url').toString());
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());

    return { header, payload };
  } catch (error) {
    throw new Error(`Failed to decode JWT: ${error.message}`);
  }
}

/**
 * Make a request to Microsoft Graph API
 */
function graphRequest(endpoint, token, method = 'GET') {
  return new Promise((resolve, reject) => {
    const url = new URL(endpoint.startsWith('http') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`);

    const options = {
      hostname: url.hostname,
      path: url.pathname + url.search,
      method,
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = data ? JSON.parse(data) : {};
          resolve({
            status: res.statusCode,
            data: json,
            success: res.statusCode >= 200 && res.statusCode < 300
          });
        } catch {
          resolve({
            status: res.statusCode,
            data: data,
            success: res.statusCode >= 200 && res.statusCode < 300
          });
        }
      });
    });

    req.on('error', reject);
    req.setTimeout(10000, () => {
      req.destroy();
      reject(new Error('Request timeout'));
    });
    req.end();
  });
}

/**
 * Test endpoints mapped to permissions
 */
const permissionTests = [
  // User permissions
  { endpoint: '/me', permission: 'User.Read', description: 'Read user profile' },
  { endpoint: '/me/photo/$value', permission: 'User.Read', description: 'Read user photo' },
  { endpoint: '/users?$top=1', permission: 'User.Read.All', description: 'List users' },

  // Mail permissions
  { endpoint: '/me/messages?$top=1', permission: 'Mail.Read / Mail.ReadWrite', description: 'Read mail' },
  { endpoint: '/me/mailFolders', permission: 'Mail.Read / Mail.ReadWrite', description: 'List mail folders' },

  // Calendar permissions
  { endpoint: '/me/events?$top=1', permission: 'Calendars.Read / Calendars.ReadWrite', description: 'Read calendar events' },
  { endpoint: '/me/calendars', permission: 'Calendars.Read / Calendars.ReadWrite', description: 'List calendars' },

  // Files permissions
  { endpoint: '/me/drive', permission: 'Files.Read / Files.ReadWrite', description: 'Access OneDrive' },
  { endpoint: '/me/drive/root/children?$top=5', permission: 'Files.Read / Files.ReadWrite', description: 'List OneDrive files' },

  // People permissions
  { endpoint: '/me/people?$top=5', permission: 'People.Read', description: 'Read relevant people' },

  // Contacts permissions
  { endpoint: '/me/contacts?$top=1', permission: 'Contacts.Read / Contacts.ReadWrite', description: 'Read contacts' },

  // Groups permissions
  { endpoint: '/me/memberOf?$top=1', permission: 'Directory.Read.All / Group.Read.All', description: 'Read group memberships' },

  // Tasks permissions
  { endpoint: '/me/todo/lists', permission: 'Tasks.Read / Tasks.ReadWrite', description: 'Read To Do lists' },

  // Teams permissions
  { endpoint: '/me/joinedTeams', permission: 'Team.ReadBasic.All', description: 'List joined teams' },
  { endpoint: '/me/chats?$top=1', permission: 'Chat.Read', description: 'List chats' },
  { endpoint: '/me/onlineMeetings?$top=1', permission: 'OnlineMeetings.Read', description: 'List online meetings' },

  // Notes permissions (OneNote)
  { endpoint: '/me/onenote/notebooks', permission: 'Notes.Read / Notes.Create', description: 'Read OneNote notebooks' },
];

/**
 * Print token information
 */
function printTokenInfo(decoded) {
  const { payload } = decoded;

  console.log(`\n${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                    TOKEN INFORMATION${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}═══════════════════════════════════════════════════════════════${colors.reset}\n`);

  // Basic info
  console.log(`${colors.bold}User:${colors.reset}        ${payload.name || 'N/A'} (${payload.upn || payload.unique_name || 'N/A'})`);
  console.log(`${colors.bold}Tenant:${colors.reset}      ${payload.tid}`);
  console.log(`${colors.bold}App:${colors.reset}         ${payload.app_displayname || 'N/A'} (${payload.appid})`);
  console.log(`${colors.bold}Audience:${colors.reset}    ${payload.aud}`);

  // Token validity
  const now = Math.floor(Date.now() / 1000);
  const issuedAt = new Date(payload.iat * 1000).toISOString();
  const expiresAt = new Date(payload.exp * 1000).toISOString();
  const isExpired = payload.exp < now;
  const timeLeft = payload.exp - now;

  console.log(`${colors.bold}Issued:${colors.reset}      ${issuedAt}`);
  console.log(`${colors.bold}Expires:${colors.reset}     ${expiresAt} ${isExpired ? colors.red + '(EXPIRED)' : colors.green + `(${Math.floor(timeLeft / 60)} min left)`}${colors.reset}`);

  // Scopes
  console.log(`\n${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                    GRANTED SCOPES (${(payload.scp || '').split(' ').length})${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}\n`);

  if (payload.scp) {
    const scopes = payload.scp.split(' ').sort();

    // Group scopes by category
    const categories = {
      'User': [],
      'Mail': [],
      'Calendar': [],
      'Files': [],
      'People': [],
      'Contacts': [],
      'Groups': [],
      'Directory': [],
      'Tasks': [],
      'Teams': [],
      'Notes': [],
      'Print': [],
      'Reports': [],
      'Sensitivity': [],
      'Other': []
    };

    scopes.forEach(scope => {
      let categorized = false;
      for (const cat of Object.keys(categories)) {
        if (scope.toLowerCase().startsWith(cat.toLowerCase()) ||
            scope.toLowerCase().includes(cat.toLowerCase())) {
          categories[cat].push(scope);
          categorized = true;
          break;
        }
      }
      if (!categorized) {
        categories['Other'].push(scope);
      }
    });

    for (const [category, categoryScopes] of Object.entries(categories)) {
      if (categoryScopes.length > 0) {
        console.log(`${colors.bold}${category}:${colors.reset}`);
        categoryScopes.forEach(scope => {
          console.log(`  ${colors.green}✓${colors.reset} ${scope}`);
        });
        console.log();
      }
    }
  } else {
    console.log(`${colors.yellow}No scopes found in token${colors.reset}`);
  }

  return payload;
}

/**
 * Test permissions against Microsoft Graph
 */
async function testPermissions(token, payload) {
  console.log(`${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                    PERMISSION TESTS${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}\n`);

  // Check if token is expired
  const now = Math.floor(Date.now() / 1000);
  if (payload.exp < now) {
    console.log(`${colors.red}Token is expired. Cannot test permissions.${colors.reset}\n`);
    return;
  }

  const results = {
    passed: [],
    failed: [],
    errors: []
  };

  for (const test of permissionTests) {
    process.stdout.write(`Testing ${test.description.padEnd(30)} `);

    try {
      const result = await graphRequest(test.endpoint, token);

      if (result.success) {
        console.log(`${colors.green}✓ PASS${colors.reset} ${colors.dim}(${result.status})${colors.reset}`);
        results.passed.push(test);
      } else {
        const errorCode = result.data?.error?.code || result.status;
        console.log(`${colors.red}✗ FAIL${colors.reset} ${colors.dim}(${errorCode})${colors.reset}`);
        results.failed.push({ ...test, error: errorCode });
      }
    } catch (error) {
      console.log(`${colors.yellow}? ERROR${colors.reset} ${colors.dim}(${error.message})${colors.reset}`);
      results.errors.push({ ...test, error: error.message });
    }

    // Small delay to avoid rate limiting
    await new Promise(resolve => setTimeout(resolve, 100));
  }

  // Summary
  console.log(`\n${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}                    SUMMARY${colors.reset}`);
  console.log(`${colors.bold}${colors.cyan}───────────────────────────────────────────────────────────────${colors.reset}\n`);

  console.log(`${colors.green}Passed:${colors.reset} ${results.passed.length}`);
  console.log(`${colors.red}Failed:${colors.reset} ${results.failed.length}`);
  console.log(`${colors.yellow}Errors:${colors.reset} ${results.errors.length}`);

  if (results.failed.length > 0) {
    console.log(`\n${colors.bold}Failed tests (missing permissions):${colors.reset}`);
    results.failed.forEach(test => {
      console.log(`  ${colors.red}✗${colors.reset} ${test.permission} - ${test.description}`);
    });
  }

  console.log();
}

/**
 * Prompt for token input
 */
async function promptForToken() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  return new Promise(resolve => {
    console.log(`\n${colors.bold}Microsoft Graph Token Permission Tester${colors.reset}\n`);
    console.log('Paste your JWT token (it will be hidden):');

    // Hide input
    process.stdout.write('Token: ');

    let token = '';
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.setEncoding('utf8');

    process.stdin.on('data', (char) => {
      if (char === '\n' || char === '\r' || char === '\u0004') {
        process.stdin.setRawMode(false);
        process.stdout.write('\n');
        rl.close();
        resolve(token.trim());
      } else if (char === '\u0003') {
        // Ctrl+C
        process.exit();
      } else if (char === '\u007F' || char === '\b') {
        // Backspace
        token = token.slice(0, -1);
      } else {
        token += char;
      }
    });
  });
}

/**
 * Main function
 */
async function main() {
  let token = process.argv[2];

  if (!token) {
    // Check if stdin has data (piped input)
    if (!process.stdin.isTTY) {
      const chunks = [];
      for await (const chunk of process.stdin) {
        chunks.push(chunk);
      }
      token = Buffer.concat(chunks).toString().trim();
    } else {
      token = await promptForToken();
    }
  }

  if (!token) {
    console.error(`${colors.red}Error: No token provided${colors.reset}`);
    console.log(`\nUsage: node tests/test-token-permissions.cjs <token>`);
    console.log(`       echo "token" | node tests/test-token-permissions.cjs`);
    process.exit(1);
  }

  try {
    // Decode and display token info
    const decoded = decodeJwt(token);
    const payload = printTokenInfo(decoded);

    // Test permissions
    await testPermissions(token, payload);

  } catch (error) {
    console.error(`${colors.red}Error: ${error.message}${colors.reset}`);
    process.exit(1);
  }
}

// Run
main().catch(console.error);
