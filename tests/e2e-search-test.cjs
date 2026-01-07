/**
 * E2E Test Script for MCP Microsoft Office
 * Tests authentication, mail retrieval, and unified search
 */

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');

// Configuration
const API_BASE = process.env.API_BASE || 'http://localhost:3000';
const TOKEN_FILE = path.join(__dirname, '..', 'token.txt');

// Read the MS Graph token
const graphToken = fs.readFileSync(TOKEN_FILE, 'utf8').trim();

// Test state
let mcpBearerToken = null;
let sessionCookies = [];
let testResults = { passed: 0, failed: 0, tests: [] };

/**
 * Make HTTP request with cookie support
 */
function makeRequest(method, url, body = null, headers = {}) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const isHttps = parsedUrl.protocol === 'https:';
    const client = isHttps ? https : http;

    const options = {
      hostname: parsedUrl.hostname,
      port: parsedUrl.port || (isHttps ? 443 : 80),
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
      options.headers['Content-Length'] = Buffer.byteLength(JSON.stringify(body));
    }

    const req = client.request(options, (res) => {
      // Store set-cookie headers
      const setCookies = res.headers['set-cookie'];
      if (setCookies) {
        setCookies.forEach(cookie => {
          const cookieName = cookie.split('=')[0];
          // Update or add cookie
          sessionCookies = sessionCookies.filter(c => !c.startsWith(cookieName + '='));
          sessionCookies.push(cookie.split(';')[0]);
        });
      }

      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, data: JSON.parse(data), headers: res.headers });
        } catch {
          resolve({ status: res.statusCode, data, headers: res.headers });
        }
      });
    });

    req.on('error', reject);
    if (body) req.write(JSON.stringify(body));
    req.end();
  });
}

/**
 * Log test result
 */
function logTest(name, passed, details = '') {
  const status = passed ? '✅ PASS' : '❌ FAIL';
  console.log(`${status}: ${name}`);
  if (details) console.log(`   ${details}`);
  testResults.tests.push({ name, passed, details });
  if (passed) testResults.passed++; else testResults.failed++;
}

/**
 * Test 1: Health check
 */
async function testHealth() {
  console.log('\n--- Test 1: Health Check ---');
  try {
    const res = await makeRequest('GET', `${API_BASE}/api/health`);
    logTest('Health endpoint', res.status === 200 && res.data.status === 'ok');
  } catch (err) {
    logTest('Health endpoint', false, err.message);
  }
}

/**
 * Test 2: Login with external Graph token
 */
async function testLogin() {
  console.log('\n--- Test 2: Login with External Token ---');
  try {
    const res = await makeRequest('POST', `${API_BASE}/api/auth/external-token/login`, {
      access_token: graphToken
    });

    if (res.status === 200 && res.data.success) {
      mcpBearerToken = graphToken; // Keep for fallback
      logTest('Login with Graph token', true, `Authenticated as ${res.data.user?.email || 'user'}`);
      console.log(`   Session cookies: ${sessionCookies.length > 0 ? sessionCookies.map(c => c.split('=')[0]).join(', ') : 'none'}`);
    } else {
      logTest('Login with Graph token', false, JSON.stringify(res.data));
    }
  } catch (err) {
    logTest('Login with Graph token', false, err.message);
  }
}

/**
 * Test 3: Get tools list
 */
async function testGetTools() {
  console.log('\n--- Test 3: Get Tools ---');
  try {
    // Tools endpoint doesn't require auth, but we pass session cookies anyway
    const res = await makeRequest('GET', `${API_BASE}/api/tools`);

    if (res.status === 200 && res.data.tools) {
      const toolNames = res.data.tools.map(t => t.name);
      const hasSearch = toolNames.includes('search');
      const hasReadMail = toolNames.includes('readMail');

      logTest('Get tools list', true, `${res.data.tools.length} tools found`);
      logTest('Has unified search tool', hasSearch, hasSearch ? 'search tool present' : 'search tool missing');
      logTest('Has readMail tool', hasReadMail);

      console.log('   Available search-related tools:', toolNames.filter(n => n.toLowerCase().includes('search')).join(', '));
    } else {
      logTest('Get tools list', false, JSON.stringify(res.data));
    }
  } catch (err) {
    logTest('Get tools list', false, err.message);
  }
}

/**
 * Test 4: Get emails and find specific one
 */
async function testGetMail() {
  console.log('\n--- Test 4: Get Mail ---');
  const targetSubject = '[EXTERNAL] Your order J21804300101 - PO 5100092434, has been Shipped.';

  try {
    // Uses session cookies for auth
    const res = await makeRequest('GET', `${API_BASE}/api/v1/mail?limit=50`);

    if (res.status === 200) {
      const emails = res.data.emails || res.data.messages || res.data;
      if (Array.isArray(emails)) {
        logTest('Get mail endpoint', true, `Got ${emails.length} emails`);

        // Look for the specific email
        const foundEmail = emails.find(e =>
          e.subject && e.subject.includes('J21804300101')
        );

        if (foundEmail) {
          logTest('Find specific email', true, `Found: "${foundEmail.subject.substring(0, 50)}..."`);
        } else {
          logTest('Find specific email', false, 'Email with J21804300101 not found in first 50');
          console.log('   Subjects found:', emails.slice(0, 5).map(e => e.subject?.substring(0, 40)).join(', '));
        }
      } else {
        logTest('Get mail endpoint', false, 'Response not an array');
      }
    } else {
      logTest('Get mail endpoint', false, `Status ${res.status}: ${JSON.stringify(res.data)}`);
    }
  } catch (err) {
    logTest('Get mail endpoint', false, err.message);
  }
}

/**
 * Test 5: Unified Search - search for files
 */
async function testUnifiedSearch() {
  console.log('\n--- Test 5: Unified Search ---');

  try {
    // Test search for files (driveItem) - uses session cookies
    const res = await makeRequest('POST', `${API_BASE}/api/v1/search`, {
      query: 'test',
      entityTypes: ['driveItem'],
      limit: 5
    });

    if (res.status === 200) {
      const results = res.data.results || [];
      logTest('Unified search (files)', true, `Got ${results.length} results`);
      if (results.length > 0) {
        console.log('   First result:', results[0].name || results[0].subject || 'N/A');
      }
    } else {
      logTest('Unified search (files)', false, `Status ${res.status}: ${JSON.stringify(res.data)}`);
    }

    // Test search for messages
    const res2 = await makeRequest('POST', `${API_BASE}/api/v1/search`, {
      query: 'order',
      entityTypes: ['message'],
      limit: 5
    });

    if (res2.status === 200) {
      const results = res2.data.results || [];
      logTest('Unified search (messages)', true, `Got ${results.length} results`);
      if (results.length > 0) {
        console.log('   First result subject:', results[0].subject?.substring(0, 50) || 'N/A');
      }
    } else {
      logTest('Unified search (messages)', false, `Status ${res2.status}: ${JSON.stringify(res2.data)}`);
    }

    // Test cross-entity search
    const res3 = await makeRequest('POST', `${API_BASE}/api/v1/search`, {
      query: 'project',
      entityTypes: ['message', 'driveItem'],
      limit: 5
    });

    if (res3.status === 200) {
      const results = res3.data.results || [];
      logTest('Unified search (cross-entity)', true, `Got ${results.length} results across message+driveItem`);
    } else {
      logTest('Unified search (cross-entity)', false, `Status ${res3.status}: ${JSON.stringify(res3.data)}`);
    }

  } catch (err) {
    logTest('Unified search', false, err.message);
  }
}

/**
 * Run all tests
 */
async function runTests() {
  console.log('='.repeat(60));
  console.log('MCP Microsoft Office - E2E Test Suite');
  console.log('='.repeat(60));
  console.log(`API Base: ${API_BASE}`);
  console.log(`Token file: ${TOKEN_FILE}`);

  await testHealth();
  await testLogin();

  if (!mcpBearerToken) {
    console.log('\n❌ Cannot continue without authentication');
    process.exit(1);
  }

  await testGetTools();
  await testGetMail();
  await testUnifiedSearch();

  // Summary
  console.log('\n' + '='.repeat(60));
  console.log('TEST SUMMARY');
  console.log('='.repeat(60));
  console.log(`✅ Passed: ${testResults.passed}`);
  console.log(`❌ Failed: ${testResults.failed}`);
  console.log(`📊 Total:  ${testResults.passed + testResults.failed}`);

  process.exit(testResults.failed > 0 ? 1 : 0);
}

// Run
runTests().catch(err => {
  console.error('Test runner error:', err);
  process.exit(1);
});
