/**
 * E2E Test using MCP Adapter
 * Tests the full flow: Test Script → MCP Adapter → MCP Server → Graph API
 * This mimics exactly how Claude Desktop uses the adapter
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

// Configuration - same as Claude Desktop would use
const ADAPTER_PATH = path.join(__dirname, '..', 'mcp-adapter.cjs');
const MCP_SERVER_URL = process.env.MCP_SERVER_URL || 'http://localhost:3000';

// Read MCP Bearer Token from environment or generate test token
const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN || '';

console.log('=== E2E Adapter Test ===');
console.log('Adapter:', ADAPTER_PATH);
console.log('Server:', MCP_SERVER_URL);
console.log('Token:', MCP_BEARER_TOKEN ? `${MCP_BEARER_TOKEN.substring(0, 20)}...` : 'NOT SET');

if (!MCP_BEARER_TOKEN) {
    console.error('\nERROR: MCP_BEARER_TOKEN not set. Get a token from the web UI and set it:');
    console.error('  export MCP_BEARER_TOKEN="your-token-here"');
    process.exit(1);
}

// JSON-RPC message ID counter
let messageId = 0;

// Start the adapter process
const adapter = spawn('node', [ADAPTER_PATH], {
    env: {
        ...process.env,
        MCP_SERVER_URL,
        MCP_BEARER_TOKEN
    },
    stdio: ['pipe', 'pipe', 'pipe']
});

// Capture stderr for debugging
adapter.stderr.on('data', (data) => {
    const output = data.toString().trim();
    if (output) {
        console.log('[ADAPTER STDERR]', output);
    }
});

// Parse JSON-RPC responses from stdout
const rl = readline.createInterface({
    input: adapter.stdout,
    crlfDelay: Infinity
});

const pendingRequests = new Map();

rl.on('line', (line) => {
    try {
        const response = JSON.parse(line);
        const request = pendingRequests.get(response.id);
        if (request) {
            pendingRequests.delete(response.id);
            request.resolve(response);
        }
    } catch (e) {
        console.error('[PARSE ERROR]', e.message, 'Line:', line.substring(0, 100));
    }
});

// Send JSON-RPC request
function sendRequest(method, params = {}) {
    return new Promise((resolve, reject) => {
        const id = messageId++;
        const request = {
            jsonrpc: '2.0',
            method,
            params,
            id
        };

        const timeout = setTimeout(() => {
            pendingRequests.delete(id);
            reject(new Error(`Request ${method} timed out`));
        }, 30000);

        pendingRequests.set(id, {
            resolve: (response) => {
                clearTimeout(timeout);
                resolve(response);
            },
            reject
        });

        adapter.stdin.write(JSON.stringify(request) + '\n');
    });
}

// Test scenarios - exactly what Claude sends
const tests = [
    {
        name: 'Initialize',
        method: 'initialize',
        params: {
            protocolVersion: '2025-06-18',
            capabilities: {},
            clientInfo: { name: 'e2e-test', version: '1.0.0' }
        }
    },
    {
        name: 'List Tools',
        method: 'tools/list',
        params: {}
    },
    {
        name: 'Search Mail (from:kristerm)',
        method: 'tools/call',
        params: {
            name: 'searchMail',
            arguments: {
                query: 'from:kristerm@microsoft.com',
                limit: 5
            }
        }
    },
    {
        name: 'Read Mail (inbox)',
        method: 'tools/call',
        params: {
            name: 'readMail',
            arguments: {}
        }
    }
];

async function runTests() {
    console.log('\n--- Starting Tests ---\n');

    const results = [];

    for (const test of tests) {
        console.log(`\n[TEST] ${test.name}`);
        console.log(`  Method: ${test.method}`);
        if (test.params.name) console.log(`  Tool: ${test.params.name}`);

        try {
            const response = await sendRequest(test.method, test.params);

            if (response.error) {
                console.log(`  ❌ ERROR: ${response.error.message}`);
                results.push({ name: test.name, success: false, error: response.error.message });
            } else {
                // Check for mock data
                const resultStr = JSON.stringify(response.result);
                const isMock = resultStr.includes('search-mock') || resultStr.includes('Search System');

                if (isMock) {
                    console.log(`  ⚠️  MOCK DATA returned (search fix needed)`);
                    results.push({ name: test.name, success: false, error: 'Mock data returned' });
                } else {
                    console.log(`  ✅ SUCCESS`);

                    // Show preview of result
                    if (test.method === 'tools/call') {
                        try {
                            const content = response.result?.content?.[0]?.text;
                            if (content) {
                                const parsed = JSON.parse(content);
                                if (Array.isArray(parsed) && parsed.length > 0) {
                                    console.log(`  📧 Found ${parsed.length} items`);
                                    console.log(`     First: ${parsed[0].subject || parsed[0].displayName || 'N/A'}`);
                                }
                            }
                        } catch (e) { /* Ignore preview errors */ }
                    }

                    results.push({ name: test.name, success: true });
                }
            }
        } catch (e) {
            console.log(`  ❌ TIMEOUT/ERROR: ${e.message}`);
            results.push({ name: test.name, success: false, error: e.message });
        }
    }

    // Summary
    console.log('\n=== Summary ===');
    for (const r of results) {
        const status = r.success ? '✅' : '❌';
        console.log(`${status} ${r.name}${r.error ? ` (${r.error})` : ''}`);
    }

    // Cleanup
    adapter.stdin.end();
    adapter.kill();

    const allPassed = results.every(r => r.success);
    console.log(`\n${allPassed ? '✅ All tests passed!' : '❌ Some tests failed'}`);
    process.exit(allPassed ? 0 : 1);
}

// Handle adapter exit
adapter.on('exit', (code) => {
    if (code !== 0 && code !== null) {
        console.error(`Adapter exited with code ${code}`);
    }
});

// Wait a moment for adapter to initialize, then run tests
setTimeout(runTests, 2000);
