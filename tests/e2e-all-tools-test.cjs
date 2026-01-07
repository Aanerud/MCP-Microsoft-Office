/**
 * E2E Test for ALL new tools
 * Tests: Todo, Contacts, Groups, Teams modules
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

const ADAPTER_PATH = path.join(__dirname, '..', 'mcp-adapter.cjs');
const MCP_SERVER_URL = process.env.MCP_SERVER_URL || 'http://localhost:3000';
const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN || fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();

console.log('=== E2E All Tools Test ===');
console.log('Server:', MCP_SERVER_URL);

let messageId = 0;
const pendingRequests = new Map();

const adapter = spawn('node', [ADAPTER_PATH], {
    env: { ...process.env, MCP_SERVER_URL, MCP_BEARER_TOKEN },
    stdio: ['pipe', 'pipe', 'pipe']
});

adapter.stderr.on('data', (data) => {
    const output = data.toString().trim();
    if (output && !output.includes('Token refresh')) {
        console.log('[STDERR]', output.substring(0, 100));
    }
});

const rl = readline.createInterface({ input: adapter.stdout, crlfDelay: Infinity });

rl.on('line', (line) => {
    try {
        const response = JSON.parse(line);
        const request = pendingRequests.get(response.id);
        if (request) {
            pendingRequests.delete(response.id);
            request.resolve(response);
        }
    } catch (e) { /* ignore */ }
});

function sendRequest(method, params = {}) {
    return new Promise((resolve, reject) => {
        const id = messageId++;
        const timeout = setTimeout(() => {
            pendingRequests.delete(id);
            reject(new Error('Timeout'));
        }, 30000);

        pendingRequests.set(id, {
            resolve: (response) => { clearTimeout(timeout); resolve(response); }
        });

        adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params, id }) + '\n');
    });
}

// All tests organized by module
const tests = [
    // Initialize
    { name: 'Initialize', method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1.0' } } },

    // === TODO MODULE ===
    { name: 'TODO: listTaskLists', method: 'tools/call', params: { name: 'listTaskLists', arguments: {} } },
    { name: 'TODO: listTasks', method: 'tools/call', params: { name: 'listTasks', arguments: {} } },

    // === CONTACTS MODULE ===
    { name: 'CONTACTS: listContacts', method: 'tools/call', params: { name: 'listContacts', arguments: {} } },
    { name: 'CONTACTS: searchContacts', method: 'tools/call', params: { name: 'searchContacts', arguments: { query: 'John' } } },

    // === GROUPS MODULE ===
    { name: 'GROUPS: listMyGroups', method: 'tools/call', params: { name: 'listMyGroups', arguments: {} } },
    { name: 'GROUPS: listGroups', method: 'tools/call', params: { name: 'listGroups', arguments: { limit: 5 } } },

    // === TEAMS MODULE ===
    { name: 'TEAMS: listChats', method: 'tools/call', params: { name: 'listChats', arguments: { limit: 5 } } },
    { name: 'TEAMS: listJoinedTeams', method: 'tools/call', params: { name: 'listJoinedTeams', arguments: {} } },
    { name: 'TEAMS: listOnlineMeetings', method: 'tools/call', params: { name: 'listOnlineMeetings', arguments: { limit: 5 } } },

    // === MAIL (verify search still works) ===
    { name: 'MAIL: searchMail', method: 'tools/call', params: { name: 'searchMail', arguments: { query: 'from:kristerm@microsoft.com', limit: 3 } } },

    // === CALENDAR ===
    { name: 'CALENDAR: getEvents', method: 'tools/call', params: { name: 'getEvents', arguments: { limit: 3 } } },

    // === FILES ===
    { name: 'FILES: listFiles', method: 'tools/call', params: { name: 'listFiles', arguments: {} } },

    // === PEOPLE ===
    { name: 'PEOPLE: getRelevantPeople', method: 'tools/call', params: { name: 'getRelevantPeople', arguments: { limit: 3 } } },
];

async function runTests() {
    console.log('\n--- Running Tests ---\n');
    const results = [];

    for (const test of tests) {
        process.stdout.write(`[${test.name}] `);

        try {
            const response = await sendRequest(test.method, test.params);

            if (response.error) {
                console.log(`❌ ${response.error.message.substring(0, 60)}`);
                results.push({ name: test.name, success: false, error: response.error.message });
            } else {
                const resultStr = JSON.stringify(response.result);
                const isMock = resultStr.includes('mock') || resultStr.includes('Mock');

                if (isMock) {
                    console.log('⚠️  MOCK DATA');
                    results.push({ name: test.name, success: false, error: 'Mock data' });
                } else {
                    // Try to extract count
                    let count = '';
                    try {
                        const content = response.result?.content?.[0]?.text;
                        if (content) {
                            const parsed = JSON.parse(content);
                            if (Array.isArray(parsed)) count = `(${parsed.length} items)`;
                            else if (parsed.taskLists) count = `(${parsed.taskLists.length} lists)`;
                            else if (parsed.tasks) count = `(${parsed.tasks.length} tasks)`;
                            else if (parsed.contacts) count = `(${parsed.contacts.length} contacts)`;
                            else if (parsed.groups) count = `(${parsed.groups.length} groups)`;
                            else if (parsed.chats) count = `(${parsed.chats.length} chats)`;
                            else if (parsed.teams) count = `(${parsed.teams.length} teams)`;
                            else if (parsed.meetings) count = `(${parsed.meetings.length} meetings)`;
                            else if (parsed.people) count = `(${parsed.people.length} people)`;
                            else if (parsed.files) count = `(${parsed.files.length} files)`;
                            else if (parsed.events) count = `(${parsed.events.length} events)`;
                        }
                    } catch (e) { /* ignore */ }

                    console.log(`✅ ${count}`);
                    results.push({ name: test.name, success: true });
                }
            }
        } catch (e) {
            console.log(`❌ ${e.message}`);
            results.push({ name: test.name, success: false, error: e.message });
        }
    }

    // Summary
    console.log('\n=== Summary ===');
    const passed = results.filter(r => r.success).length;
    const failed = results.filter(r => !r.success).length;

    console.log(`\nPassed: ${passed}/${results.length}`);
    if (failed > 0) {
        console.log('\nFailed tests:');
        results.filter(r => !r.success).forEach(r => {
            console.log(`  ❌ ${r.name}: ${r.error}`);
        });
    }

    adapter.stdin.end();
    adapter.kill();

    process.exit(failed > 0 ? 1 : 0);
}

setTimeout(runTests, 2000);
