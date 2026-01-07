/**
 * Quick test for unified search with "Krister"
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const readline = require('readline');

const ADAPTER_PATH = path.join(__dirname, '..', 'mcp-adapter.cjs');
const MCP_SERVER_URL = process.env.MCP_SERVER_URL || 'http://localhost:3000';
const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN || fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();

console.log('=== Krister Search Test ===');
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
        console.log('[STDERR]', output.substring(0, 150));
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

async function runTest() {
    console.log('\n--- Initializing ---');
    await sendRequest('initialize', {
        protocolVersion: '2025-06-18',
        capabilities: {},
        clientInfo: { name: 'test', version: '1.0' }
    });

    console.log('\n--- Searching for "Krister" ---\n');

    // Test 1: Unified search for emails from Krister (using 'search' tool)
    console.log('[Mail Search] Searching emails from kristerm@microsoft.com...');
    try {
        const mailResult = await sendRequest('tools/call', {
            name: 'search',
            arguments: { query: 'from:kristerm@microsoft.com', entityTypes: ['message'], limit: 10 }
        });

        if (mailResult.error) {
            console.log('  ❌ Error:', mailResult.error.message);
        } else {
            const content = JSON.parse(mailResult.result?.content?.[0]?.text || '{}');
            const emails = content.results || content.emails || [];
            console.log(`  ✅ Found ${emails.length} emails`);
            emails.slice(0, 3).forEach((email, i) => {
                console.log(`     ${i+1}. "${email.subject}" from ${email.from?.email || email.from?.address || 'unknown'}`);
            });
        }
    } catch (e) {
        console.log('  ❌ Error:', e.message);
    }

    // Test 1b: Unified search for emails containing "Krister"
    console.log('\n[Mail Search 2] Searching emails containing "Krister"...');
    try {
        const mailResult2 = await sendRequest('tools/call', {
            name: 'search',
            arguments: { query: 'Krister', entityTypes: ['message'], limit: 5 }
        });

        if (mailResult2.error) {
            console.log('  ❌ Error:', mailResult2.error.message);
        } else {
            const content = JSON.parse(mailResult2.result?.content?.[0]?.text || '{}');
            const emails = content.results || content.emails || [];
            console.log(`  ✅ Found ${emails.length} emails mentioning "Krister"`);
        }
    } catch (e) {
        console.log('  ❌ Error:', e.message);
    }

    // Test 2: getEvents to find meetings with Krister
    console.log('\n[Calendar Search] Getting events (check for Krister meetings)...');
    try {
        const calResult = await sendRequest('tools/call', {
            name: 'getEvents',
            arguments: { limit: 20 }
        });

        if (calResult.error) {
            console.log('  ❌ Error:', calResult.error.message);
        } else {
            const content = JSON.parse(calResult.result?.content?.[0]?.text || '{}');
            const events = content.events || content.results || [];
            const kristerEvents = events.filter(e =>
                JSON.stringify(e).toLowerCase().includes('krister')
            );
            console.log(`  ✅ Found ${events.length} events total, ${kristerEvents.length} with "Krister"`);
            kristerEvents.slice(0, 3).forEach((event, i) => {
                console.log(`     ${i+1}. "${event.subject}" on ${event.start?.dateTime || event.start}`);
            });
        }
    } catch (e) {
        console.log('  ❌ Error:', e.message);
    }

    // Test 3: findPeople for Krister
    console.log('\n[People Search] Finding people named "Krister"...');
    try {
        const peopleResult = await sendRequest('tools/call', {
            name: 'findPeople',
            arguments: { name: 'Krister', limit: 5 }
        });

        if (peopleResult.error) {
            console.log('  ❌ Error:', peopleResult.error.message);
        } else {
            const content = JSON.parse(peopleResult.result?.content?.[0]?.text || '{}');
            const people = content.people || content.results || [];
            console.log(`  ✅ Found ${people.length} people`);
            people.slice(0, 3).forEach((person, i) => {
                console.log(`     ${i+1}. ${person.displayName || person.name} - ${person.jobTitle || 'N/A'}`);
            });
        }
    } catch (e) {
        console.log('  ❌ Error:', e.message);
    }

    // Test 4: Unified search for files containing "Krister"
    console.log('\n[Files Search] Searching files for "Krister"...');
    try {
        const filesResult = await sendRequest('tools/call', {
            name: 'search',
            arguments: { query: 'Krister', entityTypes: ['driveItem'], limit: 5 }
        });

        if (filesResult.error) {
            console.log('  ❌ Error:', filesResult.error.message);
        } else {
            const content = JSON.parse(filesResult.result?.content?.[0]?.text || '{}');
            const files = content.results || content.files || [];
            console.log(`  ✅ Found ${files.length} files`);
            files.slice(0, 3).forEach((file, i) => {
                console.log(`     ${i+1}. ${file.name} (${file.webUrl ? 'has URL' : 'no URL'})`);
            });
        }
    } catch (e) {
        console.log('  ❌ Error:', e.message);
    }

    console.log('\n=== Test Complete ===');
    adapter.stdin.end();
    adapter.kill();
    process.exit(0);
}

setTimeout(runTest, 2000);
