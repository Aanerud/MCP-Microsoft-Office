const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const BEARER = fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();
const adapter = spawn('node', [path.join(__dirname, '..', 'mcp-adapter.cjs')], {
    env: { ...process.env, MCP_SERVER_URL: 'http://localhost:3000', MCP_BEARER_TOKEN: BEARER },
    stdio: ['pipe', 'pipe', 'pipe']
});

const rl = readline.createInterface({ input: adapter.stdout, crlfDelay: Infinity });

let id = 0;
const pending = new Map();

rl.on('line', line => {
    try {
        const r = JSON.parse(line);
        if (pending.has(r.id)) {
            pending.get(r.id)(r);
            pending.delete(r.id);
        }
    } catch (e) {}
});

function call(method, params) {
    return new Promise((resolve, reject) => {
        const reqId = id++;
        const timeout = setTimeout(() => { pending.delete(reqId); reject(new Error('Timeout')); }, 30000);
        pending.set(reqId, r => { clearTimeout(timeout); resolve(r); });
        adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params, id: reqId }) + '\n');
    });
}

async function run() {
    console.log('=== Unified Search Test ===\n');

    await call('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1.0' } });
    console.log('✅ Initialized\n');

    // Test 1: Search all types for "Krister"
    console.log('Test 1: Search all types for "Krister"...');
    const r1 = await call('tools/call', { name: 'search', arguments: { query: 'Krister', limit: 3 } });
    if (r1.error) {
        console.log('❌ Error:', r1.error.message);
    } else {
        const data = JSON.parse(r1.result.content[0].text);
        console.log(`✅ Found ${data.results?.length || 0} results across ${data.entityTypes?.join(', ')}`);
        (data.results || []).slice(0, 5).forEach(r => console.log(`   - [${r.entityType}] ${r.subject || r.displayName || r.name || r.id}`));
    }

    // Test 2: Search only emails
    console.log('\nTest 2: Search emails only for "from:kristerm@microsoft.com"...');
    const r2 = await call('tools/call', { name: 'search', arguments: { query: 'from:kristerm@microsoft.com', entityTypes: ['message'], limit: 5 } });
    if (r2.error) {
        console.log('❌ Error:', r2.error.message);
    } else {
        const data = JSON.parse(r2.result.content[0].text);
        console.log(`✅ Found ${data.results?.length || 0} emails`);
        (data.results || []).slice(0, 3).forEach(r => console.log(`   - ${r.subject} (from: ${r.from?.email})`));
    }

    // Test 3: Search files
    console.log('\nTest 3: Search files for "report"...');
    const r3 = await call('tools/call', { name: 'search', arguments: { query: 'report', entityTypes: ['driveItem'], limit: 5 } });
    if (r3.error) {
        console.log('❌ Error:', r3.error.message);
    } else {
        const data = JSON.parse(r3.result.content[0].text);
        console.log(`✅ Found ${data.results?.length || 0} files`);
        (data.results || []).slice(0, 3).forEach(r => console.log(`   - ${r.name}`));
    }

    // Test 4: Search people
    console.log('\nTest 4: Search people for "Krister"...');
    const r4 = await call('tools/call', { name: 'search', arguments: { query: 'Krister', entityTypes: ['person'], limit: 5 } });
    if (r4.error) {
        console.log('❌ Error:', r4.error.message);
    } else {
        const data = JSON.parse(r4.result.content[0].text);
        console.log(`✅ Found ${data.results?.length || 0} people`);
        (data.results || []).slice(0, 3).forEach(r => console.log(`   - ${r.displayName} (${r.jobTitle || 'N/A'})`));
    }

    console.log('\n=== Tests Complete ===');
    adapter.kill();
    process.exit(0);
}

adapter.stderr.on('data', () => {});
setTimeout(run, 2000);
