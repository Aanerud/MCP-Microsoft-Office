/**
 * Test unified search across Microsoft 365
 */

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const BEARER = fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();
const adapter = spawn('node', [path.join(__dirname, '..', 'mcp-adapter.cjs')], {
    env: { ...process.env, MCP_SERVER_URL: 'http://localhost:3000', MCP_BEARER_TOKEN: BEARER },
    stdio: ['pipe', 'pipe', 'inherit']
});

let id = 0;
const results = [];

adapter.stdout.on('data', d => {
    try {
        const r = JSON.parse(d.toString());
        if (r.result?.content) {
            const text = r.result.content[0]?.text;
            if (text) {
                const data = JSON.parse(text);
                results.push({ id: r.id, data });
            }
        } else if (r.result?.tools) {
            // Check if search tool exists
            const searchTool = r.result.tools.find(t => t.name === 'search');
            if (searchTool) {
                console.log('✅ Unified search tool found');
                console.log('   Description:', searchTool.description.substring(0, 80) + '...');
            } else {
                console.log('❌ search tool NOT found');
            }
            // Check that searchMail and searchFiles are removed
            const searchMail = r.result.tools.find(t => t.name === 'searchMail');
            const searchFiles = r.result.tools.find(t => t.name === 'searchFiles');
            if (!searchMail) console.log('✅ searchMail removed');
            else console.log('❌ searchMail still present');
            if (!searchFiles) console.log('✅ searchFiles removed');
            else console.log('❌ searchFiles still present');
        } else if (r.result) {
            console.log('✅ Init OK');
        }
    } catch (e) {}
});

console.log('=== Testing Unified Search ===\n');

setTimeout(() => {
    console.log('Initializing...');
    adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1.0' } }, id: id++ }) + '\n');
}, 1000);

setTimeout(() => {
    console.log('\nListing tools...');
    adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'tools/list', params: {}, id: id++ }) + '\n');
}, 2500);

setTimeout(() => {
    console.log('\n--- Test 1: Search all entity types for "Krister" ---');
    adapter.stdin.write(JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/call',
        params: {
            name: 'search',
            arguments: { query: 'Krister', limit: 5 }
        },
        id: id++
    }) + '\n');
}, 4000);

setTimeout(() => {
    console.log('\n--- Test 2: Search only emails for "from:kristerm" ---');
    adapter.stdin.write(JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/call',
        params: {
            name: 'search',
            arguments: {
                query: 'from:kristerm@microsoft.com',
                entityTypes: ['message'],
                limit: 5
            }
        },
        id: id++
    }) + '\n');
}, 7000);

setTimeout(() => {
    console.log('\n--- Test 3: Search files for "report" ---');
    adapter.stdin.write(JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/call',
        params: {
            name: 'search',
            arguments: {
                query: 'report',
                entityTypes: ['driveItem'],
                limit: 5
            }
        },
        id: id++
    }) + '\n');
}, 10000);

setTimeout(() => {
    console.log('\n--- Test 4: Search people for "Krister" ---');
    adapter.stdin.write(JSON.stringify({
        jsonrpc: '2.0',
        method: 'tools/call',
        params: {
            name: 'search',
            arguments: {
                query: 'Krister',
                entityTypes: ['person'],
                limit: 5
            }
        },
        id: id++
    }) + '\n');
}, 13000);

setTimeout(() => {
    console.log('\n=== Results Summary ===');
    results.forEach((r, i) => {
        if (r.data.results) {
            console.log(`\nTest ${i + 1}: Found ${r.data.results.length} results across ${r.data.entityTypes?.join(', ') || 'all'} types`);
            r.data.results.slice(0, 3).forEach(result => {
                console.log(`  - [${result.entityType}] ${result.subject || result.displayName || result.name || result.id}`);
            });
        } else {
            console.log(`\nTest ${i + 1}:`, JSON.stringify(r.data).substring(0, 200));
        }
    });

    adapter.stdin.end();
    adapter.kill();
    process.exit(0);
}, 18000);
