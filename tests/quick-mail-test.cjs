const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const BEARER = fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();
const adapter = spawn('node', [path.join(__dirname, '..', 'mcp-adapter.cjs')], {
    env: { ...process.env, MCP_SERVER_URL: 'http://localhost:3000', MCP_BEARER_TOKEN: BEARER },
    stdio: ['pipe', 'pipe', 'inherit']
});

let id = 0;
adapter.stdout.on('data', d => {
    try {
        const r = JSON.parse(d.toString());
        if (r.result?.content) {
            const text = r.result.content[0]?.text;
            if (text) {
                const data = JSON.parse(text);
                console.log('Result:', JSON.stringify(data, null, 2).substring(0, 1500));
            }
        } else if (r.result) {
            console.log('Init OK');
        }
    } catch (e) {}
});

setTimeout(() => {
    adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1.0' } }, id: id++ }) + '\n');
}, 1000);

setTimeout(() => {
    console.log('Calling searchMail...');
    adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'tools/call', params: { name: 'searchMail', arguments: { query: 'from:kristerm@microsoft.com', limit: 5 } }, id: id++ }) + '\n');
}, 3000);

setTimeout(() => {
    adapter.stdin.end();
    adapter.kill();
    process.exit(0);
}, 20000);
