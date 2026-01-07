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

rl.on('line', line => {
    try {
        const r = JSON.parse(line);
        if (r.result?.tools) {
            const tools = r.result.tools;
            const search = tools.find(t => t.name === 'search');
            const searchMail = tools.find(t => t.name === 'searchMail');
            const searchFiles = tools.find(t => t.name === 'searchFiles');
            console.log('\n=== Tool Check ===');
            console.log('Total tools:', tools.length);
            console.log('search:', search ? 'FOUND ✅' : 'NOT FOUND ❌');
            console.log('searchMail:', searchMail ? 'STILL EXISTS ❌' : 'REMOVED ✅');
            console.log('searchFiles:', searchFiles ? 'STILL EXISTS ❌' : 'REMOVED ✅');
            if (search) {
                console.log('\nSearch tool schema:');
                console.log(JSON.stringify(search, null, 2));
            }
            adapter.kill();
            process.exit(0);
        }
    } catch (e) {}
});

adapter.stderr.on('data', d => {
    const s = d.toString();
    if (!s.includes('Token')) console.log('[stderr]', s.substring(0, 100));
});

setTimeout(() => adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '1.0' } }, id: 0 }) + '\n'), 1000);
setTimeout(() => adapter.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'tools/list', params: {}, id: 1 }) + '\n'), 2000);
setTimeout(() => { console.log('Timeout'); process.exit(1); }, 10000);
