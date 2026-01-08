# Claude Code Memory - MCP Microsoft Office

## E2E Testing Workflow

### Prerequisites
1. **Fresh Graph Token**: Copy a valid Microsoft Graph token to `token.txt` in project root
2. **Local Server**: Must be running on localhost:3000

### Step-by-Step Testing

```bash
# 1. Start the local server
npm run dev:web

# 2. Exchange Graph token for MCP bearer token
node tests/auth-helper.cjs
# This reads token.txt and saves MCP token to tests/mcp-bearer-token.txt

# 3. Run E2E tests (examples)
node tests/test-calendar-today.cjs    # Calendar date filtering
node tests/quick-mail-test.cjs        # Mail search
node tests/e2e-search-test.cjs        # Unified search
```

### Test Structure

Tests use the MCP adapter pattern:
```javascript
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

// Load MCP bearer token (NOT the raw Graph token)
const BEARER = fs.readFileSync(path.join(__dirname, 'mcp-bearer-token.txt'), 'utf8').trim();

// Spawn MCP adapter connected to local server
const adapter = spawn('node', [path.join(__dirname, '..', 'mcp-adapter.cjs')], {
    env: { ...process.env, MCP_SERVER_URL: 'http://localhost:3000', MCP_BEARER_TOKEN: BEARER },
    stdio: ['pipe', 'pipe', 'inherit']
});

// Send JSON-RPC messages via stdin, receive via stdout
adapter.stdin.write(JSON.stringify({
    jsonrpc: '2.0',
    method: 'tools/call',
    params: { name: 'getEvents', arguments: { timeframe: 'today' } },
    id: 1
}) + '\n');
```

### Important Notes

- **Token Flow**: Graph token → auth-helper.cjs → MCP bearer token → tests
- **tests/ is gitignored**: Test files stay local, don't try to commit them
- **Server logs**: Check server output for debugging (parameters passed, Graph API calls)
- **Multi-day events**: calendarView API returns events that OVERLAP with date range, not just events starting on that date

### Common Issues

1. **"Authentication required"**: Run `node tests/auth-helper.cjs` to refresh MCP token
2. **"No stored token"**: Ensure token.txt has a valid Graph token, then run auth-helper
3. **Wrong events returned**: Check if date params are being passed through the full chain (controller → module → service → Graph API)
