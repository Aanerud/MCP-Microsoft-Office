/**
 * Helper script to authenticate and get MCP bearer token for testing
 */

const fs = require('fs');
const path = require('path');
const http = require('http');

const tokenPath = path.join(__dirname, '..', 'token.txt');
const graphToken = fs.readFileSync(tokenPath, 'utf8').trim();

console.log('Graph token length:', graphToken.length);

// Store cookies for session
let cookies = [];

function makeRequest(method, path, data = null) {
    return new Promise((resolve, reject) => {
        const postData = data ? JSON.stringify(data) : null;

        const options = {
            hostname: 'localhost',
            port: 3000,
            path: path,
            method: method,
            headers: {
                'Content-Type': 'application/json',
                'Cookie': cookies.join('; ')
            }
        };

        if (postData) {
            options.headers['Content-Length'] = Buffer.byteLength(postData);
        }

        const req = http.request(options, (res) => {
            // Store cookies from response
            const setCookies = res.headers['set-cookie'];
            if (setCookies) {
                cookies = setCookies.map(c => c.split(';')[0]);
            }

            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                try {
                    resolve({ status: res.statusCode, data: JSON.parse(data) });
                } catch (e) {
                    resolve({ status: res.statusCode, data: data });
                }
            });
        });

        req.on('error', reject);
        if (postData) req.write(postData);
        req.end();
    });
}

async function main() {
    try {
        // Step 1: Login with external token
        console.log('\n1. Logging in with Graph token...');
        const loginResult = await makeRequest('POST', '/api/auth/external-token/login', { access_token: graphToken });
        console.log('   Status:', loginResult.status);

        if (loginResult.status !== 200 || !loginResult.data.success) {
            console.error('   Login failed:', loginResult.data);
            process.exit(1);
        }

        console.log('   Authenticated as:', loginResult.data.user?.name);

        // Step 2: Generate MCP token
        console.log('\n2. Generating MCP Bearer Token...');
        const tokenResult = await makeRequest('POST', '/api/auth/generate-mcp-token');
        console.log('   Status:', tokenResult.status);

        if (tokenResult.status === 200 && (tokenResult.data.token || tokenResult.data.access_token)) {
            const mcpToken = tokenResult.data.token || tokenResult.data.access_token;
            console.log('\n=== MCP Bearer Token ===');
            console.log(mcpToken);
            console.log('\nExport with:');
            console.log(`export MCP_BEARER_TOKEN="${mcpToken}"`);

            // Write to file for easy use
            const tokenFile = path.join(__dirname, 'mcp-bearer-token.txt');
            fs.writeFileSync(tokenFile, mcpToken);
            console.log(`\nToken saved to: ${tokenFile}`);
        } else {
            console.error('   Failed to generate token:', tokenResult.data);
        }

    } catch (e) {
        console.error('Error:', e.message);
        process.exit(1);
    }
}

main();
