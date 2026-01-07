/**
 * E2E Test for Mail Search
 * Tests the searchMail functionality with real Graph API calls
 */

const fs = require('fs');
const path = require('path');
const https = require('https');

// Read the Graph token from token.txt
const tokenPath = path.join(__dirname, '..', 'token.txt');
const graphToken = fs.readFileSync(tokenPath, 'utf8').trim();

if (!graphToken) {
    console.error('No token found in token.txt');
    process.exit(1);
}

console.log('Token loaded, length:', graphToken.length);

// Test queries - same as Claude would send
const testQueries = [
    'from:kristerm@microsoft.com',
    'subject:meeting',
    'Krister'
];

async function testGraphSearch(query) {
    return new Promise((resolve, reject) => {
        // Build the search URL - this is what the fixed code does
        const cleanQuery = query.trim();
        const searchUrl = `/v1.0/me/messages?$search="${encodeURIComponent(cleanQuery)}"&$top=5`;

        console.log('\n--- Testing query:', query);
        console.log('URL:', searchUrl);

        const options = {
            hostname: 'graph.microsoft.com',
            path: searchUrl,
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${graphToken}`,
                'Content-Type': 'application/json'
            }
        };

        const req = https.request(options, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                console.log('Status:', res.statusCode);

                if (res.statusCode === 200) {
                    try {
                        const result = JSON.parse(data);
                        const emails = result.value || [];
                        console.log('Results:', emails.length, 'emails found');

                        if (emails.length > 0) {
                            console.log('First email:');
                            console.log('  Subject:', emails[0].subject);
                            console.log('  From:', emails[0].from?.emailAddress?.name || 'N/A');
                            console.log('  Date:', emails[0].receivedDateTime);
                        }
                        resolve({ success: true, count: emails.length, emails });
                    } catch (e) {
                        console.error('Parse error:', e.message);
                        resolve({ success: false, error: e.message });
                    }
                } else {
                    console.error('Error response:', data.substring(0, 500));
                    resolve({ success: false, status: res.statusCode, error: data });
                }
            });
        });

        req.on('error', (e) => {
            console.error('Request error:', e.message);
            resolve({ success: false, error: e.message });
        });

        req.end();
    });
}

async function runTests() {
    console.log('=== E2E Mail Search Test ===\n');

    const results = [];

    for (const query of testQueries) {
        const result = await testGraphSearch(query);
        results.push({ query, ...result });
    }

    console.log('\n=== Summary ===');
    for (const r of results) {
        const status = r.success ? '✅' : '❌';
        console.log(`${status} "${r.query}" - ${r.count || 0} results`);
    }

    const allPassed = results.every(r => r.success);
    process.exit(allPassed ? 0 : 1);
}

runTests();
