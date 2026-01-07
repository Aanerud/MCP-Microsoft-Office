#!/usr/bin/env node
/**
 * Test storage flow for external token
 */

const path = require('path');
const fs = require('fs');

// Change to project root
process.chdir(path.join(__dirname, '..'));

// Set required env vars
process.env.NODE_ENV = 'development';
process.env.MICROSOFT_CLIENT_ID = 'test';
process.env.JWT_SECRET = 'test-secret-for-local-testing';

async function main() {
  console.log('\n=== Testing Storage Flow ===\n');

  // Read token
  const token = fs.readFileSync('token.txt', 'utf8').trim();
  const parts = token.split('.');
  const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());

  const userEmail = payload.upn || payload.unique_name;
  const userId = `ms365:${userEmail}`;

  console.log('User email:', userEmail);
  console.log('User ID:', userId);
  console.log('');

  // Initialize storage service
  const StorageService = require('../src/core/storage-service.cjs');
  await StorageService.init();
  console.log('Storage initialized\n');

  // Test key generation
  const STORAGE_KEYS = {
    TOKEN: 'external-graph-token',
    METADATA: 'external-token-metadata',
    SOURCE: 'token-source'
  };

  const sourceKey = `${userId}:${STORAGE_KEYS.SOURCE}`;
  const tokenKey = `${userId}:${STORAGE_KEYS.TOKEN}`;

  console.log('Keys that will be used:');
  console.log('  sourceKey:', sourceKey);
  console.log('  tokenKey:', tokenKey);
  console.log('');

  // Step 1: Store the token (simulating loginWithToken)
  console.log('Step 1: Storing token...');
  await StorageService.setSecureSetting(tokenKey, token, userId);
  console.log('  Token stored');

  await StorageService.setSecureSetting(sourceKey, 'external', userId);
  console.log('  Source set to "external"');
  console.log('');

  // Step 2: Retrieve the token source
  console.log('Step 2: Retrieving token source...');
  const retrievedSource = await StorageService.getSecureSetting(sourceKey, userId);
  console.log('  Retrieved source:', retrievedSource);
  console.log('  Is external?', retrievedSource === 'external');
  console.log('');

  // Step 3: Retrieve the token
  console.log('Step 3: Retrieving token...');
  const retrievedToken = await StorageService.getSecureSetting(tokenKey, userId);
  if (retrievedToken) {
    console.log('  Token retrieved! Length:', retrievedToken.length);
    console.log('  Matches original?', retrievedToken === token);
  } else {
    console.log('  ERROR: Token NOT FOUND!');
  }
  console.log('');

  // Step 4: List all entries in database
  console.log('Step 4: Listing all secure_settings entries...');
  const db = StorageService.getDatabase();
  if (db) {
    const rows = await new Promise((resolve, reject) => {
      db.all('SELECT key, user_id, LENGTH(encrypted_value) as len FROM secure_settings', [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });

    if (rows.length === 0) {
      console.log('  No entries in database!');
    } else {
      console.log('  Found', rows.length, 'entries:');
      rows.forEach(row => {
        console.log(`    key="${row.key}" user_id="${row.user_id}" value_len=${row.len}`);
      });
    }
  }
  console.log('');

  // Step 5: Test with getActiveExternalToken
  console.log('Step 5: Testing getActiveExternalToken...');
  try {
    const ExternalTokenController = require('../src/api/controllers/external-token-controller.cjs');
    const result = await ExternalTokenController.getActiveExternalToken(userId);
    if (result) {
      console.log('  SUCCESS! Token retrieved via getActiveExternalToken');
      console.log('  Token length:', result.length);
    } else {
      console.log('  FAILED: getActiveExternalToken returned null');
    }
  } catch (err) {
    console.log('  ERROR:', err.message);
  }

  console.log('\n=== Test Complete ===\n');
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
