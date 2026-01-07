#!/usr/bin/env node
/**
 * Debug script to trace the token flow
 * Tests what getActiveExternalToken and getMostRecentToken return
 */

const path = require('path');

// Set up module paths
process.chdir(path.join(__dirname, '..'));

// Mock the monitoring service to avoid initialization issues
const MonitoringService = {
  debug: (msg, data) => console.log(`[DEBUG] ${msg}`, JSON.stringify(data, null, 2)),
  info: (msg, data) => console.log(`[INFO] ${msg}`, JSON.stringify(data, null, 2)),
  warn: (msg, data) => console.log(`[WARN] ${msg}`, JSON.stringify(data, null, 2)),
  error: (msg, data) => console.log(`[ERROR] ${msg}`, JSON.stringify(data, null, 2)),
  logError: (err) => console.log(`[ERROR] ${err.message || err}`)
};

// Override module resolution
require.cache[require.resolve('../src/core/monitoring-service.cjs')] = {
  exports: MonitoringService
};

const StorageService = require('../src/core/storage-service.cjs');
const ExternalTokenController = require('../src/api/controllers/external-token-controller.cjs');

async function main() {
  // Get userId from command line or use default test value
  const userId = process.argv[2] || 'ms365:test@example.com';

  console.log('\n=== Token Flow Debug ===\n');
  console.log(`Testing with userId: ${userId}\n`);

  // Initialize storage
  await StorageService.init();

  // 1. Check token source
  const sourceKey = `${userId}:token-source`;
  console.log(`1. Checking token source key: ${sourceKey}`);
  try {
    const source = await StorageService.getSecureSetting(sourceKey, userId);
    console.log(`   Result: ${source || 'NOT FOUND'}\n`);
  } catch (err) {
    console.log(`   Error: ${err.message}\n`);
  }

  // 2. Check external token
  const tokenKey = `${userId}:external-graph-token`;
  console.log(`2. Checking external token key: ${tokenKey}`);
  try {
    const token = await StorageService.getSecureSetting(tokenKey, userId);
    if (token) {
      console.log(`   Result: Token found (${token.substring(0, 50)}...)\n`);
    } else {
      console.log(`   Result: NOT FOUND\n`);
    }
  } catch (err) {
    console.log(`   Error: ${err.message}\n`);
  }

  // 3. Check OAuth token
  const oauthKey = `${userId}:ms-access-token`;
  console.log(`3. Checking OAuth token key: ${oauthKey}`);
  try {
    const token = await StorageService.getSecureSetting(oauthKey, userId);
    if (token) {
      console.log(`   Result: Token found (${token.substring(0, 50)}...)\n`);
    } else {
      console.log(`   Result: NOT FOUND\n`);
    }
  } catch (err) {
    console.log(`   Error: ${err.message}\n`);
  }

  // 4. Try getActiveExternalToken
  console.log(`4. Calling getActiveExternalToken("${userId}")`);
  try {
    const extToken = await ExternalTokenController.getActiveExternalToken(userId);
    if (extToken) {
      console.log(`   Result: Token found (${extToken.substring(0, 50)}...)\n`);
    } else {
      console.log(`   Result: null (no external token active)\n`);
    }
  } catch (err) {
    console.log(`   Error: ${err.message}\n`);
  }

  // 5. List all secure settings for this user
  console.log(`5. Listing all secure_settings in database:`);
  try {
    const db = StorageService.getDatabase();
    if (db) {
      const rows = await new Promise((resolve, reject) => {
        db.all('SELECT key, user_id FROM secure_settings', [], (err, rows) => {
          if (err) reject(err);
          else resolve(rows);
        });
      });

      if (rows.length === 0) {
        console.log('   No entries found\n');
      } else {
        rows.forEach(row => {
          console.log(`   - key: ${row.key}, user_id: ${row.user_id}`);
        });
        console.log();
      }
    } else {
      console.log('   Database not available\n');
    }
  } catch (err) {
    console.log(`   Error: ${err.message}\n`);
  }

  console.log('=== Debug Complete ===\n');
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
