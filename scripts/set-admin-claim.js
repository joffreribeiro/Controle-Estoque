#!/usr/bin/env node
// Usage: node scripts/set-admin-claim.js /path/to/serviceAccountKey.json user@example.com
// Requires: npm install firebase-admin

const admin = require('firebase-admin');
const path = require('path');

async function main() {
  const [,, svcPath, email] = process.argv;
  if (!svcPath || !email) {
    console.error('Usage: node scripts/set-admin-claim.js /path/to/serviceAccountKey.json user@example.com');
    process.exit(1);
  }

  const fullPath = path.resolve(svcPath);
  let serviceAccount;
  try {
    serviceAccount = require(fullPath);
  } catch (err) {
    console.error('Failed to load service account JSON:', err.message || err);
    process.exit(1);
  }

  try {
    admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
  } catch (e) {
    // ignore if already initialized
  }

  try {
    const user = await admin.auth().getUserByEmail(email);
    await admin.auth().setCustomUserClaims(user.uid, { admin: true });
    console.log(`Success: set custom claim { admin: true } for ${email}`);
    // Optionally print instructions
    console.log('Note: The user may need to sign out and sign in again to refresh tokens.');
  } catch (err) {
    console.error('Error setting admin claim:', err.message || err);
    process.exit(1);
  }
}

main();
