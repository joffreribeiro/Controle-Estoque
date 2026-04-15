const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, '..', 'app.js');
const s = fs.readFileSync(p, 'utf8');
const counts = {
  backticks: (s.match(/`/g) || []).length,
  singleQuotes: (s.match(/'/g) || []).length,
  doubleQuotes: (s.match(/"/g) || []).length,
  opensBrace: (s.match(/\{/g) || []).length,
  closesBrace: (s.match(/\}/g) || []).length,
};
console.log(JSON.stringify(counts, null, 2));
console.log('\n=== last 400 chars ===\n');
console.log(s.slice(-400));
