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
  length: s.length
};
fs.writeFileSync(path.join(__dirname,'count_chars_output.txt'), JSON.stringify(counts, null, 2) + '\n\nLast400:\n' + s.slice(-400));
console.log('Wrote output to scripts/count_chars_output.txt');
