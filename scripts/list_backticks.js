const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, '..', 'app.js');
const txt = fs.readFileSync(p, 'utf8');
let line = 1, col = 0;
let escape = false;
for (let i = 0; i < txt.length; i++) {
    const ch = txt[i];
    const next = txt[i+1];
    col++;
    if (ch === '\n') { line++; col = 0; }
    if (ch === '\\') { escape = !escape; continue; }
    if (ch === '`' && !escape) {
        // print context
        const start = Math.max(0, i-40);
        const end = Math.min(txt.length, i+40);
        const ctx = txt.slice(start,end).replace(/\n/g,'\\n');
        console.log('Backtick at i=' + i + ' line=' + line + ' col=' + col + ' ctx=' + ctx);
    }
    if (ch !== '\\') escape = false;
}
