const fs = require('fs');
const code = fs.readFileSync('app.js', 'utf8');
let depth = 0, i = 0, len = code.length;
let opens = [];

while (i < len) {
  const c = code[i];

  // skip strings and template literals
  if (c === '"' || c === "'" || c === '`') {
    const q = c;
    i++;
    while (i < len) {
      if (code[i] === '\\') { i += 2; continue; }
      if (q === '`' && code[i] === '$' && code[i + 1] === '{') {
        i += 2;
        let td = 1;
        while (i < len && td > 0) {
          if (code[i] === '\\') { i += 2; continue; }
          if (code[i] === '{') td++;
          if (code[i] === '}') td--;
          i++;
        }
        continue;
      }
      if (code[i] === q) { i++; break; }
      i++;
    }
    continue;
  }

  // skip line comments
  if (c === '/' && code[i + 1] === '/') {
    while (i < len && code[i] !== '\n') i++;
    continue;
  }

  // skip block comments
  if (c === '/' && code[i + 1] === '*') {
    i += 2;
    while (i < len && !(code[i] === '*' && code[i + 1] === '/')) i++;
    i += 2;
    continue;
  }

  // skip regex literals (basic heuristic)
  if (c === '/' && i > 0) {
    const prev = code.substring(Math.max(0, i - 20), i).trimEnd();
    const lastChar = prev[prev.length - 1];
    if ('=(:,;[!&|?+{'.includes(lastChar)) {
      i++;
      while (i < len) {
        if (code[i] === '\\') { i += 2; continue; }
        if (code[i] === '/') { i++; while (i < len && /[gimsuy]/.test(code[i])) i++; break; }
        if (code[i] === '\n') break;
        i++;
      }
      continue;
    }
  }

  if (c === '{') {
    depth++;
    if (depth <= 2) {
      const lineNum = code.substring(0, i).split('\n').length;
      const lineText = code.split('\n')[lineNum - 1].trim().substring(0, 120);
      opens.push({ d: depth, ln: lineNum, line: lineText });
    }
  }

  if (c === '}') {
    if (depth <= 2) {
      const lineNum = code.substring(0, i).split('\n').length;
      if (opens.length > 0 && opens[opens.length - 1].d === depth) {
        opens.pop();
      } else {
        console.log(`UNMATCHED CLOSE at line ${lineNum}, depth ${depth}`);
      }
    }
    depth--;
  }

  i++;
}

console.log('Final depth:', depth);
console.log('Unmatched opens:', opens.length);
opens.forEach(o => console.log(`  OPEN d=${o.d} line=${o.ln}: ${o.line}`));
