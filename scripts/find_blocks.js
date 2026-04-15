// Proper JS parser that handles template literals with nested ${} expressions
const fs = require('fs');
const src = fs.readFileSync('app.js', 'utf8');
const len = src.length;

let depth = 0;
let pos = 0;
let lineNum = 1;

// Stack for tracking opens
const openStack = [];

// Template literal nesting: stack of depths when we entered ${}
const templateStack = [];

function lineAt(p) {
  let ln = 1;
  for (let i = 0; i < p && i < len; i++) {
    if (src[i] === '\n') ln++;
  }
  return ln;
}

function contextAt(ln) {
  const lines = src.split('\n');
  return (lines[ln - 1] || '').trim().substring(0, 120);
}

while (pos < len) {
  const ch = src[pos];
  
  // Track line numbers
  if (ch === '\n') { lineNum++; pos++; continue; }
  
  // Line comment
  if (ch === '/' && pos + 1 < len && src[pos + 1] === '/') {
    while (pos < len && src[pos] !== '\n') pos++;
    continue;
  }
  
  // Block comment
  if (ch === '/' && pos + 1 < len && src[pos + 1] === '*') {
    pos += 2;
    while (pos < len) {
      if (src[pos] === '\n') lineNum++;
      if (src[pos] === '*' && pos + 1 < len && src[pos + 1] === '/') { pos += 2; break; }
      pos++;
    }
    continue;
  }
  
  // Single/double quoted strings
  if (ch === "'" || ch === '"') {
    const q = ch;
    pos++;
    while (pos < len) {
      if (src[pos] === '\\') { pos += 2; continue; }
      if (src[pos] === '\n') lineNum++;
      if (src[pos] === q) { pos++; break; }
      pos++;
    }
    continue;
  }
  
  // Template literal
  if (ch === '`') {
    pos++;
    while (pos < len) {
      if (src[pos] === '\\') { pos += 2; continue; }
      if (src[pos] === '\n') lineNum++;
      if (src[pos] === '$' && pos + 1 < len && src[pos + 1] === '{') {
        // Enter template expression - push current state
        templateStack.push('template');
        pos += 2;
        depth++;
        openStack.push({ line: lineNum, type: 'template-expr' });
        break; // back to main loop to parse the expression
      }
      if (src[pos] === '`') { pos++; break; } // end of template
      pos++;
    }
    continue;
  }
  
  // Opening brace
  if (ch === '{') {
    const prevDepth = depth;
    depth++;
    openStack.push({ line: lineNum, depth: depth, context: contextAt(lineNum) });
    if (prevDepth === 0) {
      console.log(`OPEN { at line ${lineNum} (0->1): ${contextAt(lineNum)}`);
    }
    if (prevDepth === 1) {
      // only log first few
    }
    pos++;
    continue;
  }
  
  // Closing brace
  if (ch === '}') {
    const matchedOpen = openStack.pop();
    
    // Check if this closes a template expression
    if (templateStack.length > 0 && depth === openStack.length + 1) {
      // Check if the matched open was a template-expr
      if (matchedOpen && matchedOpen.type === 'template-expr') {
        templateStack.pop();
        depth--;
        pos++;
        // Continue parsing the template literal
        while (pos < len) {
          if (src[pos] === '\\') { pos += 2; continue; }
          if (src[pos] === '\n') lineNum++;
          if (src[pos] === '$' && pos + 1 < len && src[pos + 1] === '{') {
            templateStack.push('template');
            pos += 2;
            depth++;
            openStack.push({ line: lineNum, type: 'template-expr' });
            break;
          }
          if (src[pos] === '`') { pos++; break; }
          pos++;
        }
        continue;
      }
    }
    
    if (depth === 1) {
      console.log(`CLOSE } at line ${lineNum} (1->0): opened at line ${matchedOpen?.line}: ${matchedOpen?.context}`);
    }
    // Report last closes
    if (lineNum > 11380) {
      console.log(`LATE CLOSE } at line ${lineNum} (depth ${depth}->${depth-1}): opened at line ${matchedOpen?.line}: ${matchedOpen?.context}`);
    }
    depth--;
    pos++;
    continue;
  }
  
  pos++;
}

console.log('\nFinal depth:', depth);
console.log('Total lines:', lineNum);
