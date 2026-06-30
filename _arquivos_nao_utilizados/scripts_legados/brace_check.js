const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const lines = src.split('\n');

// Proper JS brace tracker with template literal support
let depth = 0;
let i = 0; // char index
const len = src.length;

// State stack for template literals: when we see ${ inside a template, push state
const stateStack = []; // 'template' means we're in a template literal waiting for }

let lineNum = 1;
let lineStart = 0;

// Track depth at start of each line
const lineDepths = new Array(lines.length + 1);
lineDepths[1] = 0;

const functionDepths = [];
let prevDepth = 0;
let depthChanges = []; // where depth increases unexpectedly

while (i < len) {
    const ch = src[i];
    
    // Track line numbers
    if (ch === '\n') {
        lineNum++;
        lineStart = i + 1;
        lineDepths[lineNum] = depth;
        
        // Check if line starts a function
        const lineText = lines[lineNum - 2] || ''; // previous line (0-indexed)
        const trimmed = lineText.trim();
        if (/^(async\s+)?function\s+\w+/.test(trimmed) || /^(async\s+)?function\s*\(/.test(trimmed)) {
            functionDepths.push({ lineNum: lineNum - 1, depth: lineDepths[lineNum - 1], text: trimmed.substring(0, 80) });
        }
        i++;
        continue;
    }
    
    // Block comment
    if (ch === '/' && i + 1 < len && src[i + 1] === '*') {
        i += 2;
        while (i < len && !(src[i] === '*' && i + 1 < len && src[i + 1] === '/')) {
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; lineDepths[lineNum] = depth; }
            i++;
        }
        i += 2; // skip */
        continue;
    }
    
    // Line comment
    if (ch === '/' && i + 1 < len && src[i + 1] === '/') {
        i += 2;
        while (i < len && src[i] !== '\n') i++;
        continue;
    }
    
    // Single or double quoted string
    if (ch === '"' || ch === "'") {
        const quote = ch;
        i++;
        while (i < len && src[i] !== quote) {
            if (src[i] === '\\') i++; // skip escaped char
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; lineDepths[lineNum] = depth; }
            i++;
        }
        i++; // skip closing quote
        continue;
    }
    
    // Template literal
    if (ch === '`') {
        i++;
        while (i < len) {
            if (src[i] === '\\') { i += 2; continue; }
            if (src[i] === '`') { i++; break; } // end of template
            if (src[i] === '$' && i + 1 < len && src[i + 1] === '{') {
                // Template expression - push state and count this as a real brace
                stateStack.push('template');
                depth++;
                i += 2;
                // Now parse normally until matching }
                break;
            }
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; lineDepths[lineNum] = depth; }
            i++;
        }
        continue;
    }
    
    if (ch === '{') {
        depth++;
        i++;
        continue;
    }
    
    if (ch === '}') {
        depth--;
        // Check if we're closing a template expression
        if (stateStack.length > 0 && stateStack[stateStack.length - 1] === 'template') {
            stateStack.pop();
            // We're back in template literal, continue parsing it
            i++;
            while (i < len) {
                if (src[i] === '\\') { i += 2; continue; }
                if (src[i] === '`') { i++; break; }
                if (src[i] === '$' && i + 1 < len && src[i + 1] === '{') {
                    stateStack.push('template');
                    depth++;
                    i += 2;
                    break;
                }
                if (src[i] === '\n') { lineNum++; lineStart = i + 1; lineDepths[lineNum] = depth; }
                i++;
            }
            continue;
        }
        i++;
        continue;
    }
    
    i++;
}

// Check last line for function
const lastTrimmed = (lines[lines.length - 1] || '').trim();
if (/^(async\s+)?function\s+\w+/.test(lastTrimmed)) {
    functionDepths.push({ lineNum: lines.length, depth, text: lastTrimmed.substring(0, 80) });
}

console.log('Final depth:', depth);

// Find where depth changes between consecutive top-level functions
console.log('\nAll top-level function declarations with depth (expect 0 for truly top-level):');
let prevFnDepth = 0;
for (const f of functionDepths) {
    if (f.depth !== prevFnDepth) {
        console.log(`  *** DEPTH CHANGE ${prevFnDepth} -> ${f.depth} *** Line ${f.lineNum}: ${f.text}`);
    }
    prevFnDepth = f.depth;
}

// Show depth at key lines
console.log('\nDepth at key transition points:');
for (let ln = 1; ln <= lines.length; ln++) {
    const t = lines[ln-1].trim();
    if (/^(async\s+)?function\s+\w+/.test(t) && lineDepths[ln] !== undefined) {
        if (lineDepths[ln] !== lineDepths[ln-1]) {
            // only show if different from surrounding context
        }
    }
}

// Show all depth transitions
console.log('\nLines where depth increases by opening a function but never returns:');
let lastDepthAtFunctionStart = {};
for (const f of functionDepths) {
    console.log(`  Line ${f.lineNum}: depth=${f.depth} ${f.text}`);
}
