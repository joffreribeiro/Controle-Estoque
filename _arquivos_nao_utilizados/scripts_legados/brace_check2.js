// Find the second missing } in app.js by tracking brace depth with proper JS parsing
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');

// Tokenizer approach: extract only structural braces (not in strings/comments/templates)
let i = 0;
const len = src.length;
let depth = 0;
let lineNum = 1;

// Stack for template literal nesting
// Each entry is 'template' meaning we're inside a ${...} inside a template
const tmplStack = [];

const events = []; // {line, type: 'open'|'close', depth_before, depth_after}

function skipBlockComment() {
    i += 2; // skip /*
    while (i < len - 1) {
        if (src[i] === '\n') lineNum++;
        if (src[i] === '*' && src[i+1] === '/') { i += 2; return; }
        i++;
    }
}

function skipLineComment() {
    i += 2; // skip //
    while (i < len && src[i] !== '\n') i++;
}

function skipString(quote) {
    i++; // skip opening quote
    while (i < len) {
        if (src[i] === '\\') { i += 2; continue; }
        if (src[i] === '\n') lineNum++;
        if (src[i] === quote) { i++; return; }
        i++;
    }
}

function parseTemplateContent() {
    // We're inside a template literal (after ` or after } closing a ${} expression)
    while (i < len) {
        if (src[i] === '\\') { i += 2; continue; }
        if (src[i] === '\n') { lineNum++; i++; continue; }
        if (src[i] === '`') { i++; return; } // end of template
        if (src[i] === '$' && i + 1 < len && src[i+1] === '{') {
            // Enter template expression
            tmplStack.push('template');
            i += 2; // skip ${
            return; // return to main parse loop to parse the expression
        }
        i++;
    }
}

while (i < len) {
    const ch = src[i];
    
    if (ch === '\n') { lineNum++; i++; continue; }
    
    // Block comment
    if (ch === '/' && i + 1 < len && src[i+1] === '*') { skipBlockComment(); continue; }
    
    // Line comment  
    if (ch === '/' && i + 1 < len && src[i+1] === '/') { skipLineComment(); continue; }
    
    // Strings
    if (ch === '"' || ch === "'") { skipString(ch); continue; }
    
    // Template literal start
    if (ch === '`') { i++; parseTemplateContent(); continue; }
    
    // Opening brace
    if (ch === '{') {
        depth++;
        i++;
        continue;
    }
    
    // Closing brace
    if (ch === '}') {
        // Check if this closes a template expression (${...})
        if (tmplStack.length > 0 && tmplStack[tmplStack.length - 1] === 'template') {
            // This } closes the ${} expression - don't change depth since ${ didn't either
            tmplStack.pop();
            // Continue parsing template content
            parseTemplateContent();
        } else {
            depth--;
        }
        i++;
        continue;
    }
    
    i++;
}

console.log('Final depth (with proper template handling):', depth);
console.log('');

// Now do a second pass: track depth at each function declaration  
i = 0;
depth = 0;
lineNum = 1;
tmplStack.length = 0;

const funcDepths = [];
let currentLineStart = 0;

function getLineText() {
    let end = src.indexOf('\n', currentLineStart);
    if (end === -1) end = len;
    return src.substring(currentLineStart, end);
}

// Reset and reparse, this time tracking line-level info
i = 0;
depth = 0;
lineNum = 1;
currentLineStart = 0;

let depthAtLineStart = 0;

while (i < len) {
    const ch = src[i];
    
    if (ch === '\n') {
        lineNum++;
        i++;
        currentLineStart = i;
        depthAtLineStart = depth;
        
        // Check if NEW line starts with function keyword
        let lineEnd = src.indexOf('\n', currentLineStart);
        if (lineEnd === -1) lineEnd = len;
        const line = src.substring(currentLineStart, lineEnd).trim();
        if (/^(async\s+)?function\s+\w+/.test(line)) {
            funcDepths.push({ line: lineNum, depth: depth, text: line.substring(0, 90) });
        }
        continue;
    }
    
    if (ch === '/' && i + 1 < len && src[i+1] === '*') { skipBlockComment(); continue; }
    if (ch === '/' && i + 1 < len && src[i+1] === '/') { skipLineComment(); continue; }
    if (ch === '"' || ch === "'") { skipString(ch); continue; }
    if (ch === '`') { i++; parseTemplateContent(); continue; }
    
    if (ch === '{') { depth++; i++; continue; }
    if (ch === '}') {
        if (tmplStack.length > 0 && tmplStack[tmplStack.length - 1] === 'template') {
            tmplStack.pop();
            parseTemplateContent();
        } else {
            depth--;
        }
        i++;
        continue;
    }
    
    i++;
}

console.log('Second pass final depth:', depth);
console.log('');

// Find where depth changes between consecutive functions
let prevDepth = -1;
for (const f of funcDepths) {
    if (prevDepth !== -1 && f.depth !== prevDepth) {
        console.log(`*** DEPTH CHANGE ${prevDepth} -> ${f.depth} at line ${f.line}: ${f.text}`);
    }
    prevDepth = f.depth;
}

console.log('\nAll function depths:');
for (const f of funcDepths) {
    console.log(`  Line ${f.line}: depth=${f.depth}  ${f.text}`);
}
