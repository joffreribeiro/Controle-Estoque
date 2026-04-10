// Deep brace analysis: tracks nesting depth through the file, 
// properly skipping strings, template literals, comments, and regex.
const fs = require('fs');
const src = fs.readFileSync(require('path').join(__dirname, '..', 'app.js'), 'utf8');
const lines = src.split('\n');

let depth = 0;
let i = 0;
const len = src.length;

// Track line number
let lineNum = 1;
let lineStart = 0;

// For each line, record depth at start
const depthAtLineStart = [0]; // index 0 unused, line 1 = index 1

function col() { return i - lineStart; }

const depthDropsToZero = []; // lines where depth returns to 0

while (i < len) {
    const ch = src[i];
    
    // Track line numbers
    if (ch === '\n') {
        lineNum++;
        lineStart = i + 1;
        depthAtLineStart[lineNum] = depth;
        i++;
        continue;
    }
    
    // Skip single-line comments
    if (ch === '/' && src[i+1] === '/') {
        while (i < len && src[i] !== '\n') i++;
        continue;
    }
    
    // Skip multi-line comments
    if (ch === '/' && src[i+1] === '*') {
        i += 2;
        while (i < len && !(src[i] === '*' && src[i+1] === '/')) {
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; depthAtLineStart[lineNum] = depth; }
            i++;
        }
        i += 2; // skip */
        continue;
    }
    
    // Skip template literals (backtick strings) - track nested ${} 
    if (ch === '`') {
        i++;
        let tmplDepth = 0;
        while (i < len) {
            if (src[i] === '\\') { i += 2; continue; }
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; depthAtLineStart[lineNum] = depth; }
            if (src[i] === '`' && tmplDepth === 0) { i++; break; }
            if (src[i] === '$' && src[i+1] === '{') { tmplDepth++; i += 2; continue; }
            if (src[i] === '}' && tmplDepth > 0) { tmplDepth--; i++; continue; }
            if (src[i] === '{' && tmplDepth > 0) { tmplDepth++; i++; continue; }
            // Nested template literal inside ${}
            if (src[i] === '`' && tmplDepth > 0) {
                // recursively skip nested template
                i++;
                let nestedTmpl = 0;
                while (i < len) {
                    if (src[i] === '\\') { i += 2; continue; }
                    if (src[i] === '\n') { lineNum++; lineStart = i + 1; depthAtLineStart[lineNum] = depth; }
                    if (src[i] === '`' && nestedTmpl === 0) { i++; break; }
                    if (src[i] === '$' && src[i+1] === '{') { nestedTmpl++; i += 2; continue; }
                    if (src[i] === '}' && nestedTmpl > 0) { nestedTmpl--; i++; continue; }
                    if (src[i] === '{' && nestedTmpl > 0) { nestedTmpl++; i++; continue; }
                    i++;
                }
                continue;
            }
            i++;
        }
        continue;
    }
    
    // Skip regular strings
    if (ch === '"' || ch === "'") {
        const quote = ch;
        i++;
        while (i < len && src[i] !== quote) {
            if (src[i] === '\\') i++;
            if (src[i] === '\n') { lineNum++; lineStart = i + 1; depthAtLineStart[lineNum] = depth; }
            i++;
        }
        i++; // skip closing quote
        continue;
    }
    
    // Track braces
    if (ch === '{') {
        depth++;
        i++;
        continue;
    }
    if (ch === '}') {
        depth--;
        if (depth === 0) {
            depthDropsToZero.push(lineNum);
        }
        if (depth < 0) {
            console.log(`*** DEPTH WENT NEGATIVE at line ${lineNum} ***`);
        }
        i++;
        continue;
    }
    
    i++;
}

console.log(`Final depth: ${depth}`);
console.log(`Total lines: ${lineNum}`);
console.log('');

if (depth !== 0) {
    console.log('FILE IS NOT BALANCED!');
} else {
    console.log('File braces are balanced (ignoring strings/comments/templates).');
}

console.log('');

// Now find top-level functions and check if depth returns to 0 properly
// Show each function start and the depth at that line
const funcPattern = /^(async\s+)?function\s+(\w+)/;
console.log('=== Top-level function analysis (depth at function start) ===');
let problemCount = 0;
for (let ln = 1; ln <= lines.length; ln++) {
    const line = lines[ln - 1];
    const trimmed = line.trim();
    const indent = line.length - line.trimStart().length;
    const m = trimmed.match(funcPattern);
    if (m && indent === 0) {
        const d = depthAtLineStart[ln] || 0;
        if (d !== 0) {
            console.log(`*** DEPTH ${d} *** at line ${ln}: ${m[0]} -- should be 0 for top-level function!`);
            problemCount++;
        }
    }
}

if (problemCount === 0) {
    console.log('All top-level functions start at depth 0. No structural issues found.');
} else {
    console.log(`\nFound ${problemCount} functions starting at wrong depth.`);
    
    // Find where depth should have been 0 but wasn't
    // Show the last place depth was 0 before each problem
    console.log('\n=== Detailed depth tracking around problems ===');
    for (let ln = 1; ln <= lines.length; ln++) {
        const line = lines[ln - 1];
        const trimmed = line.trim();
        const indent = line.length - line.trimStart().length;
        const m = trimmed.match(funcPattern);
        if (m && indent === 0) {
            const d = depthAtLineStart[ln] || 0;
            if (d !== 0) {
                // Find the last line where depth was 0 before this
                let lastZero = 0;
                for (let j = ln - 1; j >= 1; j--) {
                    if ((depthAtLineStart[j] || 0) === 0) {
                        lastZero = j;
                        break;
                    }
                }
                console.log(`\nFunction "${m[0]}" at line ${ln} has depth ${d}`);
                console.log(`Last depth=0 was at line ${lastZero}: ${lines[lastZero - 1]?.substring(0, 80)}`);
                // Show depth transitions near lastZero
                for (let j = Math.max(1, lastZero - 2); j <= Math.min(lines.length, lastZero + 10); j++) {
                    const dd = depthAtLineStart[j] || 0;
                    const marker = dd === 0 ? '' : ` [depth=${dd}]`;
                    console.log(`  ${j}: ${lines[j-1]?.substring(0, 80)}${marker}`);
                }
            }
        }
    }
}
