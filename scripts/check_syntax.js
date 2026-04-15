const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, '..', 'app.js');
const txt = fs.readFileSync(p, 'utf8');

let line = 1, col = 0;
let inSingle = false, inDouble = false, inBacktick = false, inLineComment = false, inBlockComment = false;
let escape = false;
let globalStack = [];
let templateStack = []; // stack of {line,col,nested}
let backtickCount = 0;
let dollarBraceCount = 0;
let issues = [];

for (let i = 0; i < txt.length; i++) {
    const ch = txt[i];
    const next = txt[i+1];
    col++;
    if (ch === '\n') { line++; col = 0; inLineComment = false; }

    if (inLineComment) { escape = false; continue; }
    if (inBlockComment) {
        if (ch === '*' && next === '/') { inBlockComment = false; i++; col++; }
        escape = false; continue;
    }

    if (!inSingle && !inDouble && !inBacktick) {
        if (ch === '/' && next === '/') { inLineComment = true; i++; col++; continue; }
        if (ch === '/' && next === '*') { inBlockComment = true; i++; col++; continue; }
    }

    if (escape) { escape = false; continue; }

    if (ch === '\\') { escape = true; continue; }

    if (!inDouble && !inBacktick && ch === "'") { inSingle = !inSingle; continue; }
    if (!inSingle && !inBacktick && ch === '"') { inDouble = !inDouble; continue; }

    if (!inSingle && !inDouble && ch === '`') {
        inBacktick = !inBacktick;
        backtickCount++;
        // when closing a backtick we should clear any dangling template stack
        if (!inBacktick && templateStack.length) {
            issues.push({type:'template_unclosed', message:'Template expression ${...} not closed before backtick end', line, col, stack: templateStack.slice()});
        }
        continue;
    }

    // detect ${ inside backtick
    if (inBacktick && ch === '$' && next === '{' && !inSingle && !inDouble) {
        templateStack.push({line, col, nested:0});
        dollarBraceCount++;
        i++; col++; // skip {
        continue;
    }

    // braces handling
    if (ch === '{') {
        if (templateStack.length > 0) {
            // inside template expression
            templateStack[templateStack.length-1].nested++;
        } else if (!inSingle && !inDouble && !inBacktick) {
            globalStack.push({line,col});
        }
        continue;
    }
    if (ch === '}') {
        if (templateStack.length > 0) {
            const top = templateStack[templateStack.length-1];
            if (top.nested > 0) top.nested--; else templateStack.pop();
        } else if (!inSingle && !inDouble && !inBacktick) {
            if (globalStack.length === 0) {
                issues.push({type:'unexpected_closing_brace', message:'Extra closing }', line, col});
            } else {
                globalStack.pop();
            }
        }
        continue;
    }
}

const report = {
    backtickCount,
    unmatchedBacktick: backtickCount % 2 !== 0,
    dollarBraceCount,
    templateStackDepth: templateStack.length,
    globalBraceDepth: globalStack.length,
    issues
};

console.log('=== SYNTAX CHECK REPORT ===');
console.log(JSON.stringify(report, null, 2));

if (templateStack.length) {
    console.log('\nUnclosed template stacks:');
    templateStack.forEach(t => console.log('  at line', t.line, 'col', t.col, 'nested', t.nested));
}
if (globalStack.length) {
    console.log('\nUnclosed global braces (first 10):');
    globalStack.slice(0,10).forEach(t => console.log('  at line', t.line, 'col', t.col));
}
if (issues.length) {
    console.log('\nIssues found:');
    issues.forEach((it, idx) => console.log(idx+1, it.type, it.message, 'at', it.line, it.col));
}

if (report.unmatchedBacktick || report.templateStackDepth || report.globalBraceDepth || issues.length) process.exitCode = 2; else process.exitCode = 0;
