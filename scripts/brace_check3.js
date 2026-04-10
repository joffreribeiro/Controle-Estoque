// Approach: find where consecutive top-level functions are separated improperly.
// For each function start, scan backwards to find its expected closing } or absence thereof.
const fs = require('fs');
const path = require('path');
const lines = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8').split('\n');

// Find all function declaration lines (at column 0 or very low indent)
const funcLines = [];
for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();
    // Function at start of line (no indent or very little)
    const indent = line.length - line.trimStart().length;
    if (/^(async\s+)?function\s+\w+/.test(trimmed) && indent <= 0) {
        funcLines.push({ lineNum: i + 1, text: trimmed.substring(0, 80), indent });
    }
}

console.log(`Found ${funcLines.length} function declarations at column 0\n`);

// For each pair of consecutive functions, check if there's a closing } between them
for (let fi = 0; fi < funcLines.length - 1; fi++) {
    const curr = funcLines[fi];
    const next = funcLines[fi + 1];
    
    // Look for a line that is just "}" between curr and next
    let foundClose = false;
    let lastCloseLine = -1;
    
    // Scan backwards from next function to curr function
    for (let j = next.lineNum - 2; j >= curr.lineNum; j--) {
        const line = lines[j].trim();
        if (line === '}') {
            foundClose = true;
            lastCloseLine = j + 1;
            break;
        }
        // Also count }; or }) etc but primary is just }
        if (line === '};' || line === '});') {
            foundClose = true;
            lastCloseLine = j + 1;
            break;
        }
        // Skip blanks and comments
        if (line === '' || line.startsWith('//') || line.startsWith('/*') || line.startsWith('*')) continue;
        // If we hit actual code, the function might not be closed yet
        break;
    }
    
    if (!foundClose) {
        // Check what's between them - just look at the lines between
        const gapStart = curr.lineNum;
        const gapEnd = next.lineNum - 1;
        
        // Look for ANY } at indent 0 in the entire range
        let anyCloseAtIndent0 = false;
        for (let j = gapEnd - 1; j >= curr.lineNum; j--) {
            if (lines[j].trim() === '}' && (lines[j].length - lines[j].trimStart().length) === 0) {
                anyCloseAtIndent0 = true;
                lastCloseLine = j + 1;
                break;
            }
        }
        
        if (!anyCloseAtIndent0) {
            console.log(`*** POSSIBLE MISSING } ***`);
            console.log(`  Function: ${curr.text} (line ${curr.lineNum})`);
            console.log(`  Next function: ${next.text} (line ${next.lineNum})`);
            console.log(`  No "}" at indent 0 found between lines ${curr.lineNum} and ${next.lineNum}`);
            
            // Show what's just before the next function
            const previewStart = Math.max(next.lineNum - 5, curr.lineNum);
            console.log(`  Lines before next function:`);
            for (let j = previewStart - 1; j < next.lineNum - 1; j++) {
                console.log(`    ${j + 1}: ${lines[j]}`);
            }
            console.log('');
        }
    }
}
