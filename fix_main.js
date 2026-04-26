const fs = require('fs');
const p = 'apps/word-addin/src/main.ts';
let lines = fs.readFileSync(p, 'utf8').split('\n');

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];
  // Fix 1: early return for normal format
  if (line.includes('if (format === "normal" || !format) return;')) {
    lines[i] = line.replace('if (format === "normal" || !format) return;', 'if (!format) return;');
    console.log('Fixed line', i + 1, ':', lines[i]);
  }
  // Fix 2: add normal mapping
  if (line.includes('heading1: ["Heading 1", "标题 1", "标题1"]')) {
    lines[i] = '    normal: ["Normal", "正文"],' + '\r' + '\n' + line;
    console.log('Fixed line', i + 1, ': inserted normal mapping');
  }
}

fs.writeFileSync(p, lines.join('\n'), 'utf8');
console.log('Done');
