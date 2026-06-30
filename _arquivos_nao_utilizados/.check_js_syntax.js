const fs = require('fs');
const vm = require('vm');
try {
  const code = fs.readFileSync('app.js', 'utf8');
  new vm.Script(code);
  console.log('OK: compilado com sucesso');
} catch (e) {
  console.error('SYNTAX_ERROR:', e && e.toString());
  if (e && e.stack) console.error(e.stack);
  process.exit(2);
}
