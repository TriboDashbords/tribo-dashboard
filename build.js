// Збірка: src/server/*.js -> dist/server.js, src/web/index.html -> dist/index.html
// Запуск: node build.js   (потім node tests/run.js)
const fs = require('fs');
const ORDER = ['Config', 'Storage', 'Groups', 'Ingest', 'ExcelImport', 'Setup', 'TestData', 'Dashboard', 'PlanFact', 'GmailImport', 'Auth', 'Api'];
const parts = ORDER.map(n => `// ===================== ${n}.js =====================\n` + fs.readFileSync(`src/server/${n}.js`, 'utf8'));
const out = '// ЗІБРАНО АВТОМАТИЧНО з src/server (node build.js). Не редагуйте цей файл напряму.\n' + parts.join('\n\n');
fs.mkdirSync('dist', { recursive: true });
fs.writeFileSync('dist/server.js', out);
const stamp = new Date().toISOString().slice(0, 16).replace('T', ' ');
fs.writeFileSync('dist/index.html', fs.readFileSync('src/web/index.html', 'utf8').replace(/__BUILD__/g, stamp + ' UTC'));
console.log('dist/server.js', out.length, 'bytes; dist/index.html', fs.statSync('dist/index.html').size, 'bytes');
