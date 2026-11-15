'use strict';
/** Copies the paste-ready bundle from dist/ into site/downloads/ so the Pages site serves the current build. */
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const out = path.join(root, 'site', 'downloads');
fs.mkdirSync(out, { recursive: true });
['Code.gs', 'appsscript.json'].forEach((f) => fs.copyFileSync(path.join(root, 'dist', f), path.join(out, f)));
fs.writeFileSync(path.join(root, 'site', '_headers'),
  '/downloads/*\n  Content-Type: text/plain; charset=utf-8\n  Cache-Control: no-cache\n\n' +
  '/*\n  X-Content-Type-Options: nosniff\n  Referrer-Policy: strict-origin-when-cross-origin\n');
console.log('site/downloads/ updated from dist/');
