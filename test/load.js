'use strict';
/** Loads every src/*.js file into one VM context with the GAS mock, exactly like Apps Script's global scope. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { createGas } = require('./gas-mock');

const SRC = path.join(__dirname, '..', 'src');

function sourceFiles() {
  return fs.readdirSync(SRC).filter((f) => f.endsWith('.js')).sort();
}

function bundleSource() {
  return sourceFiles().map((f) => `// ===== ${f} =====\n` + fs.readFileSync(path.join(SRC, f), 'utf8')).join('\n\n');
}

function loadApp(options = {}) {
  const gas = createGas(options);
  // Captured console: expected error logs (e.g. deliberate failure tests) don't clutter the output.
  gas.state.console = [];
  const quietConsole = {
    log: (...a) => gas.state.console.push(['log', a.join(' ')]),
    warn: (...a) => gas.state.console.push(['warn', a.join(' ')]),
    error: (...a) => gas.state.console.push(['error', a.join(' ')]),
  };
  const context = vm.createContext(Object.assign({ console: options.verbose ? console : quietConsole }, gas.globals));
  vm.runInContext(bundleSource(), context, { filename: 'apps-script-bundle.js' });
  return { app: context, gas };
}

module.exports = { loadApp, bundleSource, sourceFiles };
