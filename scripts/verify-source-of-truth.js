const fs = require('fs');
const path = require('path');

const workspaceRoot = path.resolve(__dirname, '..');
const rootEntries = fs.readdirSync(workspaceRoot, {withFileTypes: true});
const rootFiles = rootEntries.filter(entry => entry.isFile()).map(entry => entry.name);
const problems = [];

if (!rootFiles.includes('appsscript.json')) {
  problems.push('Missing root appsscript.json manifest.');
}

if (!rootFiles.includes('projectConfig.js')) {
  problems.push('Missing root projectConfig.js spreadsheet binding file.');
}

const rootGsFiles = rootFiles.filter(name => name.endsWith('.gs'));
if (rootGsFiles.length > 0) {
  problems.push(`Root .gs files found: ${rootGsFiles.join(', ')}`);
}

const copyNamedFiles = rootFiles.filter(name => /^Copy( \d+)? of /i.test(name));
if (copyNamedFiles.length > 0) {
  problems.push(`Copy-named root files found: ${copyNamedFiles.join(', ')}`);
}

const gsJsFiles = rootFiles.filter(name => name.endsWith('.gs.js'));
if (gsJsFiles.length > 0) {
  problems.push(`Legacy .gs.js root files found: ${gsJsFiles.join(', ')}`);
}

const deployableJsFiles = rootFiles.filter(name => name.endsWith('.js')).sort();
if (deployableJsFiles.length === 0) {
  problems.push('No root .js Apps Script files found.');
}

if (problems.length > 0) {
  console.error('Single source-of-truth check failed.');
  problems.forEach(problem => console.error(`- ${problem}`));
  process.exit(1);
}

console.log('Single source-of-truth check passed.');
console.log('Canonical deployable files:');
deployableJsFiles.forEach(name => console.log(`- ${name}`));