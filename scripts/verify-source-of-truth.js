const fs = require('fs');
const path = require('path');

const workspaceRoot = path.resolve(__dirname, '..');
const normalizeText = value => value.replace(/\r\n/g, '\n');
const sharedFetchPaths = [
  path.join(workspaceRoot, 'Google_Apps_Scripts', 'bkk-league-data', 'fetchFixtures.js'),
  path.join(workspaceRoot, 'Google_Apps_Scripts', 'team-sheet', 'fetchFixtures.js')
];

const PROJECTS = {
  'bkk-league-data': {
    folder: path.join(workspaceRoot, 'Google_Apps_Scripts', 'bkk-league-data'),
    required: ['appsscript.json', 'projectConfig.js', 'fetchFixtures.js'],
    forbidden: [/^build/, 'refreshPrediction.js']
  },
  'team-sheet': {
    folder: path.join(workspaceRoot, 'Google_Apps_Scripts', 'team-sheet'),
    required: ['appsscript.json', 'projectConfig.js', 'fetchFixtures.js', 'onEdit_Trigger.js', 'refreshPrediction.js'],
    forbidden: []
  }
};

const problems = [];

for (const [name, config] of Object.entries(PROJECTS)) {
  if (!fs.existsSync(config.folder)) {
    problems.push(`[${name}] Folder not found: ${config.folder}`);
    continue;
  }

  const entries = fs.readdirSync(config.folder, { withFileTypes: true });
  const files = entries.filter(e => e.isFile()).map(e => e.name);

  // Check required files
  for (const req of config.required) {
    if (!files.includes(req)) {
      problems.push(`[${name}] Missing required file: ${req}`);
    }
  }

  // Check forbidden files
  for (const rule of config.forbidden) {
    const matched = files.filter(f =>
      rule instanceof RegExp ? rule.test(f) : f === rule
    );
    matched.forEach(f => problems.push(`[${name}] Forbidden file present: ${f}`));
  }

  // Check for .gs files (should never be present)
  const gsFiles = files.filter(f => f.endsWith('.gs'));
  if (gsFiles.length > 0) {
    problems.push(`[${name}] Legacy .gs files found: ${gsFiles.join(', ')}`);
  }

  // Check for copy-named files
  const copyFiles = files.filter(f => /^Copy( \d+)? of /i.test(f));
  if (copyFiles.length > 0) {
    problems.push(`[${name}] Copy-named files found: ${copyFiles.join(', ')}`);
  }

  const jsFiles = files.filter(f => f.endsWith('.js')).sort();
  console.log(`[${name}] Deployable files (${jsFiles.length}): ${jsFiles.join(', ')}`);
}

const existingSharedFetchPaths = sharedFetchPaths.filter(fs.existsSync);
if (existingSharedFetchPaths.length === sharedFetchPaths.length) {
  const [canonicalFetchPath, ...otherFetchPaths] = sharedFetchPaths;
  const canonicalFetch = normalizeText(fs.readFileSync(canonicalFetchPath, 'utf8'));

  otherFetchPaths.forEach(fetchPath => {
    const currentFetch = normalizeText(fs.readFileSync(fetchPath, 'utf8'));
    if (currentFetch !== canonicalFetch) {
      problems.push(
        `[shared] fetchFixtures.js mismatch: ${path.relative(workspaceRoot, fetchPath)} must match ${path.relative(workspaceRoot, canonicalFetchPath)}`
      );
    }
  });
}

if (problems.length > 0) {
  console.error('\nSource-of-truth check FAILED:');
  problems.forEach(p => console.error(`  - ${p}`));
  process.exit(1);
}

console.log('\nSource-of-truth check passed.');
