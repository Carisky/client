const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');
const { runPrismaPostgresGenerate } = require('./prisma-postgres-utils');

const root = path.resolve(__dirname, '..');
const forgeCommand = process.argv[2] === 'package' ? 'package' : 'make';
const forgeCli = path.join(root, 'node_modules', '@electron-forge', 'cli', 'dist', 'electron-forge.js');

if (!fs.existsSync(forgeCli)) {
  throw new Error('Electron Forge not found. Run npm install in client first.');
}

runPrismaPostgresGenerate('prod', { cwd: root });

childProcess.execFileSync(process.execPath, [forgeCli, forgeCommand], {
  cwd: root,
  stdio: 'inherit',
});
