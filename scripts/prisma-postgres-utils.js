const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');

function runPrismaPostgresGenerate(target = 'dev', { cwd } = {}) {
  const root = cwd || path.resolve(__dirname, '..');
  const prismaCli = path.join(root, 'node_modules', 'prisma', 'build', 'index.js');

  if (!fs.existsSync(prismaCli)) {
    throw new Error('Prisma CLI not found. Run npm install in client first.');
  }

  const env = {
    ...process.env,
    PRISMA_ENV_TARGET: target === 'prod' ? 'prod' : 'dev',
  };

  childProcess.execFileSync(
    process.execPath,
    [prismaCli, 'generate', '--config', 'prisma.postgres.config.ts'],
    {
      cwd: root,
      env,
      stdio: 'inherit',
    },
  );
}

module.exports = {
  runPrismaPostgresGenerate,
};
