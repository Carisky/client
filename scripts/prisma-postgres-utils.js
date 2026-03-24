const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');

function getPrismaCliPath(root) {
  return path.join(root, 'node_modules', 'prisma', 'build', 'index.js');
}

function ensurePrismaCli(root) {
  const prismaCli = getPrismaCliPath(root);
  if (!fs.existsSync(prismaCli)) {
    throw new Error('Prisma CLI not found. Run npm install in client first.');
  }
  return prismaCli;
}

function buildPrismaEnv(target = 'dev') {
  return {
    ...process.env,
    PRISMA_ENV_TARGET: target === 'prod' ? 'prod' : 'dev',
  };
}

function runPrismaPostgresCli(args, target = 'dev', { cwd } = {}) {
  const root = cwd || path.resolve(__dirname, '..');
  const prismaCli = ensurePrismaCli(root);
  const env = buildPrismaEnv(target);

  childProcess.execFileSync(
    process.execPath,
    [prismaCli, ...args, '--config', 'prisma.postgres.config.ts'],
    {
      cwd: root,
      env,
      stdio: 'inherit',
    },
  );
}

function runPrismaPostgresGenerate(target = 'dev', { cwd } = {}) {
  runPrismaPostgresCli(['generate'], target, { cwd });
}

module.exports = {
  buildPrismaEnv,
  runPrismaPostgresCli,
  runPrismaPostgresGenerate,
};
