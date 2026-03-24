const childProcess = require('child_process');
const fs = require('fs');
const path = require('path');
const { buildPrismaEnv } = require('./prisma-postgres-utils');

const root = path.resolve(__dirname, '..');
const prismaCli = path.join(root, 'node_modules', 'prisma', 'build', 'index.js');
const migrationsDir = path.join(root, 'prisma-postgres', 'migrations');

function normalizeSql(sql) {
  return String(sql || '')
    .replace(/\r\n/g, '\n')
    .split('\n')
    .filter((line) => !line.trim().startsWith('--'))
    .join('\n')
    .replace(/\s+/g, ' ')
    .replace(/\s*;\s*/g, ';')
    .trim();
}

function readCombinedMigrationsSql() {
  if (!fs.existsSync(migrationsDir)) {
    throw new Error(`Migrations directory not found: ${migrationsDir}`);
  }

  const entries = fs
    .readdirSync(migrationsDir, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort((a, b) => a.localeCompare(b));

  if (!entries.length) {
    throw new Error('No Prisma migrations found in prisma-postgres/migrations.');
  }

  return entries
    .map((entry) => {
      const filePath = path.join(migrationsDir, entry, 'migration.sql');
      if (!fs.existsSync(filePath)) {
        throw new Error(`Missing migration.sql in ${entry}`);
      }
      return fs.readFileSync(filePath, 'utf8');
    })
    .join('\n');
}

function generateSqlFromSchema() {
  if (!fs.existsSync(prismaCli)) {
    throw new Error('Prisma CLI not found. Run npm install in client first.');
  }

  return childProcess.execFileSync(
    process.execPath,
    [
      prismaCli,
      'migrate',
      'diff',
      '--from-empty',
      '--to-schema',
      'prisma-postgres/schema.prisma',
      '--script',
      '--config',
      'prisma.postgres.config.ts',
    ],
    {
      cwd: root,
      env: buildPrismaEnv('dev'),
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    },
  );
}

function main() {
  try {
    const fromMigrations = normalizeSql(readCombinedMigrationsSql());
    const fromSchema = normalizeSql(generateSqlFromSchema());

    if (fromMigrations !== fromSchema) {
      console.error(
        '[prisma-postgres-check] Migrations SQL does not match the current schema SQL.',
      );
      process.exit(1);
    }

    console.log(
      '[prisma-postgres-check] Migrations SQL matches the current Prisma schema.',
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[prisma-postgres-check] ${message}`);
    process.exit(1);
  }
}

main();
