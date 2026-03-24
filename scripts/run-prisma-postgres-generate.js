const path = require('path');
const { runPrismaPostgresGenerate } = require('./prisma-postgres-utils');

const root = path.resolve(__dirname, '..');
const target = process.argv[2] === 'prod' ? 'prod' : 'dev';

runPrismaPostgresGenerate(target, { cwd: root });
