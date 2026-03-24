const path = require('path');
const { runPrismaPostgresCli } = require('./prisma-postgres-utils');

const root = path.resolve(__dirname, '..');

function printUsage() {
  console.log(`
Usage:
  node scripts/run-prisma-postgres-command.js [--target dev|prod] [--allow-prod] <prisma args...>

Examples:
  node scripts/run-prisma-postgres-command.js db push
  node scripts/run-prisma-postgres-command.js migrate status
  node scripts/run-prisma-postgres-command.js migrate dev --create-only --name init_auth
  node scripts/run-prisma-postgres-command.js --target prod --allow-prod migrate deploy
  `);
}

function parseArgs(argv) {
  const args = [...argv];
  let target = 'dev';
  let allowProd = false;
  const prismaArgs = [];

  while (args.length > 0) {
    const current = args.shift();
    if (current === '--help' || current === '-h') {
      return { help: true };
    }
    if (current === '--target') {
      const value = String(args.shift() || '').trim().toLowerCase();
      if (value !== 'dev' && value !== 'prod') {
        throw new Error('Invalid --target. Use dev or prod.');
      }
      target = value;
      continue;
    }
    if (current === '--allow-prod') {
      allowProd = true;
      continue;
    }
    prismaArgs.push(current);
  }

  return {
    help: false,
    target,
    allowProd,
    prismaArgs,
  };
}

function isProdReadonlyCommand(prismaArgs) {
  if (prismaArgs[0] === 'generate') return true;
  if (prismaArgs[0] === 'validate') return true;
  if (prismaArgs[0] === 'format') return true;
  if (prismaArgs[0] === 'db' && prismaArgs[1] === 'pull') return true;
  if (prismaArgs[0] === 'migrate' && prismaArgs[1] === 'status') return true;
  if (prismaArgs[0] === 'migrate' && prismaArgs[1] === 'diff') return true;
  return false;
}

function assertSafeCommand(target, prismaArgs, allowProd) {
  if (!prismaArgs.length) {
    throw new Error('Missing Prisma CLI arguments.');
  }

  if (target !== 'prod') return;

  if (prismaArgs[0] === 'migrate' && prismaArgs[1] === 'dev') {
    throw new Error('Refusing to run `prisma migrate dev` against prod. Use migrate deploy.');
  }

  if (isProdReadonlyCommand(prismaArgs)) return;

  if (!allowProd) {
    throw new Error(
      'Refusing to run a mutating Prisma command against prod without --allow-prod.',
    );
  }
}

function main() {
  try {
    const parsed = parseArgs(process.argv.slice(2));
    if (parsed.help) {
      printUsage();
      process.exit(0);
    }

    assertSafeCommand(parsed.target, parsed.prismaArgs, parsed.allowProd);
    runPrismaPostgresCli(parsed.prismaArgs, parsed.target, { cwd: root });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[prisma-postgres] ${message}`);
    process.exit(1);
  }
}

main();
