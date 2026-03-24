const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { config: loadDotenv } = require('dotenv');
const { Client } = require('pg');

const LOGIN_PATTERN = /^[a-z0-9._-]{3,64}$/;
const TOKEN_TTL_MS = 24 * 60 * 60 * 1000;

function printUsage() {
  console.log(`
Usage:
  node scripts/seed-postgres-super-admin.js --login <login> --full-name "<Full Name>" [--target dev|prod] [--allow-prod]

Examples:
  node scripts/seed-postgres-super-admin.js --login root.admin --full-name "Root Admin"
  node scripts/seed-postgres-super-admin.js --target prod --allow-prod --login root.admin --full-name "Root Admin"
  `);
}

function parseArgs(argv) {
  const args = [...argv];
  const parsed = {
    target: 'dev',
    allowProd: false,
    login: '',
    fullName: '',
    help: false,
  };

  while (args.length > 0) {
    const current = args.shift();
    if (current === '--help' || current === '-h') {
      parsed.help = true;
      return parsed;
    }
    if (current === '--target') {
      const value = String(args.shift() || '').trim().toLowerCase();
      if (value !== 'dev' && value !== 'prod') {
        throw new Error('Invalid --target. Use dev or prod.');
      }
      parsed.target = value;
      continue;
    }
    if (current === '--allow-prod') {
      parsed.allowProd = true;
      continue;
    }
    if (current === '--login') {
      parsed.login = String(args.shift() || '').trim().toLowerCase();
      continue;
    }
    if (current === '--full-name') {
      parsed.fullName = String(args.shift() || '').trim().replace(/\s+/g, ' ');
      continue;
    }
    throw new Error(`Unknown argument: ${current}`);
  }

  return parsed;
}

function resolveEnvFile(target) {
  const root = path.resolve(__dirname, '..');
  const envFileName = target === 'prod' ? '.env' : '.env.test';
  return {
    root,
    envFilePath: path.join(root, envFileName),
  };
}

function loadDatabaseUrl(target) {
  const ctx = resolveEnvFile(target);
  if (!fs.existsSync(ctx.envFilePath)) {
    throw new Error(`Env file not found: ${ctx.envFilePath}`);
  }
  loadDotenv({ path: ctx.envFilePath, quiet: true });
  const databaseUrl = String(process.env.POSTGRES_DATABASE_URL || '').trim();
  if (!databaseUrl) {
    throw new Error(`Missing POSTGRES_DATABASE_URL in ${ctx.envFilePath}`);
  }
  return { ...ctx, databaseUrl };
}

function validateInput(login, fullName) {
  if (!login) throw new Error('Missing --login.');
  if (!LOGIN_PATTERN.test(login)) {
    throw new Error(
      'Invalid login. Use 3-64 chars: lowercase letters, digits, dot, underscore, dash.',
    );
  }
  if (!fullName || fullName.length < 3 || fullName.length > 160) {
    throw new Error('Invalid --full-name. Use 3-160 characters.');
  }
}

function hashToken(token) {
  return crypto.createHash('sha256').update(token).digest('hex');
}

function generateToken() {
  const token = crypto.randomBytes(24).toString('base64url');
  return {
    token,
    tokenHash: hashToken(token),
    tokenPreview: `${token.slice(0, 6)}...${token.slice(-4)}`,
    expiresAt: new Date(Date.now() + TOKEN_TTL_MS),
  };
}

async function ensureAuthTablesExist(client) {
  const result = await client.query(
    `SELECT to_regclass('"public"."AuthUser"') AS "authUser", to_regclass('"public"."AuthOneTimeToken"') AS "tokenTable"`,
  );
  const row = result.rows[0] || {};
  if (!row.authUser || !row.tokenTable) {
    throw new Error(
      'Auth tables do not exist yet. Run Prisma migrate deploy/db push first.',
    );
  }
}

async function main() {
  try {
    const args = parseArgs(process.argv.slice(2));
    if (args.help) {
      printUsage();
      process.exit(0);
    }

    if (args.target === 'prod' && !args.allowProd) {
      throw new Error(
        'Refusing to seed prod without --allow-prod.',
      );
    }

    validateInput(args.login, args.fullName);
    const { databaseUrl, envFilePath } = loadDatabaseUrl(args.target);
    const client = new Client({ connectionString: databaseUrl });
    await client.connect();

    try {
      await ensureAuthTablesExist(client);
      await client.query('BEGIN');

      const existingSuperAdmin = await client.query(
        'SELECT "id", "login" FROM "AuthUser" WHERE "systemRole" = $1 LIMIT 1',
        ['SUPER_ADMIN'],
      );

      if (existingSuperAdmin.rowCount > 0) {
        throw new Error(
          `Super admin already exists (${existingSuperAdmin.rows[0].login}). Use admin panel token reset instead of seed.`,
        );
      }

      const existingLogin = await client.query(
        'SELECT "id" FROM "AuthUser" WHERE "login" = $1 LIMIT 1',
        [args.login],
      );

      if (existingLogin.rowCount > 0) {
        throw new Error(`Login already exists: ${args.login}`);
      }

      const userId = crypto.randomUUID();
      const tokenId = crypto.randomUUID();
      const token = generateToken();
      const now = new Date();

      await client.query(
        `
          INSERT INTO "AuthUser" (
            "id", "login", "fullName", "systemRole", "isActive",
            "passwordHash", "passwordSalt", "passwordSetAt",
            "mustSetPassword", "createdAt", "updatedAt"
          )
          VALUES ($1, $2, $3, 'SUPER_ADMIN', true, NULL, NULL, NULL, true, $4, $4)
        `,
        [userId, args.login, args.fullName, now],
      );

      await client.query(
        `
          INSERT INTO "AuthOneTimeToken" (
            "id", "userId", "issuedById", "tokenHash", "tokenPreview",
            "expiresAt", "usedAt", "revokedAt", "createdAt"
          )
          VALUES ($1, $2, NULL, $3, $4, $5, NULL, NULL, $6)
        `,
        [
          tokenId,
          userId,
          token.tokenHash,
          token.tokenPreview,
          token.expiresAt,
          now,
        ],
      );

      await client.query('COMMIT');

      console.log('[seed-super-admin] Super admin seeded successfully.');
      console.log(`[seed-super-admin] Env file: ${envFilePath}`);
      console.log(`[seed-super-admin] Login: ${args.login}`);
      console.log(`[seed-super-admin] Full Name: ${args.fullName}`);
      console.log(`[seed-super-admin] Token expires at: ${token.expiresAt.toISOString()}`);
      console.log('[seed-super-admin] One-time token (store securely, shown once):');
      console.log(token.token);
    } catch (error) {
      await client.query('ROLLBACK').catch(() => {});
      throw error;
    } finally {
      await client.end();
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[seed-super-admin] ${message}`);
    process.exit(1);
  }
}

void main();
