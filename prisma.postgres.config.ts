import * as fs from 'fs';
import * as path from 'path';
import { config as loadDotenv } from 'dotenv';
import { defineConfig } from 'prisma/config';

const envTarget = process.env.PRISMA_ENV_TARGET === 'prod' ? 'prod' : 'dev';
const envFileName = envTarget === 'prod' ? '.env' : '.env.test';
const envFilePath = path.resolve(__dirname, envFileName);

if (fs.existsSync(envFilePath)) {
  loadDotenv({ path: envFilePath, quiet: true });
}

export default defineConfig({
  schema: 'prisma-postgres/schema.prisma',
  datasource: {
    url: process.env.POSTGRES_DATABASE_URL,
  },
});
