import { PrismaPg } from '@prisma/adapter-pg';
import { app } from 'electron';
import { config as loadDotenv } from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';
import { PrismaClient } from '../generated/postgres-prisma/client';

let postgresPrisma: PrismaClient | null = null;
let postgresEnvLoaded = false;

function resolvePostgresEnvFilePath(): string {
  if (app.isPackaged) {
    return path.resolve(process.resourcesPath, '.env');
  }

  return path.resolve(process.cwd(), '.env.test');
}

function loadPostgresEnvFile(): string {
  const envFilePath = resolvePostgresEnvFilePath();

  if (!postgresEnvLoaded) {
    if (!fs.existsSync(envFilePath)) {
      throw new Error(`Postgres env file not found: ${envFilePath}`);
    }

    loadDotenv({ path: envFilePath, quiet: true });
    postgresEnvLoaded = true;
  }

  return envFilePath;
}

export function getPostgresPrisma(): PrismaClient {
  if (postgresPrisma) return postgresPrisma;

  const envFilePath = loadPostgresEnvFile();
  const databaseUrl = process.env.POSTGRES_DATABASE_URL?.trim();
  if (!databaseUrl) {
    throw new Error(`Missing POSTGRES_DATABASE_URL in ${envFilePath}`);
  }

  const adapter = new PrismaPg({ connectionString: databaseUrl });
  postgresPrisma = new PrismaClient({ adapter });
  return postgresPrisma;
}

export async function connectPostgresPrisma(): Promise<PrismaClient> {
  const client = getPostgresPrisma();
  await client.$connect();
  return client;
}

export async function disposePostgresPrisma(): Promise<void> {
  if (!postgresPrisma) return;

  await postgresPrisma.$disconnect();
  postgresPrisma = null;
}
