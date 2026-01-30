import type { Prisma } from '@prisma/client';
import { PrismaClient } from '@prisma/client';
import { PrismaBetterSqlite3 } from '@prisma/adapter-better-sqlite3';
import { app } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import * as xlsx from 'xlsx';
import { RAPORT_COLUMNS } from '../raportColumns';

const DEFAULT_PAGE_SIZE = 250;
const IMPORT_BATCH_SIZE = 50;

let prisma: PrismaClient | null = null;

type ImportProgress = {
  stage: 'reading' | 'parsing' | 'importing' | 'finalizing' | 'done';
  message: string;
  current: number;
  total: number;
};

function toSqliteUrl(dbFilePath: string): string {
  const normalized = dbFilePath.replace(/\\/g, '/');
  return `file:${normalized}`;
}

function resolvePrismaQueryEnginePath(): string | null {
  if (process.platform !== 'win32') return null;
  const engineName = 'query_engine-windows.dll.node';

  const candidates = [
    path.join(__dirname, 'native_modules', 'client', engineName),
    path.join(__dirname, '.prisma', 'client', engineName),
  ];

  const expanded: string[] = [];
  for (const c of candidates) {
    if (c.includes('app.asar')) expanded.push(c.replace('app.asar', 'app.asar.unpacked'));
    expanded.push(c);
  }

  for (const c of expanded) {
    const isAsarPath = c.includes('app.asar') && !c.includes('app.asar.unpacked');
    if (isAsarPath) continue; // never load .node directly from asar
    if (fs.existsSync(c)) return c;
  }
  return null;
}

export function getDbFilePath(): string {
  if (!app.isPackaged) {
    return path.join(process.cwd(), 'db', 'dev.sqlite');
  }
  return path.join(app.getPath('userData'), 'raport.sqlite');
}

export async function getPrisma(): Promise<PrismaClient> {
  if (prisma) return prisma;

  const dbFilePath = getDbFilePath();
  fs.mkdirSync(path.dirname(dbFilePath), { recursive: true });
  fs.closeSync(fs.openSync(dbFilePath, 'a'));

  // Useful for tools/debugging (not required when using adapter).
  process.env.DATABASE_URL = toSqliteUrl(dbFilePath);

  const enginePath = resolvePrismaQueryEnginePath();
  if (enginePath) process.env.PRISMA_QUERY_ENGINE_LIBRARY = enginePath;

  const adapter = new PrismaBetterSqlite3({ url: toSqliteUrl(dbFilePath) });
  prisma = new PrismaClient({ adapter });
  await ensureSchema(prisma);
  return prisma;
}

async function ensureSchema(client: PrismaClient): Promise<void> {
  const createMeta = `
    CREATE TABLE IF NOT EXISTS "raport_meta" (
      "id" INTEGER PRIMARY KEY,
      "importedAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "sourceFile" TEXT NULL,
      "rowCount" INTEGER NOT NULL DEFAULT 0
    );
  `;

  const cols = RAPORT_COLUMNS.map((c) => `"${c.field}" TEXT NULL`).join(',\n      ');
  const createRows = `
    CREATE TABLE IF NOT EXISTS "raport_rows" (
      "id" INTEGER PRIMARY KEY AUTOINCREMENT,
      "rowNumber" INTEGER NULL,
      ${cols}
    );
  `;

  await client.$executeRawUnsafe(createMeta);
  await client.$executeRawUnsafe(createRows);

  await client.raportMeta.upsert({
    where: { id: 1 },
    create: { id: 1, rowCount: 0, sourceFile: null },
    update: {},
  });
}

function normalizeCell(value: unknown): string | null {
  if (value == null) return null;
  const s = String(value).trim();
  return s.length === 0 ? null : s;
}

function findHeaderRow(rows: unknown[][]): number {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const hasAny = row.some((cell) => normalizeCell(cell) !== null);
    if (hasAny) return i;
  }
  return 0;
}

function isEmptyRow(row: unknown[]): boolean {
  return !(row || []).some((cell) => normalizeCell(cell) !== null);
}

export async function importRaportFromXlsx(filePath: string): Promise<{ rowCount: number; sourceFile: string }>;
export async function importRaportFromXlsx(
  filePath: string,
  onProgress: (p: ImportProgress) => void,
): Promise<{ rowCount: number; sourceFile: string }>;
export async function importRaportFromXlsx(
  filePath: string,
  onProgress?: (p: ImportProgress) => void,
): Promise<{ rowCount: number; sourceFile: string }> {
  const client = await getPrisma();
  const report = (p: ImportProgress) => onProgress?.(p);

  report({ stage: 'reading', message: 'Wczytywanie pliku…', current: 0, total: 0 });

  try {
    fs.accessSync(filePath, fs.constants.R_OK);
  } catch {
    throw new Error(`Nie można odczytać pliku: ${filePath}`);
  }

  const wb = xlsx.readFile(filePath, { cellDates: true });

  report({ stage: 'parsing', message: 'Analiza arkusza…', current: 0, total: 0 });

  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const aoa = xlsx.utils.sheet_to_json(ws, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: false,
  }) as unknown[][];

  const headerRowIndex = findHeaderRow(aoa);
  const startDataIndex = headerRowIndex + 1;
  const dataRows = aoa.slice(startDataIndex);

  const totalRows = dataRows.length;
  report({ stage: 'importing', message: 'Importowanie…', current: 0, total: totalRows });

  await clearRaportData();
  await new Promise<void>((resolve) => setImmediate(resolve));

  const operations: Prisma.PrismaPromise<unknown>[] = [];
  let rowCount = 0;
  let processed = 0;

  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i] || [];
    processed++;
    if (isEmptyRow(row)) continue;

    const record: Record<string, string | number | null> = {
      rowNumber: i + 1,
    };

    for (let colIndex = 0; colIndex < RAPORT_COLUMNS.length; colIndex++) {
      const col = RAPORT_COLUMNS[colIndex];
      const value = normalizeCell(row[colIndex]);
      if (value !== null) record[col.field] = value;
    }

    operations.push(
      client.raportRow.create({
        data: record as Prisma.RaportRowUncheckedCreateInput,
      }),
    );
    rowCount++;

    if (operations.length >= IMPORT_BATCH_SIZE) {
      await client.$transaction(operations);
      operations.length = 0;
      report({ stage: 'importing', message: 'Importowanie…', current: processed, total: totalRows });
      await new Promise<void>((resolve) => setImmediate(resolve));
    }
  }

  if (operations.length > 0) {
    await client.$transaction(operations);
    report({ stage: 'importing', message: 'Importowanie…', current: processed, total: totalRows });
    await new Promise<void>((resolve) => setImmediate(resolve));
  }

  report({ stage: 'finalizing', message: 'Zapisywanie informacji…', current: processed, total: totalRows });

  await client.raportMeta.update({
    where: { id: 1 },
    data: {
      importedAt: new Date(),
      sourceFile: path.basename(filePath),
      rowCount,
    },
  });

  report({ stage: 'done', message: 'Zakończono.', current: processed, total: totalRows });
  return { rowCount, sourceFile: path.basename(filePath) };
}

export async function clearRaportData(): Promise<void> {
  const client = await getPrisma();

  await client.$executeRawUnsafe('DELETE FROM "raport_rows";');
  try {
    await client.$executeRawUnsafe('DELETE FROM sqlite_sequence WHERE name = "raport_rows";');
  } catch {
    // sqlite_sequence may not exist yet
  }

  await client.raportMeta.update({
    where: { id: 1 },
    data: {
      importedAt: new Date(),
      sourceFile: null,
      rowCount: 0,
    },
  });
}

export async function getRaportMeta(): Promise<{ importedAt: string | null; sourceFile: string | null; rowCount: number }> {
  const client = await getPrisma();
  const meta = await client.raportMeta.findUnique({ where: { id: 1 } });
  return {
    importedAt: meta?.importedAt ? meta.importedAt.toISOString() : null,
    sourceFile: meta?.sourceFile ?? null,
    rowCount: meta?.rowCount ?? 0,
  };
}

export async function getDbInfo(): Promise<{ filePath: string; exists: boolean }> {
  const filePath = getDbFilePath();
  return { filePath, exists: fs.existsSync(filePath) };
}

export async function getRaportPage(params: {
  page: number;
  pageSize?: number;
}): Promise<{
  page: number;
  pageSize: number;
  total: number;
  columns: Array<{ field: string; label: string }>;
  rows: Array<Record<string, string | null>>;
}> {
  const client = await getPrisma();
  const pageSize = Math.max(1, Math.min(1000, params.pageSize ?? DEFAULT_PAGE_SIZE));
  const page = Math.max(1, params.page);
  const skip = (page - 1) * pageSize;

  const [total, rows] = await Promise.all([
    client.raportRow.count(),
    client.raportRow.findMany({
      skip,
      take: pageSize,
      orderBy: { id: 'asc' },
    }),
  ]);

  const normalizedRows = rows.map((r) => {
    const rAny = r as unknown as Record<string, unknown>;
    const out: Record<string, string | null> = { id: String(r.id) };
    for (const col of RAPORT_COLUMNS) {
      out[col.field] = (rAny[col.field] as string | null | undefined) ?? null;
    }
    return out;
  });

  return {
    page,
    pageSize,
    total,
    columns: RAPORT_COLUMNS.map((c) => ({ field: c.field, label: c.label })),
    rows: normalizedRows,
  };
}

export async function disposePrisma(): Promise<void> {
  if (!prisma) return;
  await prisma.$disconnect();
  prisma = null;
}
