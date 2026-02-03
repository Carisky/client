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

  const createMrnBatch = `
    CREATE TABLE IF NOT EXISTS "mrn_batch" (
      "id" INTEGER PRIMARY KEY AUTOINCREMENT,
      "createdAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      "numer_mrn" TEXT NOT NULL,
      "raportRowId" INTEGER NULL,
      "rowNumber" INTEGER NULL,
      "nr_sad" TEXT NULL,
      "data_mrn" TEXT NULL,
      "zglaszajacy" TEXT NULL
    );
  `;

  const createValidationManual = `
    CREATE TABLE IF NOT EXISTS "validation_manual" (
      "raportRowId" INTEGER PRIMARY KEY,
      "verifiedAt" DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
  `;

  await client.$executeRawUnsafe(createMeta);
  await client.$executeRawUnsafe(createRows);
  await client.$executeRawUnsafe(createMrnBatch);
  await client.$executeRawUnsafe(createValidationManual);

  await client.$executeRawUnsafe('CREATE INDEX IF NOT EXISTS "mrn_batch_numer_mrn_idx" ON "mrn_batch" ("numer_mrn");');
  await client.$executeRawUnsafe('CREATE INDEX IF NOT EXISTS "mrn_batch_raportRowId_idx" ON "mrn_batch" ("raportRowId");');

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

  await clearMrnBatch();

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

async function clearMrnBatch(): Promise<void> {
  const client = await getPrisma();

  await client.$executeRawUnsafe('DELETE FROM "mrn_batch";');
  try {
    await client.$executeRawUnsafe('DELETE FROM sqlite_sequence WHERE name = "mrn_batch";');
  } catch {
    // sqlite_sequence may not exist yet
  }
}

export async function rebuildMrnBatch(): Promise<{ rowsInserted: number; groups: number; scannedAt: string | null }> {
  const client = await getPrisma();

  await clearMrnBatch();

  const inserted = (await client.$executeRawUnsafe(`
    INSERT INTO "mrn_batch" ("numer_mrn", "raportRowId", "rowNumber", "nr_sad", "data_mrn", "zglaszajacy")
    SELECT
      TRIM("numer_mrn") as "numer_mrn",
      "id" as "raportRowId",
      "rowNumber",
      "nr_sad",
      "data_mrn",
      "zglaszajacy"
    FROM "raport_rows"
    WHERE "numer_mrn" IS NOT NULL
      AND TRIM("numer_mrn") <> ''
      AND TRIM("numer_mrn") IN (
        SELECT TRIM("numer_mrn") as "numer_mrn"
        FROM "raport_rows"
        WHERE "numer_mrn" IS NOT NULL AND TRIM("numer_mrn") <> ''
        GROUP BY TRIM("numer_mrn")
        HAVING COUNT(*) > 1
      );
  `)) as unknown as number;

  const groupsRows = (await client.$queryRawUnsafe(`
    SELECT COUNT(*) as "cnt"
    FROM (
      SELECT "numer_mrn"
      FROM "mrn_batch"
      GROUP BY "numer_mrn"
    );
  `)) as Array<{ cnt: number }>;

  const scannedAtRows = (await client.$queryRawUnsafe(`
    SELECT MAX("createdAt") as "maxCreatedAt"
    FROM "mrn_batch";
  `)) as Array<{ maxCreatedAt: string | null }>;

  return {
    rowsInserted: Number.isFinite(inserted) ? inserted : 0,
    groups: groupsRows?.[0]?.cnt ?? 0,
    scannedAt: scannedAtRows?.[0]?.maxCreatedAt ?? null,
  };
}

export async function getMrnBatchMeta(): Promise<{ scannedAt: string | null; groups: number; rows: number }> {
  const client = await getPrisma();

  const rows = await client.mrnBatchRow.count();
  const groupsRows = (await client.$queryRawUnsafe(`
    SELECT COUNT(*) as "cnt"
    FROM (
      SELECT "numer_mrn"
      FROM "mrn_batch"
      GROUP BY "numer_mrn"
    );
  `)) as Array<{ cnt: number }>;

  const scannedAtRows = (await client.$queryRawUnsafe(`
    SELECT MAX("createdAt") as "maxCreatedAt"
    FROM "mrn_batch";
  `)) as Array<{ maxCreatedAt: string | null }>;

  return {
    scannedAt: scannedAtRows?.[0]?.maxCreatedAt ?? null,
    groups: groupsRows?.[0]?.cnt ?? 0,
    rows,
  };
}

export async function getMrnBatchGroups(params?: {
  limit?: number;
}): Promise<Array<{ numer_mrn: string; count: number }>> {
  const client = await getPrisma();
  const limit = Math.max(1, Math.min(2000, params?.limit ?? 500));

  const rows = (await client.$queryRawUnsafe(
    `
      SELECT "numer_mrn" as "numer_mrn", COUNT(*) as "count"
      FROM "mrn_batch"
      GROUP BY "numer_mrn"
      ORDER BY "count" DESC, "numer_mrn" ASC
      LIMIT ?;
    `,
    limit,
  )) as Array<{ numer_mrn: string; count: number }>;

  return rows.map((r) => ({
    numer_mrn: String(r.numer_mrn),
    count: Number(r.count) || 0,
  }));
}

export async function getMrnBatchRows(numerMrn: string): Promise<{
  numer_mrn: string;
  rows: Array<Record<string, string | null>>;
}> {
  const client = await getPrisma();
  const numer_mrn = String(numerMrn ?? '').trim();
  if (!numer_mrn) return { numer_mrn: '', rows: [] };

  const batch = await client.mrnBatchRow.findMany({
    where: { numer_mrn },
    select: { raportRowId: true },
    orderBy: { raportRowId: 'asc' },
  });
  const ids = batch.map((b) => b.raportRowId).filter((v): v is number => typeof v === 'number');
  if (ids.length === 0) return { numer_mrn, rows: [] };

  const rows = await client.raportRow.findMany({
    where: { id: { in: ids } },
    orderBy: { id: 'asc' },
  });

  const normalized = rows.map((r) => {
    const rAny = r as unknown as Record<string, unknown>;
    const out: Record<string, string | null> = {
      id: String(r.id),
      rowNumber: r.rowNumber == null ? null : String(r.rowNumber),
    };
    for (const col of RAPORT_COLUMNS) {
      out[col.field] = (rAny[col.field] as string | null | undefined) ?? null;
    }
    return out;
  });

  return { numer_mrn, rows: normalized };
}

function toYmdRange(period: string): { start: string; end: string } {
  const m = String(period ?? '').trim();

  const yearMatch = /^(\d{4})$/.exec(m);
  if (yearMatch) return { start: `${yearMatch[1]}-01-01`, end: `${yearMatch[1]}-12-31` };

  const match = /^(\d{4})-(0[1-9]|1[0-2])$/.exec(m);
  if (!match) {
    // fallback: current month
    const now = new Date();
    const y = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const start = `${y}-${mm}-01`;
    const end = `${y}-${mm}-${String(new Date(y, now.getMonth() + 1, 0).getDate()).padStart(2, '0')}`;
    return { start, end };
  }
  const y = Number(match[1]);
  const mon = Number(match[2]);
  const lastDay = new Date(y, mon, 0).getDate();
  return { start: `${match[1]}-${match[2]}-01`, end: `${match[1]}-${match[2]}-${String(lastDay).padStart(2, '0')}` };
}

type ValidationDateGrouping = 'day' | 'days2' | 'days3' | 'week' | 'month' | 'months2';

function normalizeValidationDateGrouping(value: unknown): ValidationDateGrouping {
  const v = String(value ?? '').trim();
  if (v === 'days2') return 'days2';
  if (v === 'days3') return 'days3';
  if (v === 'week') return 'week';
  if (v === 'month') return 'month';
  if (v === 'months2') return 'months2';
  return 'day';
}

function parseYmdStrict(ymd: string): { y: number; m: number; d: number } | null {
  const m = /^(\d{4})-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/.exec(String(ymd ?? '').trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mon = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(Date.UTC(y, mon - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mon - 1 || dt.getUTCDate() !== d) return null;
  return { y, m: mon, d };
}

const MS_PER_DAY = 24 * 60 * 60 * 1000;

function ymdToUtcMs(ymd: string): number | null {
  const p = parseYmdStrict(ymd);
  if (!p) return null;
  return Date.UTC(p.y, p.m - 1, p.d);
}

function utcMsToYmd(ms: number): string {
  return new Date(ms).toISOString().slice(0, 10);
}

function addDaysYmd(ymd: string, days: number): string | null {
  const ms = ymdToUtcMs(ymd);
  if (ms == null) return null;
  const d = Number(days);
  if (!Number.isFinite(d)) return null;
  return utcMsToYmd(ms + Math.trunc(d) * MS_PER_DAY);
}

function addMonthsYmdClamped(ymd: string, months: number): string | null {
  const p = parseYmdStrict(ymd);
  if (!p) return null;
  const m = Number(months);
  if (!Number.isFinite(m)) return null;
  const targetFirst = new Date(Date.UTC(p.y, p.m - 1 + Math.trunc(m), 1));
  const lastDay = new Date(Date.UTC(targetFirst.getUTCFullYear(), targetFirst.getUTCMonth() + 1, 0)).getUTCDate();
  const dd = Math.min(p.d, lastDay);
  return utcMsToYmd(Date.UTC(targetFirst.getUTCFullYear(), targetFirst.getUTCMonth(), dd));
}

function getValidationGroupingConfig(params: { grouping?: unknown }): { grouping: ValidationDateGrouping } {
  const grouping = normalizeValidationDateGrouping(params.grouping);
  return { grouping };
}

function bucketEndFromStart(startYmd: string, grouping: ValidationDateGrouping): string | null {
  if (!startYmd) return null;
  if (grouping === 'day') return startYmd;
  if (grouping === 'days2') return addDaysYmd(startYmd, 1);
  if (grouping === 'days3') return addDaysYmd(startYmd, 2);
  if (grouping === 'week') return addDaysYmd(startYmd, 6);
  if (grouping === 'month') {
    const next = addMonthsYmdClamped(startYmd, 1);
    return next ? addDaysYmd(next, -1) : null;
  }
  if (grouping === 'months2') {
    const next = addMonthsYmdClamped(startYmd, 2);
    return next ? addDaysYmd(next, -1) : null;
  }
  return startYmd;
}

function bucketStartForDate(dateYmd: string, anchorYmd: string, grouping: ValidationDateGrouping): string | null {
  const dateMs = ymdToUtcMs(dateYmd);
  if (dateMs == null) return null;

  if (grouping === 'day') return dateYmd;

  if (grouping === 'days2' || grouping === 'days3' || grouping === 'week') {
    const anchorMs = ymdToUtcMs(anchorYmd);
    if (anchorMs == null) return dateYmd;
    const sizeDays = grouping === 'days2' ? 2 : grouping === 'days3' ? 3 : 7;
    const diffDays = Math.floor((dateMs - anchorMs) / MS_PER_DAY);
    const bucketIdx = Math.floor(diffDays / sizeDays);
    return utcMsToYmd(anchorMs + bucketIdx * sizeDays * MS_PER_DAY);
  }

  const stepMonths = grouping === 'month' ? 1 : 2;
  const anchorParsed = parseYmdStrict(anchorYmd);
  if (!anchorParsed) return dateYmd;
  let start = anchorYmd;

  const cmp = dateYmd.localeCompare(anchorYmd);
  if (cmp >= 0) {
    for (let i = 0; i < 60; i++) {
      const next = addMonthsYmdClamped(start, stepMonths);
      if (!next) break;
      if (next.localeCompare(dateYmd) > 0) break;
      start = next;
    }
    return start;
  }

  for (let i = 0; i < 60; i++) {
    const prev = addMonthsYmdClamped(start, -stepMonths);
    if (!prev) break;
    if (prev.localeCompare(dateYmd) <= 0) return prev;
    start = prev;
  }
  return start;
}

type ValidationCohortItem = { key: ValidationGroupKey; bucketStart: string | null; numer_mrn: string | null };

function cohortKey(item: Pick<ValidationCohortItem, 'key' | 'bucketStart'>): string | null {
  if (!item.bucketStart) return null;
  return `${JSON.stringify(item.key)}|${item.bucketStart}`;
}

function filterToMrnCohorts<T extends ValidationCohortItem>(items: T[], mrnNorm: string): T[] {
  if (!mrnNorm) return items;
  const cohorts = new Set<string>();
  for (const it of items) {
    const k = cohortKey(it);
    if (!k) continue;
    if (!mrnContains(it.numer_mrn, mrnNorm)) continue;
    cohorts.add(k);
  }
  if (cohorts.size === 0) return [];
  return items.filter((it) => {
    const k = cohortKey(it);
    return k != null && cohorts.has(k);
  });
}

function filterBucketStartsByMrn(
  items: Array<{ bucketStart: string | null; numer_mrn: string | null }>,
  mrnNorm: string,
): Set<string> {
  const set = new Set<string>();
  if (!mrnNorm) return set;
  for (const it of items) {
    if (!it.bucketStart) continue;
    if (!mrnContains(it.numer_mrn, mrnNorm)) continue;
    set.add(it.bucketStart);
  }
  return set;
}

function parseLocaleNumber(value: string | null | undefined): number | null {
  if (value == null) return null;
  let s = String(value).trim();
  if (!s) return null;

  s = s.replace(/\u00a0/g, ' ').replace(/\s+/g, '');

  const hasComma = s.includes(',');
  const hasDot = s.includes('.');
  if (hasComma && hasDot) {
    const lastComma = s.lastIndexOf(',');
    const lastDot = s.lastIndexOf('.');
    const dec = lastComma > lastDot ? ',' : '.';
    const thou = dec === ',' ? '.' : ',';
    s = s.split(thou).join('');
    if (dec === ',') s = s.replace(/,/g, '.');
  } else if (hasComma && !hasDot) {
    s = s.replace(/,/g, '.');
  }

  s = s.replace(/[^0-9.+-]/g, '');
  const n = Number.parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function quantileSorted(valuesAsc: number[], p: number): number {
  if (valuesAsc.length === 0) return Number.NaN;
  if (valuesAsc.length === 1) return valuesAsc[0];

  const pp = Math.min(1, Math.max(0, p));
  const idx = (valuesAsc.length - 1) * pp;
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return valuesAsc[lo];
  const w = idx - lo;
  return valuesAsc[lo] * (1 - w) + valuesAsc[hi] * w;
}

async function getValidationManualSet(client: PrismaClient): Promise<Set<number>> {
  const rows = (await client.$queryRawUnsafe(`
    SELECT "raportRowId" as "id"
    FROM "validation_manual";
  `)) as Array<{ id: number }>;
  const set = new Set<number>();
  for (const r of rows) {
    const id = Number(r?.id);
    if (Number.isFinite(id)) set.add(id);
  }
  return set;
}

export type ValidationGroupKey = {
  odbiorca: string;
  kraj_wysylki: string;
  warunki_dostawy: string;
  waluta: string;
  kurs_waluty: string;
  transport_na_granicy_rodzaj: string;
  kod_towaru: string;
};

function normalizeMrnQuery(value: unknown): string {
  if (typeof value !== 'string') return '';
  return value.trim().toUpperCase().replace(/[\s-]+/g, '');
}

function mrnContains(numerMrn: string | null, queryNorm: string): boolean {
  if (!queryNorm) return true;
  const v = String(numerMrn ?? '').trim().toUpperCase().replace(/[\s-]+/g, '');
  return v.includes(queryNorm);
}

export async function getValidationDefaultMonth(): Promise<{ month: string | null }> {
  const client = await getPrisma();
  const rows = (await client.$queryRawUnsafe(`
    SELECT MAX(SUBSTR(TRIM("data_mrn"), 1, 7)) as "month"
    FROM "raport_rows"
    WHERE "data_mrn" IS NOT NULL AND TRIM("data_mrn") <> '';
  `)) as Array<{ month: string | null }>;
  const month = rows?.[0]?.month ?? null;
  return { month };
}

async function queryValidationRepresentativeRows(client: PrismaClient, range: { start: string; end: string }) {
  type ValidationSadRawRow = {
    id: number;
    data_mrn: string | null;
    numer_mrn: string | null;
    nr_sad: string | null;
    zglaszajacy: string | null;
    odbiorca: string | null;
    kraj_wysylki: string | null;
    warunki_dostawy: string | null;
    waluta: string | null;
    kurs_waluty: string | null;
    transport_na_granicy_rodzaj: string | null;
    kod_towaru: string | null;
    wartosc_faktury: string | null;
    wartosc_pozycji: string | null;
    oplaty_celne_razem: string | null;
    masa_netto: string | null;
  };

  type ValidationSadTextField = Exclude<
    keyof ValidationSadRawRow,
    'id' | 'wartosc_faktury' | 'wartosc_pozycji' | 'oplaty_celne_razem' | 'masa_netto'
  >;
  type ValidationSadNumericField = 'wartosc_faktury' | 'wartosc_pozycji' | 'oplaty_celne_razem' | 'masa_netto';

  const raw = (await client.$queryRawUnsafe(
    `
      SELECT
        "id",
        "data_mrn",
        "numer_mrn",
        "nr_sad",
        "zglaszajacy",
        "odbiorca",
        "kraj_wysylki",
        "warunki_dostawy",
        "waluta",
        "kurs_waluty",
        "transport_na_granicy_rodzaj",
        "kod_towaru",
        "wartosc_faktury",
        "wartosc_pozycji",
        "oplaty_celne_razem",
        "masa_netto"
      FROM "raport_rows"
      WHERE "data_mrn" IS NOT NULL
        AND TRIM("data_mrn") <> ''
        AND TRIM("data_mrn") BETWEEN ? AND ?
        AND "numer_mrn" IS NOT NULL AND TRIM("numer_mrn") <> ''
        AND "nr_sad" IS NOT NULL AND TRIM("nr_sad") <> '';
    `,
    range.start,
    range.end,
  )) as ValidationSadRawRow[];

  // SAD może być wielopozycyjny (ten sam nr_sad pojawia się wiele razy).
  // Na potrzeby IQR agregujemy wartości liczbowe (np. masa_netto, opłaty) w jedną pozycję na (numer_mrn + nr_sad).
  const groupKey = (r: { numer_mrn: string | null; nr_sad: string | null }): string =>
    `${String(r.numer_mrn ?? '').trim()}|${String(r.nr_sad ?? '').trim()}`;

  const hasValue = (v: string | null | undefined): boolean => String(v ?? '').trim().length > 0;

  const repRank = (r: { wartosc_pozycji: string | null; wartosc_faktury: string | null }): number => {
    if (hasValue(r.wartosc_pozycji)) return 0;
    if (hasValue(r.wartosc_faktury)) return 1;
    return 2;
  };

  const mergeTextField = (rows: ValidationSadRawRow[], field: ValidationSadTextField): string | null => {
    const values = new Set<string>();
    for (const r of rows) {
      const v = String(r[field] ?? '').trim();
      if (!v) continue;
      values.add(v);
      if (values.size > 1) return 'MULTI';
    }
    return values.size === 1 ? Array.from(values)[0] : null;
  };

  const sumNumericField = (rows: ValidationSadRawRow[], field: ValidationSadNumericField): string | null => {
    let sum = 0;
    let any = false;
    for (const r of rows) {
      const n = parseLocaleNumber(r[field]);
      if (n == null || !Number.isFinite(n)) continue;
      sum += n;
      any = true;
    }
    return any ? String(sum) : null;
  };

  const bySad = new Map<string, ValidationSadRawRow[]>();
  for (const r of raw) {
    const k = groupKey(r);
    if (!k || k === '|') continue;
    const arr = bySad.get(k);
    if (arr) arr.push(r);
    else bySad.set(k, [r]);
  }

  const aggregated: Array<{
    id: number;
    data_mrn: string | null;
    numer_mrn: string | null;
    nr_sad: string | null;
    zglaszajacy: string | null;
    odbiorca: string | null;
    kraj_wysylki: string | null;
    warunki_dostawy: string | null;
    waluta: string | null;
    kurs_waluty: string | null;
    transport_na_granicy_rodzaj: string | null;
    kod_towaru: string | null;
    oplaty_celne_razem: string | null;
    masa_netto: string | null;
  }> = [];

  for (const rows of bySad.values()) {
    rows.sort((a, b) => repRank(a) - repRank(b) || a.id - b.id);
    const rep = rows[0];

    aggregated.push({
      id: rep.id,
      data_mrn: rep.data_mrn ?? null,
      numer_mrn: rep.numer_mrn ?? null,
      nr_sad: rep.nr_sad ?? null,
      // agent/odbiorca/... zwykle są stałe; jeśli jednak różnią się w pozycjach, ustawiamy "MULTI"
      zglaszajacy: mergeTextField(rows, 'zglaszajacy'),
      odbiorca: mergeTextField(rows, 'odbiorca'),
      kraj_wysylki: mergeTextField(rows, 'kraj_wysylki'),
      warunki_dostawy: mergeTextField(rows, 'warunki_dostawy'),
      waluta: mergeTextField(rows, 'waluta'),
      kurs_waluty: mergeTextField(rows, 'kurs_waluty'),
      transport_na_granicy_rodzaj: mergeTextField(rows, 'transport_na_granicy_rodzaj'),
      kod_towaru: mergeTextField(rows, 'kod_towaru'),
      oplaty_celne_razem: sumNumericField(rows, 'oplaty_celne_razem'),
      masa_netto: sumNumericField(rows, 'masa_netto'),
    });
  }

  return aggregated;
}

export async function getValidationGroups(params: {
  month: string;
  mrn?: string | null;
  grouping?: unknown;
}): Promise<{
  range: { start: string; end: string };
  groups: Array<{ key: ValidationGroupKey; count: number }>;
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);

  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const mrn = normalizeMrnQuery(params.mrn);

  const map = new Map<string, { key: ValidationGroupKey; count: number }>();
  if (!mrn) {
    for (const r of rows) {
      const key: ValidationGroupKey = {
        odbiorca: (r.odbiorca ?? '').trim(),
        kraj_wysylki: (r.kraj_wysylki ?? '').trim(),
        warunki_dostawy: (r.warunki_dostawy ?? '').trim(),
        waluta: (r.waluta ?? '').trim(),
        kurs_waluty: (r.kurs_waluty ?? '').trim(),
        transport_na_granicy_rodzaj: (r.transport_na_granicy_rodzaj ?? '').trim(),
        kod_towaru: (r.kod_towaru ?? '').trim(),
      };
      const k = JSON.stringify(key);
      const existing = map.get(k);
      if (existing) existing.count += 1;
      else map.set(k, { key, count: 1 });
    }
  } else {
    const cohorts = new Set<string>();
    for (const r of rows) {
      const data_mrn = r.data_mrn ?? null;
      const bucketStart = data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null;
      if (!bucketStart) continue;
      if (!mrnContains(r.numer_mrn, mrn)) continue;
      const key = toValidationKey(r);
      const ck = cohortKey({ key, bucketStart });
      if (ck) cohorts.add(ck);
    }

    for (const r of rows) {
      const data_mrn = r.data_mrn ?? null;
      const bucketStart = data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null;
      if (!bucketStart) continue;
      const key = toValidationKey(r);
      const ck = cohortKey({ key, bucketStart });
      if (!ck || !cohorts.has(ck)) continue;

      const k = JSON.stringify(key);
      const existing = map.get(k);
      if (existing) existing.count += 1;
      else map.set(k, { key, count: 1 });
    }
  }

  const groups = Array.from(map.values()).sort((a, b) => b.count - a.count);
  return { range, groups };
}

export async function getValidationItems(params: {
  month: string;
  key: ValidationGroupKey;
  mrn?: string | null;
  grouping?: unknown;
}): Promise<{
  range: { start: string; end: string };
  key: ValidationGroupKey;
  items: Array<{
    rowId: number;
    data_mrn: string | null;
    odbiorca: string | null;
    numer_mrn: string | null;
    coef: number | null;
    verifiedManual: boolean;
    checkable: boolean;
    outlier: boolean;
    outlierSide: 'low' | 'high' | null;
  }>;
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);

  const want = params.key;
  const items = rows
    .filter((r) => {
      const key: ValidationGroupKey = {
        odbiorca: (r.odbiorca ?? '').trim(),
        kraj_wysylki: (r.kraj_wysylki ?? '').trim(),
        warunki_dostawy: (r.warunki_dostawy ?? '').trim(),
        waluta: (r.waluta ?? '').trim(),
        kurs_waluty: (r.kurs_waluty ?? '').trim(),
        transport_na_granicy_rodzaj: (r.transport_na_granicy_rodzaj ?? '').trim(),
        kod_towaru: (r.kod_towaru ?? '').trim(),
      };
      return (
        key.odbiorca === want.odbiorca &&
        key.kraj_wysylki === want.kraj_wysylki &&
        key.warunki_dostawy === want.warunki_dostawy &&
        key.waluta === want.waluta &&
        key.kurs_waluty === want.kurs_waluty &&
        key.transport_na_granicy_rodzaj === want.transport_na_granicy_rodzaj &&
        key.kod_towaru === want.kod_towaru
      );
    })
    .map((r) => {
      const fees = parseLocaleNumber(r.oplaty_celne_razem);
      const mass = parseLocaleNumber(r.masa_netto);
      const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
      const rowIdRaw = Number(r.id);
      const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
      const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
      const data_mrn = r.data_mrn ?? null;
      return {
        rowId,
        data_mrn,
        bucketStart: data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null,
        odbiorca: r.odbiorca ?? null,
        numer_mrn: r.numer_mrn ?? null,
        coef,
        verifiedManual,
        checkable: false,
        outlier: false,
        outlierSide: null as 'low' | 'high' | null,
      };
    });

  // Outliers per day (IQR). If a day has <2 numeric values, it's ignored.
  const byDate = new Map<string, number[]>();
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.bucketStart) continue;
    if (it.verifiedManual) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const arr = byDate.get(it.bucketStart);
    if (arr) arr.push(i);
    else byDate.set(it.bucketStart, [i]);
  }

  for (const indices of byDate.values()) {
    if (indices.length < 2) continue;
    const valuesAsc = indices
      .map((i) => items[i].coef as number)
      .filter((v) => Number.isFinite(v))
      .sort((a, b) => a - b);
    if (valuesAsc.length < 2) continue;

    const q1 = quantileSorted(valuesAsc, 0.25);
    const q3 = quantileSorted(valuesAsc, 0.75);
    const iqr = q3 - q1;
    if (!Number.isFinite(iqr)) continue;
    const lower = q1 - 1.5 * iqr;
    const upper = q3 + 1.5 * iqr;

    for (const i of indices) {
      const v = items[i].coef;
      if (v == null || !Number.isFinite(v)) continue;
      items[i].checkable = true;
      if (v < lower) {
        items[i].outlier = true;
        items[i].outlierSide = 'low';
      } else if (v > upper) {
        items[i].outlier = true;
        items[i].outlierSide = 'high';
      } else {
        items[i].outlier = false;
        items[i].outlierSide = null;
      }
    }
  }

  // Mark numeric-but-not-checkable (e.g. only 1 item in a day) for UI.
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.verifiedManual) continue;
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    if (!it.checkable) {
      it.outlier = false;
      it.outlierSide = null;
    }
  }

  items.sort(
    (a, b) =>
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  const mrn = normalizeMrnQuery(params.mrn);
  let filtered = items;
  if (mrn) {
    const bucketStarts = filterBucketStartsByMrn(items, mrn);
    filtered = bucketStarts.size ? items.filter((it) => it.bucketStart && bucketStarts.has(it.bucketStart)) : [];
  }

  return {
    range,
    key: params.key,
    items: filtered.map((it) => ({
      rowId: it.rowId,
      data_mrn: it.data_mrn,
      odbiorca: it.odbiorca,
      numer_mrn: it.numer_mrn,
      coef: it.coef,
      verifiedManual: it.verifiedManual,
      checkable: it.checkable,
      outlier: it.outlier,
      outlierSide: it.outlierSide,
    })),
  };
}

export type ValidationDayFilter = 'all' | 'outliersHigh' | 'outliersLow' | 'singles';

export async function setValidationManualVerified(params: { rowId: number; verified: boolean }): Promise<{ ok: true }> {
  const client = await getPrisma();
  const id = Number(params.rowId);
  if (!Number.isFinite(id) || id <= 0) return { ok: true };

  if (params.verified) {
    await client.$executeRawUnsafe(
      `
        INSERT OR REPLACE INTO "validation_manual" ("raportRowId", "verifiedAt")
        VALUES (?, CURRENT_TIMESTAMP);
      `,
      id,
    );
  } else {
    await client.$executeRawUnsafe(`DELETE FROM "validation_manual" WHERE "raportRowId" = ?;`, id);
  }

  return { ok: true };
}

function toValidationKey(r: {
  odbiorca: string | null;
  kraj_wysylki: string | null;
  warunki_dostawy: string | null;
  waluta: string | null;
  kurs_waluty: string | null;
  transport_na_granicy_rodzaj: string | null;
  kod_towaru: string | null;
}): ValidationGroupKey {
  return {
    odbiorca: (r.odbiorca ?? '').trim(),
    kraj_wysylki: (r.kraj_wysylki ?? '').trim(),
    warunki_dostawy: (r.warunki_dostawy ?? '').trim(),
    waluta: (r.waluta ?? '').trim(),
    kurs_waluty: (r.kurs_waluty ?? '').trim(),
    transport_na_granicy_rodzaj: (r.transport_na_granicy_rodzaj ?? '').trim(),
    kod_towaru: (r.kod_towaru ?? '').trim(),
  };
}

type ValidationComputedItem = {
  rowId: number;
  data_mrn: string | null;
  bucketStart: string | null;
  numer_mrn: string | null;
  odbiorca: string | null;
  key: ValidationGroupKey;
  coef: number | null;
  verifiedManual: boolean;
  checkable: boolean;
  outlier: boolean;
  outlierSide: 'low' | 'high' | null;
};

function computeIqrFlags(items: ValidationComputedItem[]): void {
  const groupDay = new Map<string, number[]>();
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.bucketStart) continue;
    if (it.verifiedManual) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const k = `${JSON.stringify(it.key)}|${it.bucketStart}`;
    const arr = groupDay.get(k);
    if (arr) arr.push(i);
    else groupDay.set(k, [i]);
  }

  for (const indices of groupDay.values()) {
    if (indices.length < 2) continue;
    const valuesAsc = indices
      .map((i) => items[i].coef as number)
      .filter((v) => Number.isFinite(v))
      .sort((a, b) => a - b);
    if (valuesAsc.length < 2) continue;

    const q1 = quantileSorted(valuesAsc, 0.25);
    const q3 = quantileSorted(valuesAsc, 0.75);
    const iqr = q3 - q1;
    if (!Number.isFinite(iqr)) continue;
    const lower = q1 - 1.5 * iqr;
    const upper = q3 + 1.5 * iqr;

    for (const i of indices) {
      const v = items[i].coef;
      if (v == null || !Number.isFinite(v)) continue;
      items[i].checkable = true;
      if (v < lower) {
        items[i].outlier = true;
        items[i].outlierSide = 'low';
      } else if (v > upper) {
        items[i].outlier = true;
        items[i].outlierSide = 'high';
      } else {
        items[i].outlier = false;
        items[i].outlierSide = null;
      }
    }
  }

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.verifiedManual) {
      it.checkable = false;
      it.outlier = false;
      it.outlierSide = null;
      continue;
    }
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    if (!it.checkable) {
      it.outlier = false;
      it.outlierSide = null;
    }
  }
}

export async function getValidationDashboard(params: { month: string; mrn?: string | null; grouping?: unknown }): Promise<{
  range: { start: string; end: string };
  stats: { outliersHigh: number; outliersLow: number; singles: number; verifiedManual: number };
  days: Array<{ date: string; end: string; outliersHigh: number; outliersLow: number; singles: number; total: number }>;
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);

  const items: ValidationComputedItem[] = rows.map((r) => {
    const fees = parseLocaleNumber(r.oplaty_celne_razem);
    const mass = parseLocaleNumber(r.masa_netto);
    const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
    const rowIdRaw = Number(r.id);
    const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
    const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
    const data_mrn = r.data_mrn ?? null;
    return {
      rowId,
      data_mrn,
      bucketStart: data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null,
      numer_mrn: r.numer_mrn ?? null,
      odbiorca: r.odbiorca ?? null,
      key: toValidationKey(r),
      coef,
      verifiedManual,
      checkable: false,
      outlier: false,
      outlierSide: null as 'low' | 'high' | null,
    };
  });

  computeIqrFlags(items);

  const mrn = normalizeMrnQuery(params.mrn);
  const base = mrn ? filterToMrnCohorts(items, mrn) : items;

  const dayMap = new Map<string, { date: string; end: string; outliersHigh: number; outliersLow: number; singles: number; total: number }>();
  let verifiedManual = 0;
  for (const it of base) {
    if (it.verifiedManual) verifiedManual += 1;
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const date = it.bucketStart;
    const end = bucketEndFromStart(date, grouping) ?? date;
    const agg =
      dayMap.get(date) ?? { date, end, outliersHigh: 0, outliersLow: 0, singles: 0, total: 0 };
    if (!it.verifiedManual) agg.total += 1;
    if (it.outlierSide === 'high') agg.outliersHigh += 1;
    if (it.outlierSide === 'low') agg.outliersLow += 1;
    if (!it.verifiedManual && !it.checkable) agg.singles += 1;
    dayMap.set(date, agg);
  }

  const days = Array.from(dayMap.values()).sort((a, b) => a.date.localeCompare(b.date));
  const stats = {
    outliersHigh: days.reduce((s, d) => s + d.outliersHigh, 0),
    outliersLow: days.reduce((s, d) => s + d.outliersLow, 0),
    singles: days.reduce((s, d) => s + d.singles, 0),
    verifiedManual,
  };

  return { range, stats, days };
}

export async function getValidationDayItems(params: {
  month: string;
  date: string;
  filter: ValidationDayFilter;
  mrn?: string | null;
  grouping?: unknown;
}): Promise<{
  date: string;
  totals: { all: number; outliersHigh: number; outliersLow: number; singles: number; verifiedManual: number };
  items: Array<{
    rowId: number;
    data_mrn: string | null;
    numer_mrn: string | null;
    odbiorca: string | null;
    key: ValidationGroupKey;
    coef: number | null;
    verifiedManual: boolean;
    checkable: boolean;
    outlier: boolean;
    outlierSide: 'low' | 'high' | null;
  }>;
}> {
  const client = await getPrisma();
  const manual = await getValidationManualSet(client);

  const periodRange = toYmdRange(params.month);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = periodRange.start;

  const rawDate = String(params.date ?? '').trim();
  const bucketStart = bucketStartForDate(rawDate, anchor, grouping) ?? rawDate;
  const bucketEnd = bucketEndFromStart(bucketStart, grouping) ?? bucketStart;

  const queryStart = bucketStart.localeCompare(periodRange.start) < 0 ? periodRange.start : bucketStart;
  const queryEnd = bucketEnd.localeCompare(periodRange.end) > 0 ? periodRange.end : bucketEnd;
  const rows = queryStart.localeCompare(queryEnd) <= 0 ? await queryValidationRepresentativeRows(client, { start: queryStart, end: queryEnd }) : [];
  const items: ValidationComputedItem[] = rows.map((r) => {
    const fees = parseLocaleNumber(r.oplaty_celne_razem);
    const mass = parseLocaleNumber(r.masa_netto);
    const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
    const rowIdRaw = Number(r.id);
    const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
    const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
    const data_mrn = r.data_mrn ?? null;
    return {
      rowId,
      data_mrn,
      bucketStart: data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null,
      numer_mrn: r.numer_mrn ?? null,
      odbiorca: r.odbiorca ?? null,
      key: toValidationKey(r),
      coef,
      verifiedManual,
      checkable: false,
      outlier: false,
      outlierSide: null as 'low' | 'high' | null,
    };
  });

  computeIqrFlags(items);

  const mrn = normalizeMrnQuery(params.mrn);
  const inBucket = items.filter((it) => it.bucketStart === bucketStart);
  let base = inBucket;
  if (mrn) {
    const keys = new Set<string>();
    for (const it of inBucket) {
      if (!mrnContains(it.numer_mrn, mrn)) continue;
      keys.add(JSON.stringify(it.key));
    }
    base = keys.size ? inBucket.filter((it) => keys.has(JSON.stringify(it.key))) : [];
  }

  const filter: ValidationDayFilter =
    params.filter === 'outliersHigh' || params.filter === 'outliersLow' || params.filter === 'singles' ? params.filter : 'all';
  const totals = {
    all: base.length,
    outliersHigh: base.filter((it) => it.outlierSide === 'high').length,
    outliersLow: base.filter((it) => it.outlierSide === 'low').length,
    singles: base.filter((it) => !it.verifiedManual && it.coef != null && Number.isFinite(it.coef) && !it.checkable).length,
    verifiedManual: base.filter((it) => it.verifiedManual).length,
  };

  let filtered = base;
  if (filter === 'outliersHigh') filtered = base.filter((it) => it.outlierSide === 'high');
  else if (filter === 'outliersLow') filtered = base.filter((it) => it.outlierSide === 'low');
  else if (filter === 'singles')
    filtered = base.filter((it) => !it.verifiedManual && it.coef != null && Number.isFinite(it.coef) && !it.checkable);

  filtered.sort(
    (a, b) =>
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  return {
    date: bucketStart,
    totals,
    items: filtered.map((it) => ({
      rowId: it.rowId,
      data_mrn: it.data_mrn,
      numer_mrn: it.numer_mrn,
      odbiorca: it.odbiorca,
      key: it.key,
      coef: it.coef,
      verifiedManual: it.verifiedManual,
      checkable: it.checkable,
      outlier: it.outlier,
      outlierSide: it.outlierSide,
    })),
  };
}

export type ValidationOutlierError = {
  rowId: number;
  data_mrn: string | null;
  numer_mrn: string | null;
  nr_sad: string | null;
  agent_celny: string | null;
  odbiorca: string | null;
  key: ValidationGroupKey;
  coef: number;
  outlierSide: 'low' | 'high';
  limit: number;
  discrepancyPct: number | null;
};

function computeDiscrepancyPct(value: number, limit: number): number | null {
  const denom = Math.abs(limit);
  if (!Number.isFinite(denom) || denom < 1e-12) return null;
  const pct = (Math.abs(value - limit) / denom) * 100;
  return Number.isFinite(pct) ? pct : null;
}

export async function getValidationOutlierErrors(params: { month: string; mrn?: string | null; grouping?: unknown }): Promise<{
  range: { start: string; end: string };
  items: ValidationOutlierError[];
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);

  const computed: Array<
    ValidationComputedItem & {
      nr_sad: string | null;
      agent_celny: string | null;
    }
  > = rows.map((r) => {
    const fees = parseLocaleNumber(r.oplaty_celne_razem);
    const mass = parseLocaleNumber(r.masa_netto);
    const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
    const rowIdRaw = Number(r.id);
    const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
    const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
    const data_mrn = r.data_mrn ?? null;
    return {
      rowId,
      data_mrn,
      bucketStart: data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null,
      numer_mrn: r.numer_mrn ?? null,
      nr_sad: r.nr_sad ?? null,
      agent_celny: r.zglaszajacy ?? null,
      odbiorca: r.odbiorca ?? null,
      key: toValidationKey(r),
      coef,
      verifiedManual,
      checkable: false,
      outlier: false,
      outlierSide: null as 'low' | 'high' | null,
    };
  });

  const mrn = normalizeMrnQuery(params.mrn);
  const base = mrn ? filterToMrnCohorts(computed, mrn) : computed;

  // Compute outliers per (group key + day), and keep IQR bounds for discrepancy calculation.
  const groupDay = new Map<string, number[]>();
  for (let i = 0; i < base.length; i++) {
    const it = base[i];
    if (!it.bucketStart) continue;
    if (it.verifiedManual) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const k = `${JSON.stringify(it.key)}|${it.bucketStart}`;
    const arr = groupDay.get(k);
    if (arr) arr.push(i);
    else groupDay.set(k, [i]);
  }

  const bounds = new Map<string, { lower: number; upper: number }>();
  for (const [k, indices] of groupDay.entries()) {
    if (indices.length < 2) continue;
    const valuesAsc = indices
      .map((i) => base[i].coef as number)
      .filter((v) => Number.isFinite(v))
      .sort((a, b) => a - b);
    if (valuesAsc.length < 2) continue;

    const q1 = quantileSorted(valuesAsc, 0.25);
    const q3 = quantileSorted(valuesAsc, 0.75);
    const iqr = q3 - q1;
    if (!Number.isFinite(iqr)) continue;
    const lower = q1 - 1.5 * iqr;
    const upper = q3 + 1.5 * iqr;

    bounds.set(k, { lower, upper });

    for (const i of indices) {
      const v = base[i].coef;
      if (v == null || !Number.isFinite(v)) continue;
      base[i].checkable = true;
      if (v < lower) {
        base[i].outlier = true;
        base[i].outlierSide = 'low';
      } else if (v > upper) {
        base[i].outlier = true;
        base[i].outlierSide = 'high';
      } else {
        base[i].outlier = false;
        base[i].outlierSide = null;
      }
    }
  }

  for (let i = 0; i < base.length; i++) {
    const it = base[i];
    if (it.verifiedManual) continue;
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    if (!it.checkable) {
      it.outlier = false;
      it.outlierSide = null;
    }
  }

  const outliers: ValidationOutlierError[] = [];
  for (const it of base) {
    if (!it.data_mrn) continue;
    if (!it.bucketStart) continue;
    if (!it.outlierSide) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;

    const k = `${JSON.stringify(it.key)}|${it.bucketStart}`;
    const b = bounds.get(k);
    if (!b) continue;
    const limit = it.outlierSide === 'high' ? b.upper : b.lower;

    outliers.push({
      rowId: it.rowId,
      data_mrn: it.data_mrn,
      numer_mrn: it.numer_mrn,
      nr_sad: it.nr_sad,
      agent_celny: it.agent_celny,
      odbiorca: it.odbiorca,
      key: it.key,
      coef: it.coef,
      outlierSide: it.outlierSide,
      limit,
      discrepancyPct: computeDiscrepancyPct(it.coef, limit),
    });
  }

  outliers.sort(
    (a, b) =>
      String(a.agent_celny ?? '').localeCompare(String(b.agent_celny ?? '')) ||
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  return { range, items: outliers };
}
