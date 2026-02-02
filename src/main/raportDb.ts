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

  await client.$executeRawUnsafe(createMeta);
  await client.$executeRawUnsafe(createRows);
  await client.$executeRawUnsafe(createMrnBatch);

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

function toYmdRange(month: string): { start: string; end: string } {
  const m = String(month ?? '').trim();
  const match = /^(\d{4})-(\d{2})$/.exec(m);
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

export type ValidationGroupKey = {
  odbiorca: string;
  kraj_wysylki: string;
  warunki_dostawy: string;
  waluta: string;
  kurs_waluty: string;
  transport_na_granicy_rodzaj: string;
  kod_towaru: string;
};

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
  return (await client.$queryRawUnsafe(
    `
      WITH rep AS (
        SELECT
          "id",
          "data_mrn",
          "numer_mrn",
          "nr_sad",
          "odbiorca",
          "kraj_wysylki",
          "warunki_dostawy",
          "waluta",
          "kurs_waluty",
          "transport_na_granicy_rodzaj",
          "kod_towaru",
          "oplaty_celne_razem",
          "masa_netto",
          ROW_NUMBER() OVER (
            PARTITION BY TRIM(COALESCE("numer_mrn", '')), TRIM(COALESCE("nr_sad", ''))
            ORDER BY
              CASE
                WHEN "wartosc_pozycji" IS NOT NULL AND TRIM("wartosc_pozycji") <> '' THEN 0
                WHEN "wartosc_faktury" IS NOT NULL AND TRIM("wartosc_faktury") <> '' THEN 1
                ELSE 2
              END,
              "id" ASC
          ) as "rn"
        FROM "raport_rows"
        WHERE "data_mrn" IS NOT NULL
          AND TRIM("data_mrn") <> ''
          AND TRIM("data_mrn") BETWEEN ? AND ?
          AND "numer_mrn" IS NOT NULL AND TRIM("numer_mrn") <> ''
          AND "nr_sad" IS NOT NULL AND TRIM("nr_sad") <> ''
      )
      SELECT *
      FROM rep
      WHERE "rn" = 1;
    `,
    range.start,
    range.end,
  )) as Array<{
    id: number;
    data_mrn: string | null;
    numer_mrn: string | null;
    nr_sad: string | null;
    odbiorca: string | null;
    kraj_wysylki: string | null;
    warunki_dostawy: string | null;
    waluta: string | null;
    kurs_waluty: string | null;
    transport_na_granicy_rodzaj: string | null;
    kod_towaru: string | null;
    oplaty_celne_razem: string | null;
    masa_netto: string | null;
    rn: number;
  }>;
}

export async function getValidationGroups(params: { month: string }): Promise<{
  range: { start: string; end: string };
  groups: Array<{ key: ValidationGroupKey; count: number }>;
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);

  const rows = await queryValidationRepresentativeRows(client, range);

  const map = new Map<string, { key: ValidationGroupKey; count: number }>();
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

  const groups = Array.from(map.values()).sort((a, b) => b.count - a.count);
  return { range, groups };
}

export async function getValidationItems(params: { month: string; key: ValidationGroupKey }): Promise<{
  range: { start: string; end: string };
  key: ValidationGroupKey;
  items: Array<{
    data_mrn: string | null;
    odbiorca: string | null;
    numer_mrn: string | null;
    coef: number | null;
    outlier: boolean;
    outlierSide: 'low' | 'high' | null;
  }>;
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);
  const rows = await queryValidationRepresentativeRows(client, range);

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
      return {
        data_mrn: r.data_mrn ?? null,
        odbiorca: r.odbiorca ?? null,
        numer_mrn: r.numer_mrn ?? null,
        coef,
        outlier: false,
        outlierSide: null as 'low' | 'high' | null,
      };
    });

  // Outliers per day (IQR). If a day has <2 numeric values, it's ignored.
  const byDate = new Map<string, number[]>();
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.data_mrn) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const arr = byDate.get(it.data_mrn);
    if (arr) arr.push(i);
    else byDate.set(it.data_mrn, [i]);
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

  items.sort(
    (a, b) =>
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  return { range, key: params.key, items };
}
