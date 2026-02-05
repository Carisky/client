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

type AgentDzialInfo = {
  filePath: string;
  exists: boolean;
  rowCount: number;
  modifiedAt: string | null;
  error?: string;
};

function getAgentDzialMapUserFilePath(): string {
  return path.join(app.getPath('userData'), 'additional-data-maps.json');
}

function getBundledAgentDzialMapCandidates(): string[] {
  const candidates: string[] = [];
  const names = ['additional-data-maps.json', 'agent-dzial-map.json']; // legacy fallback
  for (const name of names) {
    // Dev: repo-root/client
    candidates.push(path.resolve(process.cwd(), 'resources', name));
    // Dev/webpack: resolve from compiled main folder
    candidates.push(path.resolve(__dirname, '..', '..', 'resources', name));
    // Packaged: extraResource copied under resources/resources/
    candidates.push(path.resolve(process.resourcesPath, 'resources', name));
    // Packaged fallback
    candidates.push(path.resolve(process.resourcesPath, name));
  }
  return candidates;
}

function ensureAgentDzialMapFile(): string {
  const userFile = getAgentDzialMapUserFilePath();
  const legacyUserFile = path.join(app.getPath('userData'), 'agent-dzial-map.json');
  try {
    if (fs.existsSync(userFile)) return userFile;
  } catch {
    // ignore
  }

  try {
    if (fs.existsSync(legacyUserFile)) {
      try {
        const legacyText = fs.readFileSync(legacyUserFile, 'utf8');
        const legacyParsed = JSON.parse(legacyText || '{}');
        if (legacyParsed && typeof legacyParsed === 'object' && !Array.isArray(legacyParsed)) {
          const departments = Object.entries(legacyParsed as Record<string, unknown>)
            .map(([FullName, department]) => ({ FullName, department }))
            .filter((r) => String(r.FullName ?? '').trim() && String(r.department ?? '').trim());
          fs.mkdirSync(path.dirname(userFile), { recursive: true });
          fs.writeFileSync(userFile, JSON.stringify({ departments }, null, 2) + '\n', 'utf8');
          return userFile;
        }
      } catch {
        // ignore and fallback to copy-as-is
      }

      try {
        fs.mkdirSync(path.dirname(userFile), { recursive: true });
        fs.copyFileSync(legacyUserFile, userFile);
        return userFile;
      } catch {
        // ignore
      }
    }
  } catch {
    // ignore
  }

  try {
    fs.mkdirSync(path.dirname(userFile), { recursive: true });
  } catch {
    // ignore
  }

  for (const c of getBundledAgentDzialMapCandidates()) {
    try {
      if (!fs.existsSync(c)) continue;
      fs.copyFileSync(c, userFile);
      return userFile;
    } catch {
      // ignore
    }
  }

  try {
    fs.writeFileSync(userFile, JSON.stringify({ departments: [] }, null, 2) + '\n', 'utf8');
  } catch {
    // ignore
  }
  return userFile;
}

function normalizeAgentKey(value: unknown): string {
  return String(value ?? '').trim().replace(/\s+/g, ' ').toLowerCase();
}

function loadAgentDzialMap(): { map: Map<string, string>; info: AgentDzialInfo } {
  const filePath = ensureAgentDzialMapFile();
  const info: AgentDzialInfo = { filePath, exists: false, rowCount: 0, modifiedAt: null };

  let text = '';
  try {
    info.exists = fs.existsSync(filePath);
    if (!info.exists) return { map: new Map(), info };
    const stat = fs.statSync(filePath);
    info.modifiedAt = stat.mtime ? stat.mtime.toISOString() : null;
    text = fs.readFileSync(filePath, 'utf8');
  } catch (e) {
    info.error = e instanceof Error ? e.message : String(e);
    return { map: new Map(), info };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(text || '{}');
  } catch {
    info.error = 'Niepoprawny JSON (additional-data-maps.json)';
    return { map: new Map(), info };
  }

  const out = new Map<string, string>();

  const addPair = (agentRaw: unknown, dzialRaw: unknown) => {
    const agent = normalizeAgentKey(agentRaw);
    const dzial = String(dzialRaw ?? '').trim();
    if (!agent || !dzial) return;
    out.set(agent, dzial);
  };

  const parseListItem = (item: unknown) => {
    if (Array.isArray(item) && item.length >= 2) {
      addPair(item[0], item[1]);
      return;
    }

    if (!item || typeof item !== 'object' || Array.isArray(item)) return;
    const rec = item as Record<string, unknown>;

    // Single-pair object: { "Full Name": "Department" }
    const entries = Object.entries(rec);
    if (entries.length === 1) {
      addPair(entries[0][0], entries[0][1]);
      return;
    }

    // Row object: { FullName: "...", department/departnent/oddzial/dzial/wydzial: "..." }
    const name =
      rec.FullName ??
      rec.fullName ??
      rec.fullname ??
      rec.name ??
      rec.agent ??
      rec.full_name ??
      null;
    const dept =
      rec.department ??
      rec.departnent ??
      rec.Department ??
      rec.Departnent ??
      rec.oddzial ??
      rec.Oddzial ??
      rec.dzial ??
      rec.Dzial ??
      rec.wydzial ??
      rec.Wydzial ??
      null;
    addPair(name, dept);
  };

  if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
    const rec = parsed as Record<string, unknown>;
    const departments = rec.departments;
    if (Array.isArray(departments)) {
      for (const item of departments) parseListItem(item);
    } else {
      // Legacy format: { "Full Name": "Department", ... }
      for (const [k, v] of Object.entries(rec)) {
        addPair(k, v);
      }
    }
  } else if (Array.isArray(parsed)) {
    for (const item of parsed) parseListItem(item);
  } else {
    info.error = 'Niepoprawny format słownika (oczekiwano obiektu lub tablicy wpisów)';
    return { map: new Map(), info };
  }
  info.rowCount = out.size;
  return { map: out, info };
}

export async function getAgentDzialInfo(): Promise<AgentDzialInfo> {
  return loadAgentDzialMap().info;
}

export async function clearAgentDzialMap(): Promise<AgentDzialInfo> {
  const filePath = ensureAgentDzialMapFile();
  try {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    let parsed: unknown = null;
    try {
      parsed = JSON.parse(fs.readFileSync(filePath, 'utf8') || '{}');
    } catch {
      parsed = null;
    }

    if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
      const obj = parsed as Record<string, unknown>;
      obj.departments = [];
      fs.writeFileSync(filePath, JSON.stringify(obj, null, 2) + '\n', 'utf8');
    } else {
      fs.writeFileSync(filePath, JSON.stringify({ departments: [] }, null, 2) + '\n', 'utf8');
    }
  } catch {
    // ignore
  }
  return loadAgentDzialMap().info;
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
  const { map: agentDzialMap } = loadAgentDzialMap();

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
      if (value !== null) {
        if (col.field === 'zglaszajacy') record[col.field] = normalizeAgentKey(value);
        else record[col.field] = value;
      }
    }

    const agentKey = normalizeAgentKey(record.zglaszajacy);
    const dzial = agentKey ? agentDzialMap.get(agentKey) : null;
    if (dzial) record.oddzial = dzial;

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

function normalizeTextQuery(value: unknown): string {
  if (typeof value !== 'string') return '';
  return value.trim().toLowerCase();
}

function textContains(field: string | null | undefined, query: string): boolean {
  if (!query) return true;
  const v = String(field ?? '').trim().toLowerCase();
  return v.includes(query);
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
  availableAgents: string[];
  items: ValidationOutlierError[];
}> {
  const client = await getPrisma();
  const range = toYmdRange(params.month);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);

  const availableAgentMap = new Map<string, string>();
  for (const r of rows) {
    const v = String(r.zglaszajacy ?? '').trim();
    if (!v) continue;
    const k = normalizeAgentKey(v);
    if (!k || availableAgentMap.has(k)) continue;
    availableAgentMap.set(k, v);
  }
  const availableAgents = Array.from(availableAgentMap.values()).sort((a, b) =>
    a.localeCompare(b),
  );

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

  return { range, availableAgents, items: outliers };
}

type ValidationIqrBounds = {
  q1: number;
  q3: number;
  iqr: number;
  lower: number;
  upper: number;
};

function computeIqrBoundsAndFlags(items: Array<ValidationComputedItem & { _boundsKey: string }>): Map<string, ValidationIqrBounds> {
  const groupDay = new Map<string, number[]>();
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!it.bucketStart) continue;
    if (it.verifiedManual) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const k = it._boundsKey;
    const arr = groupDay.get(k);
    if (arr) arr.push(i);
    else groupDay.set(k, [i]);
  }

  const bounds = new Map<string, ValidationIqrBounds>();
  for (const [k, indices] of groupDay.entries()) {
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

    bounds.set(k, { q1, q3, iqr, lower, upper });

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

  return bounds;
}

function a1(row1: number, col0 = 0): string {
  return xlsx.utils.encode_cell({ r: Math.max(0, row1 - 1), c: Math.max(0, col0) });
}

function addSectionTitle(ws: xlsx.WorkSheet, title: string, row1: number): number {
  xlsx.utils.sheet_add_aoa(ws, [[title]], { origin: a1(row1, 0) });
  return row1 + 1;
}

function addKeyValueMeta(ws: xlsx.WorkSheet, meta: Array<[string, string | number | null | undefined]>, row1: number): number {
  const rows = meta.map(([k, v]) => [k, v == null ? '' : v]);
  xlsx.utils.sheet_add_aoa(ws, rows, { origin: a1(row1, 0) });
  return row1 + rows.length;
}

function addJsonTable(ws: xlsx.WorkSheet, rows: Array<Record<string, unknown>>, row1: number): number {
  if (!rows.length) {
    xlsx.utils.sheet_add_aoa(ws, [['Brak danych']], { origin: a1(row1, 0) });
    return row1 + 1;
  }
  xlsx.utils.sheet_add_json(ws, rows, { origin: a1(row1, 0), skipHeader: false });
  return row1 + rows.length + 1;
}

type ValidationExportLayout = 'grouped' | 'separate';
type ValidationExportContent = 'full' | 'summary' | 'errors';
type ValidationExportColumns = 'full' | 'compact';

type NormalizedValidationExportOptions = {
  layout: ValidationExportLayout;
  content: ValidationExportContent;
  columns: ValidationExportColumns;
};

function normalizeValidationExportOptions(raw: unknown): NormalizedValidationExportOptions {
  const r = (raw ?? {}) as Record<string, unknown>;
  const layout: ValidationExportLayout = r.layout === 'grouped' || r.layout === 'separate' ? r.layout : 'separate';
  const content: ValidationExportContent = r.content === 'full' || r.content === 'summary' || r.content === 'errors' ? r.content : 'full';
  const columns: ValidationExportColumns = r.columns === 'full' || r.columns === 'compact' ? r.columns : 'full';
  return { layout, content, columns };
}

type ExportTableId = 'osobySummary' | 'osobyErrors' | 'dniSummary' | 'dniItems' | 'grupySummary' | 'grupyItems';

const EXPORT_TABLE_TITLES: Record<ExportTableId, { title: string; sheet: string; description: string }> = {
  osobySummary: { title: 'Osoby (podsumowanie)', sheet: 'Osoby_summary', description: 'Podsumowanie bledow / odchylen per agent' },
  osobyErrors: { title: 'Osoby (lista bledow / odchylen)', sheet: 'Osoby_bledy', description: 'Lista pozycji z odchyleniem (outliery IQR)' },
  dniSummary: { title: 'Dni (podsumowanie)', sheet: 'Dni_summary', description: 'Podsumowanie per dzien / bucket' },
  dniItems: { title: 'Dni (pozycje)', sheet: 'Dni_pozycje', description: 'Wszystkie pozycje z obliczeniami IQR per bucket' },
  grupySummary: { title: 'Grupy (podsumowanie)', sheet: 'Grupy_summary', description: 'Podsumowanie grup (klucz: odbiorca/kod/waluta...)' },
  grupyItems: { title: 'Grupy (pozycje)', sheet: 'Grupy_pozycje', description: 'Pozycje z obliczeniami IQR w rozbiciu na grupy' },
};

const EXPORT_HEADERS: Record<ExportTableId, { full: string[]; compact: string[] }> = {
  osobySummary: { full: ['Agent', 'Dzial', 'Errors', 'High', 'Low'], compact: ['Agent', 'Dzial', 'Errors', 'High', 'Low'] },
  osobyErrors: {
    full: [
      'Agent',
      'Dzial',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Odbiorca',
      'KodTowaru',
      'Waluta',
      'KursWaluty',
      'WarunkiDostawy',
      'KrajWysylki',
      'TransportRodzaj',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'Side',
      'Limit',
      'Q1',
      'Q3',
      'IQR',
      'Lower',
      'Upper',
      'DiscrepancyPct',
    ],
    compact: [
      'Agent',
      'Dzial',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Odbiorca',
      'KodTowaru',
      'Waluta',
      'KursWaluty',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'Side',
      'Limit',
      'DiscrepancyPct',
    ],
  },
  dniSummary: {
    full: ['DateStart', 'DateEnd', 'Total', 'OutliersHigh', 'OutliersLow', 'Singles'],
    compact: ['DateStart', 'DateEnd', 'Total', 'OutliersHigh', 'OutliersLow', 'Singles'],
  },
  dniItems: {
    full: [
      'BucketStart',
      'BucketEnd',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Agent',
      'Dzial',
      'Odbiorca',
      'KodTowaru',
      'Waluta',
      'KursWaluty',
      'WarunkiDostawy',
      'KrajWysylki',
      'TransportRodzaj',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'ManualVerified',
      'Checkable',
      'OutlierSide',
      'Limit',
      'Q1',
      'Q3',
      'IQR',
      'Lower',
      'Upper',
      'DiscrepancyPct',
      'RowId',
    ],
    compact: [
      'BucketStart',
      'BucketEnd',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Agent',
      'Dzial',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'ManualVerified',
      'OutlierSide',
      'Limit',
      'DiscrepancyPct',
      'RowId',
    ],
  },
  grupySummary: {
    full: ['Count', 'Odbiorca', 'KodTowaru', 'Waluta', 'KursWaluty', 'WarunkiDostawy', 'KrajWysylki', 'TransportRodzaj'],
    compact: ['Count', 'Odbiorca', 'KodTowaru', 'Waluta', 'KursWaluty', 'WarunkiDostawy', 'KrajWysylki', 'TransportRodzaj'],
  },
  grupyItems: {
    full: [
      'Odbiorca',
      'KodTowaru',
      'Waluta',
      'KursWaluty',
      'WarunkiDostawy',
      'KrajWysylki',
      'TransportRodzaj',
      'BucketStart',
      'BucketEnd',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Agent',
      'Dzial',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'ManualVerified',
      'Checkable',
      'OutlierSide',
      'Limit',
      'Q1',
      'Q3',
      'IQR',
      'Lower',
      'Upper',
      'DiscrepancyPct',
      'RowId',
    ],
    compact: [
      'Odbiorca',
      'KodTowaru',
      'Waluta',
      'KursWaluty',
      'WarunkiDostawy',
      'KrajWysylki',
      'TransportRodzaj',
      'BucketStart',
      'BucketEnd',
      'DataMRN',
      'MRN',
      'NrSAD',
      'Agent',
      'Dzial',
      'OplatyCelneRazem',
      'MasaNetto',
      'Coef',
      'ManualVerified',
      'OutlierSide',
      'Limit',
      'DiscrepancyPct',
      'RowId',
    ],
  },
};

function projectRowsByHeaders(rows: Array<Record<string, unknown>>, headers: string[]): Array<Record<string, unknown>> {
  return rows.map((r) => {
    const out: Record<string, unknown> = {};
    for (const h of headers) out[h] = r[h];
    return out;
  });
}

function addJsonTableWithHeader(ws: xlsx.WorkSheet, rows: Array<Record<string, unknown>>, row1: number, headers: string[]): number {
  if (!rows.length) {
    xlsx.utils.sheet_add_aoa(ws, [['Brak danych']], { origin: a1(row1, 0) });
    return row1 + 1;
  }
  xlsx.utils.sheet_add_json(ws, rows, { origin: a1(row1, 0), header: headers, skipHeader: false });
  return row1 + rows.length + 1;
}

function computeCols(rows: Array<Record<string, unknown>>, headers: string[]): Array<{ wch: number }> {
  const maxLens = headers.map((h) => Math.max(6, String(h ?? '').length));
  const sample = rows.slice(0, 200);
  for (const r of sample) {
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      const v = r[h];
      if (v == null) continue;
      const s = typeof v === 'string' ? v : typeof v === 'number' ? String(v) : typeof v === 'boolean' ? (v ? 'true' : 'false') : String(v);
      maxLens[i] = Math.max(maxLens[i], s.length);
    }
  }
  return maxLens.map((n) => ({ wch: Math.max(10, Math.min(60, n + 2)) }));
}

function applyTableDefaults(ws: xlsx.WorkSheet, rows: Array<Record<string, unknown>>, headers: string[]): void {
  if (!rows.length) {
    ws['!cols'] = [{ wch: 18 }];
    return;
  }
  if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
  ws['!cols'] = computeCols(rows, headers);
}

function safeSheetName(raw: string, used: Set<string>): string {
  const cleaned = String(raw ?? '')
    .replace(/[\][:*?/\\[]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  const base = (cleaned || 'Sheet').slice(0, 31);
  if (!used.has(base)) {
    used.add(base);
    return base;
  }
  for (let i = 2; i < 1000; i++) {
    const suffix = ` ${i}`;
    const name = `${base.slice(0, Math.max(1, 31 - suffix.length))}${suffix}`;
    if (!used.has(name)) {
      used.add(name);
      return name;
    }
  }
  const fallback = `${base.slice(0, 28)}...`;
  used.add(fallback);
  return fallback;
}

type ValidationExportPreviewSection = {
  title: string;
  rows: Array<Record<string, unknown>>;
  totalRows: number;
  truncated: boolean;
};

type ValidationExportPreviewSheet = {
  name: string;
  sections: ValidationExportPreviewSection[];
};

export type ValidationExportPreview = {
  period: string;
  grouping: string;
  range: { start: string; end: string };
  availableAgents: string[];
  meta: Array<{ key: string; value: string }>;
  sheets: ValidationExportPreviewSheet[];
};

export async function previewValidationWynikiExport(params: {
  period: string;
  mrn?: string | null;
  grouping?: unknown;
  filters?: {
    importer?: string | null;
    agent?: string[] | string | null;
    dzial?: string | null;
  };
  exportOptions?: unknown;
  limit?: number;
}): Promise<ValidationExportPreview> {
  const limit = Number.isFinite(Number(params.limit)) ? Math.max(10, Math.min(500, Number(params.limit))) : 200;
  const exportOptionsRaw = normalizeValidationExportOptions(params.exportOptions);
  const exportOptions =
    exportOptionsRaw.layout === 'grouped' ? { ...exportOptionsRaw, content: 'full' as const, columns: 'full' as const } : exportOptionsRaw;

  const client = await getPrisma();
  const range = toYmdRange(params.period);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);
  const { map: agentDzialMap, info: agentDzialInfo } = loadAgentDzialMap();

  const availableAgentMap = new Map<string, string>();
  for (const r of rows) {
    const v = String(r.zglaszajacy ?? '').trim();
    if (!v) continue;
    const k = normalizeAgentKey(v);
    if (!k || availableAgentMap.has(k)) continue;
    availableAgentMap.set(k, v);
  }
  const availableAgents = Array.from(availableAgentMap.values()).sort((a, b) =>
    a.localeCompare(b),
  );

  const filters = params.filters ?? {};
  const importerQ = normalizeTextQuery(filters.importer);
  const dzialQ = normalizeTextQuery(filters.dzial);
  const agentRaw = filters.agent;
  const agentList = Array.isArray(agentRaw)
    ? agentRaw
    : typeof agentRaw === 'string'
      ? [agentRaw]
      : [];
  const agentKeys = new Set(
    agentList.map((a) => normalizeAgentKey(a)).filter(Boolean),
  );

  const filteredRows = rows.filter((r) => {
    if (importerQ && !textContains(r.odbiorca, importerQ)) return false;
    if (agentKeys.size) {
      const ak = normalizeAgentKey(r.zglaszajacy);
      if (!ak || !agentKeys.has(ak)) return false;
    }

    if (dzialQ) {
      const agentKey = normalizeAgentKey(r.zglaszajacy);
      const dz = (agentKey ? agentDzialMap.get(agentKey) : null) ?? '';
      if (!dz || !textContains(dz, dzialQ)) return false;
    }

    return true;
  });

  type ExportItem = ValidationComputedItem & {
    nr_sad: string | null;
    agent_celny: string | null;
    dzial: string | null;
    fees: number | null;
    mass: number | null;
    bucketEnd: string | null;
    _boundsKey: string;
  };

  const items: ExportItem[] = filteredRows.map((r) => {
    const fees = parseLocaleNumber(r.oplaty_celne_razem);
    const mass = parseLocaleNumber(r.masa_netto);
    const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
    const rowIdRaw = Number(r.id);
    const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
    const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
    const data_mrn = r.data_mrn ?? null;
    const bucketStart = data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null;
    const key = toValidationKey(r);
    const boundsKey = bucketStart ? `${JSON.stringify(key)}|${bucketStart}` : `${JSON.stringify(key)}|`;
    const agentKey = normalizeAgentKey(r.zglaszajacy);
    const dzial = (agentKey ? agentDzialMap.get(agentKey) : null) ?? null;
    return {
      rowId,
      data_mrn,
      bucketStart,
      bucketEnd: bucketStart ? bucketEndFromStart(bucketStart, grouping) : null,
      numer_mrn: r.numer_mrn ?? null,
      nr_sad: r.nr_sad ?? null,
      agent_celny: r.zglaszajacy ?? null,
      dzial,
      odbiorca: r.odbiorca ?? null,
      key,
      fees,
      mass,
      coef,
      verifiedManual,
      checkable: false,
      outlier: false,
      outlierSide: null as 'low' | 'high' | null,
      _boundsKey: boundsKey,
    };
  });

  const bounds = computeIqrBoundsAndFlags(items);

  const mrnNorm = normalizeMrnQuery(params.mrn);
  const base = mrnNorm ? filterToMrnCohorts(items, mrnNorm) : items;

  const dayMap = new Map<
    string,
    { date: string; end: string; outliersHigh: number; outliersLow: number; singles: number; total: number }
  >();
  let verifiedManual = 0;
  for (const it of base) {
    if (it.verifiedManual) verifiedManual += 1;
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const date = it.bucketStart;
    const end = it.bucketEnd ?? date;
    const agg = dayMap.get(date) ?? { date, end, outliersHigh: 0, outliersLow: 0, singles: 0, total: 0 };
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

  const groupMap = new Map<string, { key: ValidationGroupKey; count: number }>();
  for (const it of base) {
    const k = JSON.stringify(it.key);
    const existing = groupMap.get(k);
    if (existing) existing.count += 1;
    else groupMap.set(k, { key: it.key, count: 1 });
  }
  const groups = Array.from(groupMap.values()).sort((a, b) => b.count - a.count);

  const outliers = base
    .filter((it) => it.outlierSide && it.coef != null && Number.isFinite(it.coef) && it.bucketStart)
    .map((it) => {
      const b = bounds.get(it._boundsKey);
      const limit = b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null;
      return {
        ...it,
        bound: b ?? null,
        limit,
        discrepancyPct: limit == null ? null : computeDiscrepancyPct(it.coef as number, limit),
      };
    });

  outliers.sort(
    (a, b) =>
      String(a.agent_celny ?? '').localeCompare(String(b.agent_celny ?? '')) ||
      (b.discrepancyPct ?? -1) - (a.discrepancyPct ?? -1) ||
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  const agentMap = new Map<string, { agent: string; dzial: string | null; total: number; high: number; low: number }>();
  for (const o of outliers) {
    const agent = String(o.agent_celny ?? '').trim() || 'вЂ”';
    const dzial = agentDzialMap.get(normalizeAgentKey(agent)) ?? null;
    const agg = agentMap.get(agent) ?? { agent, dzial, total: 0, high: 0, low: 0 };
    agg.total += 1;
    if (o.outlierSide === 'high') agg.high += 1;
    if (o.outlierSide === 'low') agg.low += 1;
    agentMap.set(agent, agg);
  }
  const agents = Array.from(agentMap.values()).sort((a, b) => b.total - a.total || a.agent.localeCompare(b.agent));

  const exportedAt = new Date().toISOString();
  const metaRows: Array<[string, string | number | null | undefined]> = [
    ['Raport', 'Wyniki (IQR)'],
    ['ExportedAt', exportedAt],
    ['Period', params.period],
    ['RangeStart', range.start],
    ['RangeEnd', range.end],
    ['Grouping', grouping],
    ['MRN filter', params.mrn ?? ''],
    ['Importer filter', filters.importer ?? ''],
    ['Agent filter', agentList.join(', ')],
    ['Dzial filter', filters.dzial ?? ''],
    ['Agent->Dzial map', `${agentDzialInfo.rowCount} (${agentDzialInfo.filePath})`],
    ['Coef formula', 'oplaty_celne_razem / masa_netto'],
    ['IQR rule', 'lower=Q1-1.5*IQR, upper=Q3+1.5*IQR'],
    ['OutliersHigh', stats.outliersHigh],
    ['OutliersLow', stats.outliersLow],
    ['Singles', stats.singles],
    ['ManualVerified', stats.verifiedManual],
    ['ItemsTotal', base.length],
  ];

  const osobySummary = agents.map((a) => ({ Agent: a.agent, Dzial: a.dzial ?? '', Errors: a.total, High: a.high, Low: a.low }));
  const osobyErrors = outliers.map((o) => {
    const b = o.bound as ValidationIqrBounds | null;
    const key = o.key;
    return {
      Agent: String(o.agent_celny ?? '').trim() || 'вЂ”',
      Dzial: o.dzial ?? '',
      DataMRN: o.data_mrn,
      MRN: o.numer_mrn,
      NrSAD: o.nr_sad,
      Odbiorca: key?.odbiorca ?? '',
      KodTowaru: key?.kod_towaru ?? '',
      Waluta: key?.waluta ?? '',
      KursWaluty: key?.kurs_waluty ?? '',
      WarunkiDostawy: key?.warunki_dostawy ?? '',
      KrajWysylki: key?.kraj_wysylki ?? '',
      TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
      OplatyCelneRazem: o.fees,
      MasaNetto: o.mass,
      Coef: o.coef,
      Side: o.outlierSide,
      Limit: o.limit,
      Q1: b?.q1 ?? null,
      Q3: b?.q3 ?? null,
      IQR: b?.iqr ?? null,
      Lower: b?.lower ?? null,
      Upper: b?.upper ?? null,
      DiscrepancyPct: o.discrepancyPct,
    };
  });

  const dniSummary = days.map((d) => ({
    DateStart: d.date,
    DateEnd: d.end,
    Total: d.total,
    OutliersHigh: d.outliersHigh,
    OutliersLow: d.outliersLow,
    Singles: d.singles,
  }));

  const dniItems = base
    .slice()
    .sort(
      (a, b) =>
        String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
        String(a.outlierSide ?? '').localeCompare(String(b.outlierSide ?? '')) ||
        String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
        String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
    )
    .map((it) => {
      const b = bounds.get(it._boundsKey) ?? null;
      const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
      const discrepancyPct =
        it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null ? computeDiscrepancyPct(it.coef, limit) : null;
      const key = it.key;
      return {
        BucketStart: it.bucketStart,
        BucketEnd: it.bucketEnd,
        DataMRN: it.data_mrn,
        MRN: it.numer_mrn,
        NrSAD: it.nr_sad,
        Agent: it.agent_celny,
        Dzial: it.dzial,
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        OplatyCelneRazem: it.fees,
        MasaNetto: it.mass,
        Coef: it.coef,
        ManualVerified: it.verifiedManual,
        Checkable: it.checkable,
        OutlierSide: it.outlierSide,
        Limit: limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: discrepancyPct,
        RowId: it.rowId,
      };
    });

  const grupySummary = groups.map((g) => ({
    Count: g.count,
    Odbiorca: g.key.odbiorca,
    KodTowaru: g.key.kod_towaru,
    Waluta: g.key.waluta,
    KursWaluty: g.key.kurs_waluty,
    WarunkiDostawy: g.key.warunki_dostawy,
    KrajWysylki: g.key.kraj_wysylki,
    TransportRodzaj: g.key.transport_na_granicy_rodzaj,
  }));

  const grupyItems = base
    .slice()
    .sort(
      (a, b) =>
        JSON.stringify(a.key).localeCompare(JSON.stringify(b.key)) ||
        String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
        String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
        String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
    )
    .map((it) => {
      const b = bounds.get(it._boundsKey) ?? null;
      const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
      const discrepancyPct =
        it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null ? computeDiscrepancyPct(it.coef, limit) : null;
      const key = it.key;
      return {
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        BucketStart: it.bucketStart,
        BucketEnd: it.bucketEnd,
        DataMRN: it.data_mrn,
        MRN: it.numer_mrn,
        NrSAD: it.nr_sad,
        Agent: it.agent_celny,
        Dzial: it.dzial,
        OplatyCelneRazem: it.fees,
        MasaNetto: it.mass,
        Coef: it.coef,
        ManualVerified: it.verifiedManual,
        Checkable: it.checkable,
        OutlierSide: it.outlierSide,
        Limit: limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: discrepancyPct,
        RowId: it.rowId,
      };
    });

  const take = (rows: Array<Record<string, unknown>>): { rows: Array<Record<string, unknown>>; totalRows: number; truncated: boolean } => {
    const totalRows = rows.length;
    if (totalRows <= limit) return { rows, totalRows, truncated: false };
    return { rows: rows.slice(0, limit), totalRows, truncated: true };
  };

  const meta = metaRows.map(([k, v]) => ({ key: String(k), value: v == null ? '' : String(v) }));

  const include: Record<ExportTableId, boolean> =
    exportOptions.content === 'summary'
      ? { osobySummary: true, osobyErrors: false, dniSummary: true, dniItems: false, grupySummary: true, grupyItems: false }
      : exportOptions.content === 'errors'
        ? { osobySummary: true, osobyErrors: true, dniSummary: false, dniItems: false, grupySummary: false, grupyItems: false }
        : { osobySummary: true, osobyErrors: true, dniSummary: true, dniItems: true, grupySummary: true, grupyItems: true };

  const preset = exportOptions.columns === 'compact' ? 'compact' : 'full';
  const tables: Record<ExportTableId, { rows: Array<Record<string, unknown>>; totalRows: number; truncated: boolean }> = {
    osobySummary: take(projectRowsByHeaders(osobySummary, EXPORT_HEADERS.osobySummary[preset])),
    osobyErrors: take(projectRowsByHeaders(osobyErrors, EXPORT_HEADERS.osobyErrors[preset])),
    dniSummary: take(projectRowsByHeaders(dniSummary, EXPORT_HEADERS.dniSummary[preset])),
    dniItems: take(projectRowsByHeaders(dniItems, EXPORT_HEADERS.dniItems[preset])),
    grupySummary: take(projectRowsByHeaders(grupySummary, EXPORT_HEADERS.grupySummary[preset])),
    grupyItems: take(projectRowsByHeaders(grupyItems, EXPORT_HEADERS.grupyItems[preset])),
  };

  const groupedSheets: ValidationExportPreviewSheet[] = [
    {
      name: 'Osoby',
      sections: [
        ...(include.osobySummary ? [{ title: EXPORT_TABLE_TITLES.osobySummary.title, ...tables.osobySummary }] : []),
        ...(include.osobyErrors ? [{ title: EXPORT_TABLE_TITLES.osobyErrors.title, ...tables.osobyErrors }] : []),
      ],
    },
    {
      name: 'Dni',
      sections: [
        ...(include.dniSummary ? [{ title: EXPORT_TABLE_TITLES.dniSummary.title, ...tables.dniSummary }] : []),
        ...(include.dniItems ? [{ title: EXPORT_TABLE_TITLES.dniItems.title, ...tables.dniItems }] : []),
      ],
    },
    {
      name: 'Grupy',
      sections: [
        ...(include.grupySummary ? [{ title: EXPORT_TABLE_TITLES.grupySummary.title, ...tables.grupySummary }] : []),
        ...(include.grupyItems ? [{ title: EXPORT_TABLE_TITLES.grupyItems.title, ...tables.grupyItems }] : []),
      ],
    },
  ].filter((s) => s.sections.length > 0);

  const separateSheets: ValidationExportPreviewSheet[] = (Object.keys(include) as ExportTableId[])
    .filter((id) => include[id])
    .map((id) => ({
      name: EXPORT_TABLE_TITLES[id].sheet,
      sections: [{ title: EXPORT_TABLE_TITLES[id].title, ...tables[id] }],
    }));

  return {
    period: params.period,
    grouping,
    range,
    availableAgents,
    meta,
    sheets: exportOptions.layout === 'grouped' ? groupedSheets : separateSheets,
  };
}

export async function exportValidationWynikiToXlsx(params: {
  period: string;
  mrn?: string | null;
  grouping?: unknown;
  filters?: {
    importer?: string | null;
    agent?: string[] | string | null;
    dzial?: string | null;
  };
  exportOptions?: unknown;
  filePath: string;
}): Promise<{ ok: true }> {
  const exportOptions = normalizeValidationExportOptions(params.exportOptions);
  const client = await getPrisma();
  const range = toYmdRange(params.period);
  const { grouping } = getValidationGroupingConfig(params);
  const anchor = range.start;
  const rows = await queryValidationRepresentativeRows(client, range);
  const manual = await getValidationManualSet(client);
  const { map: agentDzialMap, info: agentDzialInfo } = loadAgentDzialMap();

  const filters = params.filters ?? {};
  const importerQ = normalizeTextQuery(filters.importer);
  const dzialQ = normalizeTextQuery(filters.dzial);
  const agentRaw = filters.agent;
  const agentList = Array.isArray(agentRaw)
    ? agentRaw
    : typeof agentRaw === 'string'
      ? [agentRaw]
      : [];
  const agentKeys = new Set(
    agentList.map((a) => normalizeAgentKey(a)).filter(Boolean),
  );

  const filteredRows = rows.filter((r) => {
    if (importerQ && !textContains(r.odbiorca, importerQ)) return false;
    if (agentKeys.size) {
      const ak = normalizeAgentKey(r.zglaszajacy);
      if (!ak || !agentKeys.has(ak)) return false;
    }

    if (dzialQ) {
      const agentKey = normalizeAgentKey(r.zglaszajacy);
      const dz = (agentKey ? agentDzialMap.get(agentKey) : null) ?? '';
      if (!dz || !textContains(dz, dzialQ)) return false;
    }

    return true;
  });

  type ExportItem = ValidationComputedItem & {
    nr_sad: string | null;
    agent_celny: string | null;
    dzial: string | null;
    fees: number | null;
    mass: number | null;
    bucketEnd: string | null;
    _boundsKey: string;
  };

  const items: ExportItem[] = filteredRows.map((r) => {
    const fees = parseLocaleNumber(r.oplaty_celne_razem);
    const mass = parseLocaleNumber(r.masa_netto);
    const coef = fees != null && mass != null && mass !== 0 ? fees / mass : null;
    const rowIdRaw = Number(r.id);
    const rowId = Number.isFinite(rowIdRaw) ? rowIdRaw : 0;
    const verifiedManual = rowId > 0 ? manual.has(rowId) : false;
    const data_mrn = r.data_mrn ?? null;
    const bucketStart = data_mrn ? bucketStartForDate(data_mrn, anchor, grouping) : null;
    const key = toValidationKey(r);
    const boundsKey = bucketStart ? `${JSON.stringify(key)}|${bucketStart}` : `${JSON.stringify(key)}|`;
    const agentKey = normalizeAgentKey(r.zglaszajacy);
    const dzial = (agentKey ? agentDzialMap.get(agentKey) : null) ?? null;
    return {
      rowId,
      data_mrn,
      bucketStart,
      bucketEnd: bucketStart ? bucketEndFromStart(bucketStart, grouping) : null,
      numer_mrn: r.numer_mrn ?? null,
      nr_sad: r.nr_sad ?? null,
      agent_celny: r.zglaszajacy ?? null,
      dzial,
      odbiorca: r.odbiorca ?? null,
      key,
      fees,
      mass,
      coef,
      verifiedManual,
      checkable: false,
      outlier: false,
      outlierSide: null as 'low' | 'high' | null,
      _boundsKey: boundsKey,
    };
  });

  const bounds = computeIqrBoundsAndFlags(items);

  const mrnNorm = normalizeMrnQuery(params.mrn);
  const base = mrnNorm ? filterToMrnCohorts(items, mrnNorm) : items;

  const dayMap = new Map<
    string,
    { date: string; end: string; outliersHigh: number; outliersLow: number; singles: number; total: number }
  >();
  let verifiedManual = 0;
  for (const it of base) {
    if (it.verifiedManual) verifiedManual += 1;
    if (!it.bucketStart) continue;
    if (it.coef == null || !Number.isFinite(it.coef)) continue;
    const date = it.bucketStart;
    const end = it.bucketEnd ?? date;
    const agg = dayMap.get(date) ?? { date, end, outliersHigh: 0, outliersLow: 0, singles: 0, total: 0 };
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

  const groupMap = new Map<string, { key: ValidationGroupKey; count: number }>();
  for (const it of base) {
    const k = JSON.stringify(it.key);
    const existing = groupMap.get(k);
    if (existing) existing.count += 1;
    else groupMap.set(k, { key: it.key, count: 1 });
  }
  const groups = Array.from(groupMap.values()).sort((a, b) => b.count - a.count);

  const outliers = base
    .filter((it) => it.outlierSide && it.coef != null && Number.isFinite(it.coef) && it.bucketStart)
    .map((it) => {
      const b = bounds.get(it._boundsKey);
      const limit = b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null;
      return {
        ...it,
        bound: b ?? null,
        limit,
        discrepancyPct: limit == null ? null : computeDiscrepancyPct(it.coef as number, limit),
      };
    });

  outliers.sort(
    (a, b) =>
      String(a.agent_celny ?? '').localeCompare(String(b.agent_celny ?? '')) ||
      (b.discrepancyPct ?? -1) - (a.discrepancyPct ?? -1) ||
      String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
      String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
  );

  const agentMap = new Map<string, { agent: string; dzial: string | null; total: number; high: number; low: number }>();
  for (const o of outliers) {
    const agent = String(o.agent_celny ?? '').trim() || '—';
    const dzial = agentDzialMap.get(normalizeAgentKey(agent)) ?? null;
    const agg = agentMap.get(agent) ?? { agent, dzial, total: 0, high: 0, low: 0 };
    agg.total += 1;
    if (o.outlierSide === 'high') agg.high += 1;
    if (o.outlierSide === 'low') agg.low += 1;
    agentMap.set(agent, agg);
  }
  const agents = Array.from(agentMap.values()).sort((a, b) => b.total - a.total || a.agent.localeCompare(b.agent));

  const exportedAt = new Date().toISOString();
  const metaRows: Array<[string, string | number | null | undefined]> = [
    ['Raport', 'Wyniki (IQR)'],
    ['ExportedAt', exportedAt],
    ['Period', params.period],
    ['RangeStart', range.start],
    ['RangeEnd', range.end],
    ['Grouping', grouping],
    ['MRN filter', params.mrn ?? ''],
    ['Importer filter', filters.importer ?? ''],
    ['Agent filter', agentList.join(', ')],
    ['Dzial filter', filters.dzial ?? ''],
    ['Agent->Dzial map', `${agentDzialInfo.rowCount} (${agentDzialInfo.filePath})`],
    ['Coef formula', 'oplaty_celne_razem / masa_netto'],
    ['IQR rule', 'lower=Q1-1.5*IQR, upper=Q3+1.5*IQR'],
    ['OutliersHigh', stats.outliersHigh],
    ['OutliersLow', stats.outliersLow],
    ['Singles', stats.singles],
    ['ManualVerified', stats.verifiedManual],
    ['ItemsTotal', base.length],
  ];

  if (exportOptions.layout === 'separate') {
    const include: Record<ExportTableId, boolean> =
      exportOptions.content === 'summary'
        ? { osobySummary: true, osobyErrors: false, dniSummary: true, dniItems: false, grupySummary: true, grupyItems: false }
        : exportOptions.content === 'errors'
          ? { osobySummary: true, osobyErrors: true, dniSummary: false, dniItems: false, grupySummary: false, grupyItems: false }
          : { osobySummary: true, osobyErrors: true, dniSummary: true, dniItems: true, grupySummary: true, grupyItems: true };

    const preset = exportOptions.columns === 'compact' ? 'compact' : 'full';

    const osobySummary = agents.map((a) => ({ Agent: a.agent, Dzial: a.dzial ?? '', Errors: a.total, High: a.high, Low: a.low }));
    const osobyErrors = outliers.map((o) => {
      const b = o.bound as ValidationIqrBounds | null;
      const key = o.key;
      return {
        Agent: String(o.agent_celny ?? '').trim() || 'вЂ”',
        Dzial: o.dzial ?? '',
        DataMRN: o.data_mrn,
        MRN: o.numer_mrn,
        NrSAD: o.nr_sad,
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        OplatyCelneRazem: o.fees,
        MasaNetto: o.mass,
        Coef: o.coef,
        Side: o.outlierSide,
        Limit: o.limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: o.discrepancyPct,
      };
    });

    const dniSummary = days.map((d) => ({
      DateStart: d.date,
      DateEnd: d.end,
      Total: d.total,
      OutliersHigh: d.outliersHigh,
      OutliersLow: d.outliersLow,
      Singles: d.singles,
    }));

    const dniItems = base
      .slice()
      .sort(
        (a, b) =>
          String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
          String(a.outlierSide ?? '').localeCompare(String(b.outlierSide ?? '')) ||
          String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
          String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
      )
      .map((it) => {
        const b = bounds.get(it._boundsKey) ?? null;
        const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
        const discrepancyPct =
          it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null
            ? computeDiscrepancyPct(it.coef, limit)
            : null;
        const key = it.key;
        return {
          BucketStart: it.bucketStart,
          BucketEnd: it.bucketEnd,
          DataMRN: it.data_mrn,
          MRN: it.numer_mrn,
          NrSAD: it.nr_sad,
          Agent: it.agent_celny,
          Dzial: it.dzial,
          Odbiorca: key?.odbiorca ?? '',
          KodTowaru: key?.kod_towaru ?? '',
          Waluta: key?.waluta ?? '',
          KursWaluty: key?.kurs_waluty ?? '',
          WarunkiDostawy: key?.warunki_dostawy ?? '',
          KrajWysylki: key?.kraj_wysylki ?? '',
          TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
          OplatyCelneRazem: it.fees,
          MasaNetto: it.mass,
          Coef: it.coef,
          ManualVerified: it.verifiedManual,
          Checkable: it.checkable,
          OutlierSide: it.outlierSide,
          Limit: limit,
          Q1: b?.q1 ?? null,
          Q3: b?.q3 ?? null,
          IQR: b?.iqr ?? null,
          Lower: b?.lower ?? null,
          Upper: b?.upper ?? null,
          DiscrepancyPct: discrepancyPct,
          RowId: it.rowId,
        };
      });

    const grupySummary = groups.map((g) => ({
      Count: g.count,
      Odbiorca: g.key.odbiorca,
      KodTowaru: g.key.kod_towaru,
      Waluta: g.key.waluta,
      KursWaluty: g.key.kurs_waluty,
      WarunkiDostawy: g.key.warunki_dostawy,
      KrajWysylki: g.key.kraj_wysylki,
      TransportRodzaj: g.key.transport_na_granicy_rodzaj,
    }));

    const grupyItems = base
      .slice()
      .sort(
        (a, b) =>
          JSON.stringify(a.key).localeCompare(JSON.stringify(b.key)) ||
          String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
          String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
          String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
      )
      .map((it) => {
        const b = bounds.get(it._boundsKey) ?? null;
        const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
        const discrepancyPct =
          it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null
            ? computeDiscrepancyPct(it.coef, limit)
            : null;
        const key = it.key;
        return {
          Odbiorca: key?.odbiorca ?? '',
          KodTowaru: key?.kod_towaru ?? '',
          Waluta: key?.waluta ?? '',
          KursWaluty: key?.kurs_waluty ?? '',
          WarunkiDostawy: key?.warunki_dostawy ?? '',
          KrajWysylki: key?.kraj_wysylki ?? '',
          TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
          BucketStart: it.bucketStart,
          BucketEnd: it.bucketEnd,
          DataMRN: it.data_mrn,
          MRN: it.numer_mrn,
          NrSAD: it.nr_sad,
          Agent: it.agent_celny,
          Dzial: it.dzial,
          OplatyCelneRazem: it.fees,
          MasaNetto: it.mass,
          Coef: it.coef,
          ManualVerified: it.verifiedManual,
          Checkable: it.checkable,
          OutlierSide: it.outlierSide,
          Limit: limit,
          Q1: b?.q1 ?? null,
          Q3: b?.q3 ?? null,
          IQR: b?.iqr ?? null,
          Lower: b?.lower ?? null,
          Upper: b?.upper ?? null,
          DiscrepancyPct: discrepancyPct,
          RowId: it.rowId,
        };
      });

    const tableRows: Record<ExportTableId, Array<Record<string, unknown>>> = {
      osobySummary: projectRowsByHeaders(osobySummary, EXPORT_HEADERS.osobySummary[preset]),
      osobyErrors: projectRowsByHeaders(osobyErrors, EXPORT_HEADERS.osobyErrors[preset]),
      dniSummary: projectRowsByHeaders(dniSummary, EXPORT_HEADERS.dniSummary[preset]),
      dniItems: projectRowsByHeaders(dniItems, EXPORT_HEADERS.dniItems[preset]),
      grupySummary: projectRowsByHeaders(grupySummary, EXPORT_HEADERS.grupySummary[preset]),
      grupyItems: projectRowsByHeaders(grupyItems, EXPORT_HEADERS.grupyItems[preset]),
    };

    const usedSheetNames = new Set<string>();
    const metaSheetName = safeSheetName('Meta', usedSheetNames);

    const tableOrder: ExportTableId[] = ['osobySummary', 'osobyErrors', 'dniSummary', 'dniItems', 'grupySummary', 'grupyItems'];
    const plannedSheets = tableOrder
      .filter((id) => include[id])
      .map((id) => ({
        id,
        name: safeSheetName(EXPORT_TABLE_TITLES[id].sheet, usedSheetNames),
      }));

    const wb = xlsx.utils.book_new();

    const wsMeta = xlsx.utils.aoa_to_sheet([]);
    let r1 = 1;
    r1 = addSectionTitle(wsMeta, 'Meta', r1);
    r1 = addKeyValueMeta(wsMeta, metaRows, r1);
    r1 += 1;
    r1 = addSectionTitle(wsMeta, 'Arkusze', r1);
    const sheetIndex = plannedSheets.map((s) => ({
      Sheet: s.name,
      Title: EXPORT_TABLE_TITLES[s.id].title,
      Rows: tableRows[s.id].length,
      Description: EXPORT_TABLE_TITLES[s.id].description,
    }));
    addJsonTableWithHeader(wsMeta, sheetIndex, r1, ['Sheet', 'Title', 'Rows', 'Description']);
    wsMeta['!cols'] = [{ wch: 22 }, { wch: 44 }, { wch: 10 }, { wch: 60 }];
    xlsx.utils.book_append_sheet(wb, wsMeta, metaSheetName);

    for (const s of plannedSheets) {
      const id = s.id;
      const headers = EXPORT_HEADERS[id][preset];
      const rows = tableRows[id];
      const ws = xlsx.utils.aoa_to_sheet([]);
      addJsonTableWithHeader(ws, rows, 1, headers);
      applyTableDefaults(ws, rows, headers);
      xlsx.utils.book_append_sheet(wb, ws, s.name);
    }

    fs.mkdirSync(path.dirname(params.filePath), { recursive: true });
    xlsx.writeFile(wb, params.filePath, { bookType: 'xlsx', compression: true });
    return { ok: true };
  }

  const wb = xlsx.utils.book_new();

  // Osoby
  const wsOsoby = xlsx.utils.aoa_to_sheet([]);
  let r1 = 1;
  r1 = addKeyValueMeta(wsOsoby, metaRows, r1);
  r1 += 1;
  r1 = addSectionTitle(wsOsoby, 'Osoby (podsumowanie)', r1);
  r1 = addJsonTable(
    wsOsoby,
    agents.map((a) => ({ Agent: a.agent, Dzial: a.dzial ?? '', Errors: a.total, High: a.high, Low: a.low })),
    r1,
  );
  r1 += 1;
  r1 = addSectionTitle(wsOsoby, 'Osoby (lista bledow / odchylen)', r1);
  r1 = addJsonTable(
    wsOsoby,
    outliers.map((o) => {
      const b = o.bound as ValidationIqrBounds | null;
      const key = o.key;
      return {
        Agent: String(o.agent_celny ?? '').trim() || '—',
        Dzial: o.dzial ?? '',
        DataMRN: o.data_mrn,
        MRN: o.numer_mrn,
        NrSAD: o.nr_sad,
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        OplatyCelneRazem: o.fees,
        MasaNetto: o.mass,
        Coef: o.coef,
        Side: o.outlierSide,
        Limit: o.limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: o.discrepancyPct,
      };
    }),
    r1,
  );
  wsOsoby['!cols'] = [
    { wch: 18 },
    { wch: 12 },
    { wch: 18 },
    { wch: 14 },
    { wch: 22 },
    { wch: 12 },
    { wch: 10 },
    { wch: 12 },
    { wch: 16 },
    { wch: 12 },
    { wch: 18 },
    { wch: 16 },
    { wch: 12 },
    { wch: 12 },
    { wch: 8 },
    { wch: 12 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 14 },
  ];
  xlsx.utils.book_append_sheet(wb, wsOsoby, 'Osoby');

  // Dni
  const wsDni = xlsx.utils.aoa_to_sheet([]);
  r1 = 1;
  r1 = addKeyValueMeta(wsDni, metaRows, r1);
  r1 += 1;
  r1 = addSectionTitle(wsDni, 'Dni (podsumowanie)', r1);
  r1 = addJsonTable(
    wsDni,
    days.map((d) => ({
      DateStart: d.date,
      DateEnd: d.end,
      Total: d.total,
      OutliersHigh: d.outliersHigh,
      OutliersLow: d.outliersLow,
      Singles: d.singles,
    })),
    r1,
  );
  r1 += 1;
  r1 = addSectionTitle(wsDni, 'Dni (pozycje)', r1);
  const dayItems = base
    .slice()
    .sort(
      (a, b) =>
        String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
        String(a.outlierSide ?? '').localeCompare(String(b.outlierSide ?? '')) ||
        String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
        String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
    )
    .map((it) => {
      const b = bounds.get(it._boundsKey) ?? null;
      const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
      const discrepancyPct =
        it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null
          ? computeDiscrepancyPct(it.coef, limit)
          : null;
      const key = it.key;
      return {
        BucketStart: it.bucketStart,
        BucketEnd: it.bucketEnd,
        DataMRN: it.data_mrn,
        MRN: it.numer_mrn,
        NrSAD: it.nr_sad,
        Agent: it.agent_celny,
        Dzial: it.dzial,
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        OplatyCelneRazem: it.fees,
        MasaNetto: it.mass,
        Coef: it.coef,
        ManualVerified: it.verifiedManual,
        Checkable: it.checkable,
        OutlierSide: it.outlierSide,
        Limit: limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: discrepancyPct,
        RowId: it.rowId,
      };
    });
  r1 = addJsonTable(wsDni, dayItems, r1);
  xlsx.utils.book_append_sheet(wb, wsDni, 'Dni');

  // Grupy
  const wsGrupy = xlsx.utils.aoa_to_sheet([]);
  r1 = 1;
  r1 = addKeyValueMeta(wsGrupy, metaRows, r1);
  r1 += 1;
  r1 = addSectionTitle(wsGrupy, 'Grupy (podsumowanie)', r1);
  r1 = addJsonTable(
    wsGrupy,
    groups.map((g) => ({
      Count: g.count,
      Odbiorca: g.key.odbiorca,
      KodTowaru: g.key.kod_towaru,
      Waluta: g.key.waluta,
      KursWaluty: g.key.kurs_waluty,
      WarunkiDostawy: g.key.warunki_dostawy,
      KrajWysylki: g.key.kraj_wysylki,
      TransportRodzaj: g.key.transport_na_granicy_rodzaj,
    })),
    r1,
  );
  r1 += 1;
  r1 = addSectionTitle(wsGrupy, 'Grupy (pozycje)', r1);
  const groupItems = base
    .slice()
    .sort(
      (a, b) =>
        JSON.stringify(a.key).localeCompare(JSON.stringify(b.key)) ||
        String(a.bucketStart ?? '').localeCompare(String(b.bucketStart ?? '')) ||
        String(a.data_mrn ?? '').localeCompare(String(b.data_mrn ?? '')) ||
        String(a.numer_mrn ?? '').localeCompare(String(b.numer_mrn ?? '')),
    )
    .map((it) => {
      const b = bounds.get(it._boundsKey) ?? null;
      const limit = it.outlierSide ? (b ? (it.outlierSide === 'high' ? b.upper : b.lower) : null) : null;
      const discrepancyPct =
        it.outlierSide && it.coef != null && Number.isFinite(it.coef) && limit != null
          ? computeDiscrepancyPct(it.coef, limit)
          : null;
      const key = it.key;
      return {
        Odbiorca: key?.odbiorca ?? '',
        KodTowaru: key?.kod_towaru ?? '',
        Waluta: key?.waluta ?? '',
        KursWaluty: key?.kurs_waluty ?? '',
        WarunkiDostawy: key?.warunki_dostawy ?? '',
        KrajWysylki: key?.kraj_wysylki ?? '',
        TransportRodzaj: key?.transport_na_granicy_rodzaj ?? '',
        BucketStart: it.bucketStart,
        BucketEnd: it.bucketEnd,
        DataMRN: it.data_mrn,
        MRN: it.numer_mrn,
        NrSAD: it.nr_sad,
        Agent: it.agent_celny,
        Dzial: it.dzial,
        OplatyCelneRazem: it.fees,
        MasaNetto: it.mass,
        Coef: it.coef,
        ManualVerified: it.verifiedManual,
        Checkable: it.checkable,
        OutlierSide: it.outlierSide,
        Limit: limit,
        Q1: b?.q1 ?? null,
        Q3: b?.q3 ?? null,
        IQR: b?.iqr ?? null,
        Lower: b?.lower ?? null,
        Upper: b?.upper ?? null,
        DiscrepancyPct: discrepancyPct,
        RowId: it.rowId,
      };
    });
  addJsonTable(wsGrupy, groupItems, r1);
  xlsx.utils.book_append_sheet(wb, wsGrupy, 'Grupy');

  fs.mkdirSync(path.dirname(params.filePath), { recursive: true });
  xlsx.writeFile(wb, params.filePath, { bookType: 'xlsx', compression: true });
  return { ok: true };
}
