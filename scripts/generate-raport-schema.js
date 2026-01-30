/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const projectRoot = path.resolve(__dirname, '..');
const inputXlsx = process.argv[2] || path.join(projectRoot, 'data_sample', 'raport.xlsx');
const prismaSchemaPath = path.join(projectRoot, 'prisma', 'schema.prisma');
const columnsTsPath = path.join(projectRoot, 'src', 'raportColumns.ts');

function normalizeHeader(value) {
  if (value == null) return '';
  return String(value).trim();
}

function toAscii(value) {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/ł/g, 'l')
    .replace(/Ł/g, 'L');
}

function toFieldName(header, colIndex1Based) {
  const raw = normalizeHeader(header);
  if (!raw) return `col_${colIndex1Based}`;
  const ascii = toAscii(raw);
  let name = ascii
    .replace(/["“”]/g, '')
    .replace(/[()]/g, ' ')
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .toLowerCase();

  if (!name) name = `col_${colIndex1Based}`;
  if (/^[0-9]/.test(name)) name = `c_${name}`;
  return name;
}

function findHeaderRow(aoa) {
  for (let i = 0; i < aoa.length; i++) {
    const row = aoa[i] || [];
    const hasAny = row.some((cell) => normalizeHeader(cell) !== '');
    if (hasAny) return i;
  }
  return 0;
}

function ensureDir(filePath) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
}

function main() {
  if (!fs.existsSync(inputXlsx)) {
    console.error(`Nie znaleziono pliku: ${inputXlsx}`);
    process.exit(1);
  }

  const wb = xlsx.readFile(inputXlsx, { cellDates: true });
  const firstSheetName = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheetName];
  const aoa = xlsx.utils.sheet_to_json(ws, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: false,
  });

  const headerRowIndex = findHeaderRow(aoa);
  const headerRow = aoa[headerRowIndex] || [];

  // Find last non-empty column
  let lastCol = 0;
  for (let i = 0; i < headerRow.length; i++) {
    if (normalizeHeader(headerRow[i]) !== '') lastCol = i + 1;
  }
  const headers = headerRow.slice(0, lastCol);

  const used = new Map();
  const columns = headers.map((h, idx) => {
    const colIndex = idx + 1;
    const label = normalizeHeader(h);
    const base = toFieldName(label, colIndex);
    const current = used.get(base) || 0;
    used.set(base, current + 1);
    const field = current === 0 ? base : `${base}_${current + 1}`;
    return { field, label, colIndex };
  });

  const schemaLines = [];
  schemaLines.push('generator client {');
  schemaLines.push('  provider = "prisma-client-js"');
  schemaLines.push('}');
  schemaLines.push('');
  schemaLines.push('datasource db {');
  schemaLines.push('  provider = "sqlite"');
  schemaLines.push('}');
  schemaLines.push('');
  schemaLines.push('model RaportMeta {');
  schemaLines.push('  id         Int      @id');
  schemaLines.push('  importedAt DateTime @default(now())');
  schemaLines.push('  sourceFile String?');
  schemaLines.push('  rowCount   Int      @default(0)');
  schemaLines.push('');
  schemaLines.push('  @@map("raport_meta")');
  schemaLines.push('}');
  schemaLines.push('');
  schemaLines.push('model RaportRow {');
  schemaLines.push('  id        Int     @id @default(autoincrement())');
  schemaLines.push('  rowNumber Int?');
  columns.forEach(({ field }) => {
    schemaLines.push(`  ${field} String?`);
  });
  schemaLines.push('');
  schemaLines.push('  @@map("raport_rows")');
  schemaLines.push('  @@index([id])');
  schemaLines.push('}');
  schemaLines.push('');

  const columnsTsLines = [];
  columnsTsLines.push('export type RaportColumn = {');
  columnsTsLines.push('  field: string;');
  columnsTsLines.push('  label: string;');
  columnsTsLines.push('};');
  columnsTsLines.push('');
  columnsTsLines.push('export const RAPORT_COLUMNS: ReadonlyArray<RaportColumn> = [');
  for (const col of columns) {
    const labelLiteral = JSON.stringify(col.label || `Kolumna ${col.colIndex}`);
    columnsTsLines.push(`  { field: ${JSON.stringify(col.field)}, label: ${labelLiteral} },`);
  }
  columnsTsLines.push('] as const;');
  columnsTsLines.push('');

  ensureDir(prismaSchemaPath);
  ensureDir(columnsTsPath);

  fs.writeFileSync(prismaSchemaPath, schemaLines.join('\n'), 'utf8');
  fs.writeFileSync(columnsTsPath, columnsTsLines.join('\n'), 'utf8');

  console.log(`OK: wygenerowano ${path.relative(projectRoot, prismaSchemaPath)}`);
  console.log(`OK: wygenerowano ${path.relative(projectRoot, columnsTsPath)}`);
  console.log(`Arkusz: ${firstSheetName}`);
  console.log(`Kolumny: ${columns.length}`);
}

main();
