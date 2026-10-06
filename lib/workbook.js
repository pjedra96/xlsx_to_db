// Reads Excel (.xlsx) workbooks into plain records - shared by the MongoDB (index.js) and SQL (sql.js) importers.
//
// Two sheet layouts are understood, detected per sheet unless --layout is given:
//   table  - a header row (on whichever row it starts) followed by one record per row;
//            the header cells become the field names.
//   matrix - a day-of-month x month grid: the top-left cell is a label (e.g. "house Visitors"),
//            the rest of the header row holds month dates (e.g. 01/09/21) and the first column
//            holds day numbers (1st, 2nd ... 31st). Each populated cell becomes one
//            { label, date, value } record; rows that are not day numbers (e.g. "Total") are skipped.

import { readdir, stat } from 'node:fs/promises';
import path from 'node:path';
import ExcelJS from 'exceljs';

const dayPattern = /^(\d{1,2})(?:st|nd|rd|th)?$/i;
const datePattern = /^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/; // dd/mm/yy or dd/mm/yyyy

// parseArgs options shared by both importers
export const readingOptions = {
  layout: { type: 'string', default: 'auto' },
  'header-row': { type: 'string' },
  sheet: { type: 'string', multiple: true },
  key: { type: 'string' },
  'label-field': { type: 'string', default: 'label' },
  'date-field': { type: 'string', default: 'date' },
  'value-field': { type: 'string', default: 'value' },
  'dry-run': { type: 'boolean', default: false },
  help: { type: 'boolean', short: 'h', default: false },
};

export const readingUsage = `Reading:
  --layout <type>        auto | table | matrix (default: auto)
  --header-row <n>       1-based row holding the headers (default: detected, skipping title lines)
  --sheet <name>         only import this sheet; repeat for several

Writing:
  --key <a,b,...>        upsert on these fields instead of inserting
                         (matrix sheets default to <label-field>,<date-field>)
  --label-field <name>   matrix field for the top-left label (default: label)
  --date-field <name>    matrix field for the date (default: date)
  --value-field <name>   matrix field for the cell value (default: value)
  --dry-run              read and print sample records without touching the database
  -h, --help             show this help`;

// Validates the parsed reading options; throws on bad input
export function readerSettings(opts) {
  if (!['auto', 'table', 'matrix'].includes(opts.layout)) throw new Error(`Unknown --layout "${opts.layout}"`);

  const headerRow = opts['header-row'] === undefined ? undefined : Number(opts['header-row']);
  if (headerRow !== undefined && !(Number.isInteger(headerRow) && headerRow >= 1)) throw new Error('--header-row must be a positive whole number');

  return {
    layout: opts.layout,
    headerRow,
    sheets: opts.sheet,
    key: opts.key?.split(',').map(field => field.trim()),
    matrixFields: { label: opts['label-field'], date: opts['date-field'], value: opts['value-field'] },
  };
}

/*
 If the target is a directory, it uses the readdir function to get a list of file
 names in the directory. It filters the file names to only include those that end
 with .xlsx and do not start with ~$.
 It then maps each file name to the full path by joining the directory and file name
 together. The resulting full paths are pushed to the found array. If the target is
 not a directory, it simply pushes the target itself to the found array. Finally,
 it returns the found array.
*/
export async function findWorkbooks(targets) {
  const found = [];
  for (const target of targets) {
    const info = await stat(target).catch(() => null);
    if (!info) throw new Error(`Not found: ${target}`);

    if (info.isDirectory()) {
      const names = await readdir(target);
      found.push(...names.filter(name => /\.xlsx$/i.test(name) && !name.startsWith('~$')).map(name => path.join(target, name)));
    } else {
      found.push(target);
    }
  }
  return found;
}

// Yields { name, layout, records, key } per sheet; layout is null for an empty sheet
export async function* readWorkbook(file, settings) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);

  for (const worksheet of workbook.worksheets) {
    if (settings.sheets && !settings.sheets.includes(worksheet.name)) continue;

    const sheet = readSheet(worksheet, settings.headerRow);
    if (!sheet) {
      yield { name: worksheet.name, layout: null, records: [], key: [] };
      continue;
    }

    const layout = settings.layout === 'auto' ? (looksLikeMatrix(sheet) ? 'matrix' : 'table') : settings.layout;
    const records = layout === 'matrix' ? matrixRecords(sheet, settings.matrixFields) : tableRecords(sheet);
    const key = settings.key ?? (layout === 'matrix' ? [settings.matrixFields.label, settings.matrixFields.date] : []);

    const missing = key.filter(field => records.length && !records.some(record => field in record));
    if (missing.length) throw new Error(`[${worksheet.name}] --key field(s) not found: ${missing.join(', ')}`);

    yield { name: worksheet.name, layout, records, key };
  }
}

// Returns { header, columns, body } where columns are the indexes of the non-empty header cells,
// or null when the sheet has no data
function readSheet(worksheet, headerRow) {
  const grid = [];
  for (let r = 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    grid.push(Array.from({ length: worksheet.columnCount }, (unused, c) => cellValue(row.getCell(c + 1).value)));
  }

  // Without --header-row, the header is the first row as wide as the widest of the first 20,
  // which skips any title lines above the table
  const filled = row => row.filter(value => value !== null).length;
  const widest = Math.max(0, ...grid.slice(0, 20).map(filled));
  const headerIndex = headerRow !== undefined ? headerRow - 1 : grid.findIndex(row => widest > 0 && filled(row) === widest);
  const header = grid[headerIndex];
  if (!header) return null;

  const columns = header.flatMap((value, c) => (value === null ? [] : [c]));
  const body = grid.slice(headerIndex + 1).filter(row => columns.some(c => row[c] !== null));
  return { header, columns, body };
}

// Flattens ExcelJS cell values (formulas, rich text, hyperlinks, errors) to plain values; blanks become null
function cellValue(value) {
  if (value === null || value === undefined) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'string') return value.trim() || null;
  if (typeof value !== 'object') return value;
  if ('result' in value) return cellValue(value.result);
  if ('richText' in value) return cellValue(value.richText.map(part => part.text).join(''));
  if ('text' in value) return cellValue(value.text);
  return null;
}

function looksLikeMatrix({ header, columns, body }) {
  const [labelColumn, ...monthColumns] = columns;
  if (monthColumns.length === 0 || !monthColumns.every(c => headerDate(header[c]))) return false;

  const labels = body.map(row => row[labelColumn]).filter(value => value !== null);
  const days = labels.filter(value => dayNumber(value) !== null);
  return days.length > 0 && days.length >= labels.length / 2;
}

function tableRecords({ header, columns, body }) {
  const used = new Set();
  const fields = columns.map(c => [c, fieldName(header[c], used)]);

  return body.map(row => Object.fromEntries(fields.filter(([c]) => row[c] !== null).map(([c, field]) => [field, row[c]])));
}

function matrixRecords({ header, columns, body }, fields) {
  const [labelColumn, ...monthColumns] = columns;
  const label = String(header[labelColumn]);
  const records = [];

  for (const row of body) {
    const day = dayNumber(row[labelColumn]);
    if (day === null) continue; // e.g. a "Total" row

    for (const c of monthColumns) {
      const month = headerDate(header[c]);
      if (row[c] === null || !month) continue;

      const date = new Date(Date.UTC(month.getUTCFullYear(), month.getUTCMonth(), day));
      if (date.getUTCMonth() !== month.getUTCMonth()) continue; // e.g. the 31st of September

      records.push({ [fields.label]: label, [fields.date]: date, [fields.value]: toNumber(row[c]) });
    }
  }
  return records;
}

// Header cells may be real Excel dates or text such as "01/09/21" (read as day/month/year)
function headerDate(value) {
  if (value instanceof Date) return value;
  const match = datePattern.exec(String(value));
  if (!match) return null;

  const [, day, month, year] = match.map(Number);
  const date = new Date(Date.UTC(year < 100 ? 2000 + year : year, month - 1, day));
  return date.getUTCMonth() === month - 1 ? date : null;
}

function dayNumber(value) {
  const match = dayPattern.exec(String(value ?? ''));
  const day = match ? Number(match[1]) : 0;
  return day >= 1 && day <= 31 ? day : null;
}

function toNumber(value) {
  const number = typeof value === 'number' ? value : Number(String(value).replace(/,/g, ''));
  return Number.isFinite(number) ? number : value;
}

// Field names may not contain "." or start with "$" (MongoDB rules, harmless for SQL); duplicates get a numeric suffix
function fieldName(header, used) {
  const base = (header instanceof Date ? header.toISOString().slice(0, 10) : String(header)).replace(/\./g, '_').replace(/^\$+/, '') || 'field';
  let name = base;
  for (let n = 2; used.has(name); n++) name = `${base}_${n}`;
  used.add(name);
  return name;
}
