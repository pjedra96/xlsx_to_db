#!/usr/bin/env node
// Imports the rows of Excel (.xlsx) workbooks into a SQL table: PostgreSQL, MySQL/MariaDB, SQL Server or SQLite.
// Sheets are read exactly as index.js reads them (see lib/workbook.js).
// Tables are created when missing and gain a column for any new field; column types are picked from the data.
// Rows are inserted, or updated-or-inserted on --key (matrix sheets default to label + date,
// so re-running an import does not duplicate rows).
//
// Only the driver for your database needs installing: npm i pg | mysql2 | tedious | better-sqlite3

import { parseArgs } from 'node:util';
import path from 'node:path';
import knex from 'knex';
import { readingOptions, readingUsage, readerSettings, findWorkbooks, readWorkbook } from './lib/workbook.js';

const usage = `Usage: node sql.js [options] [file.xlsx | directory ...]   (default: ./spreadsheets)

Connection (put SQL_URL in a .env file or the environment, never in the code):
  --url <url>            connection URL (default: $SQL_URL), one of
                           postgres://user:password@host:5432/database   (npm i pg)
                           mysql://user:password@host:3306/database      (npm i mysql2)
                           mssql://user:password@host:1433/database      (npm i tedious)
                           sqlite:./path/to/file.db                      (npm i better-sqlite3)
  --table <name>         table name (default: the workbook's file name)

${readingUsage}`;

// URL scheme -> Knex client (which is also the npm driver package, except mssql -> tedious)
const clients = {
  postgres: 'pg',
  postgresql: 'pg',
  mysql: 'mysql2',
  mariadb: 'mysql2',
  mssql: 'mssql',
  sqlserver: 'mssql',
  sqlite: 'better-sqlite3',
};

// SQL Server allows 2100 parameters per statement, the lowest limit of the supported databases
const maxParameters = 2000;

try {
  process.loadEnvFile();
} catch {
  // No .env file - rely on the real environment
}

// Parse command-line arguments
const { values: opts, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    ...readingOptions,
    url: { type: 'string' },
    table: { type: 'string' },
  },
});

if (opts.help) {
  console.log(usage);
  process.exit(0);
}

const url = opts.url ?? process.env.SQL_URL;
const dryRun = opts['dry-run'];
let settings, files, config;
try {
  settings = readerSettings(opts);
  files = await findWorkbooks(positionals.length ? positionals : ['spreadsheets']);
  if (!dryRun) config = knexConfig(url);
} catch (err) {
  fail(err.message);
}
if (files.length === 0) fail('No .xlsx files found');

const db = dryRun ? null : knex(config);
if (db) {
  await db.raw('select 1'); // fail fast on a bad connection or missing driver
  console.log(`Connected to ${config.client} database`);
}

try {
  for (const file of files) {
    try {
      await importWorkbook(file);
    } catch (err) {
      console.error(`${path.basename(file)}: ${err.message}`);
      process.exitCode = 1;
    }
  }
} finally {
  await db?.destroy();
}

console.log(process.exitCode ? 'Finished with errors' : 'XLSX to database transfer complete');

async function importWorkbook(file) {
  const tableName = opts.table ?? path.parse(file).name;

  for await (const { name, layout, records, key } of readWorkbook(file, settings)) {
    const where = `${path.basename(file)} [${name}]`;
    if (!layout) {
      console.log(`${where}: empty, skipped`);
      continue;
    }

    const types = columnTypes(records);
    const rows = records.map(record => toRow(record, types));
    console.log(`${where}: ${layout} layout, ${rows.length} row(s) -> ${tableName}`);
    if (dryRun) {
      console.log(Object.fromEntries(types), rows.slice(0, 3));
      continue;
    }
    if (rows.length === 0) continue;

    await ensureTable(tableName, types);
    const counts = await writeRows(tableName, rows, key, types.size);
    console.log(`${where}: ${Object.entries(counts).map(([label, count]) => `${count} ${label}`).join(', ')}`);
  }
}

// Builds the Knex config from a connection URL; throws when it is missing or of an unknown kind
function knexConfig(connectionUrl) {
  if (!connectionUrl) throw new Error('No connection URL: set SQL_URL (e.g. in .env) or pass --url');

  const scheme = /^([a-z0-9]+):/i.exec(connectionUrl)?.[1].toLowerCase();
  const client = clients[scheme];
  if (!client) throw new Error(`Unsupported connection URL scheme "${scheme ?? connectionUrl}" - use one of: ${Object.keys(clients).join(', ')}`);

  if (client === 'better-sqlite3') return { client, connection: { filename: connectionUrl.slice(scheme.length + 1) }, useNullAsDefault: true };
  if (client === 'pg') return { client, connection: connectionUrl };

  // mysql2 and tedious take an options object; extra URL parameters are passed through (JSON values allowed,
  // e.g. mssql://...?options={"trustServerCertificate":true})
  const parsed = new URL(connectionUrl);
  const connection = {
    [client === 'mssql' ? 'server' : 'host']: parsed.hostname,
    database: decodeURIComponent(parsed.pathname.slice(1)),
    user: decodeURIComponent(parsed.username),
    password: decodeURIComponent(parsed.password),
  };
  if (parsed.port) connection.port = Number(parsed.port);
  for (const [option, value] of parsed.searchParams) {
    try {
      connection[option] = JSON.parse(value);
    } catch {
      connection[option] = value;
    }
  }
  return { client, connection };
}

// Picks one column type per field from the values it holds, widening on conflicts
// (integer + double -> double, date + datetime -> datetime, anything else mixed -> text)
function columnTypes(records) {
  const types = new Map();
  for (const record of records) {
    for (const [field, value] of Object.entries(record)) {
      const type = valueType(value);
      const current = types.get(field);
      if (!current || current === type) types.set(field, type);
      else if ([current, type].every(t => t === 'integer' || t === 'double')) types.set(field, 'double');
      else if ([current, type].every(t => t === 'date' || t === 'datetime')) types.set(field, 'datetime');
      else types.set(field, 'text');
    }
  }
  return types;
}

function valueType(value) {
  if (value instanceof Date) return value.getTime() % 86400000 === 0 ? 'date' : 'datetime';
  if (typeof value === 'number') return Number.isInteger(value) && Math.abs(value) <= 2147483647 ? 'integer' : 'double';
  if (typeof value === 'boolean') return 'boolean';
  return 'text';
}

// Converts a record to column values. Dates are sent as UTC text so no driver shifts them into local time.
function toRow(record, types) {
  return Object.fromEntries(
    Object.entries(record).map(([field, value]) => {
      const type = types.get(field);
      if (value instanceof Date) {
        const iso = value.toISOString();
        return [field, type === 'date' ? iso.slice(0, 10) : iso.slice(0, 23).replace('T', ' ')];
      }
      return [field, type === 'text' ? String(value) : value];
    }),
  );
}

async function ensureTable(tableName, types) {
  const addColumn = (table, field) => {
    const type = types.get(field);
    if (type === 'integer') table.integer(field);
    else if (type === 'double') table.double(field);
    else if (type === 'boolean') table.boolean(field);
    else if (type === 'date') table.date(field);
    else if (type === 'datetime') table.datetime(field, { useTz: false });
    else table.text(field);
  };

  if (!(await db.schema.hasTable(tableName))) {
    await db.schema.createTable(tableName, table => {
      if (!types.has('id')) table.increments('id');
      for (const field of types.keys()) addColumn(table, field);
    });
    console.log(`Created table ${tableName}`);
    return;
  }

  const missing = [];
  for (const field of types.keys()) {
    if (!(await db.schema.hasColumn(tableName, field))) missing.push(field);
  }
  if (missing.length) {
    await db.schema.alterTable(tableName, table => missing.forEach(field => addColumn(table, field)));
    console.log(`Added column(s) to ${tableName}: ${missing.join(', ')}`);
  }
}

// Without a key, rows are bulk inserted. With a key, each row updates the row matching its key
// values or is inserted when none matches - plain SQL that behaves the same on every database
// and needs no unique index. Each sheet is written in one transaction.
async function writeRows(tableName, rows, key, columnCount) {
  return db.transaction(async trx => {
    if (!key.length) {
      await db.batchInsert(tableName, rows, Math.max(1, Math.floor(maxParameters / columnCount))).transacting(trx);
      return { inserted: rows.length };
    }

    const counts = { inserted: 0, updated: 0 };
    for (const row of rows) {
      const match = Object.fromEntries(key.map(field => [field, row[field] ?? null]));
      const updated = await trx(tableName).where(match).update(row);
      if (updated) {
        counts.updated += updated;
      } else {
        await trx(tableName).insert(row);
        counts.inserted++;
      }
    }
    return counts;
  });
}

// This function is used to indicate an error condition and terminate the program.
function fail(message) {
  console.error(`${message}\n\n${usage}`);
  process.exit(1);
}
