#!/usr/bin/env node
// Imports the rows of Excel (.xlsx) workbooks into a MongoDB collection.
// Sheet layouts and reading rules are described in lib/workbook.js; sql.js is the SQL equivalent.

import { parseArgs } from 'node:util';
import path from 'node:path';
import { MongoClient } from 'mongodb';
import { readingOptions, readingUsage, readerSettings, findWorkbooks, readWorkbook } from './lib/workbook.js';

const usage = `Usage: node index.js [options] [file.xlsx | directory ...]   (default: ./spreadsheets)

Connection (put MONGODB_URI in a .env file or the environment, never in the code):
  --uri <uri>            MongoDB connection string (default: $MONGODB_URI)
  --db <name>            database name (default: the database in the URI)
  --collection <name>    collection name (default: the workbook's file name)

${readingUsage}`;

const batchSize = 1000;

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
    uri: { type: 'string' },
    db: { type: 'string' },
    collection: { type: 'string' },
  },
});

if (opts.help) {
  console.log(usage);
  process.exit(0);
}

const uri = opts.uri ?? process.env.MONGODB_URI;
const dryRun = opts['dry-run'];
let settings, files;
try {
  settings = readerSettings(opts);
  files = await findWorkbooks(positionals.length ? positionals : ['spreadsheets']);
} catch (err) {
  fail(err.message);
}
if (files.length === 0) fail('No .xlsx files found');
if (!dryRun && !uri) fail('No connection string: set MONGODB_URI (e.g. in .env) or pass --uri');

const client = dryRun ? null : await new MongoClient(uri).connect();
const db = client?.db(opts.db);
if (db) console.log(`Connected to database "${db.databaseName}"`);

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
  await client?.close();
}

console.log(process.exitCode ? 'Finished with errors' : 'XLSX to database transfer complete');

async function importWorkbook(file) {
  const collectionName = opts.collection ?? path.parse(file).name;

  for await (const { name, layout, records, key } of readWorkbook(file, settings)) {
    const where = `${path.basename(file)} [${name}]`;
    if (!layout) {
      console.log(`${where}: empty, skipped`);
      continue;
    }

    console.log(`${where}: ${layout} layout, ${records.length} record(s) -> ${collectionName}`);
    if (dryRun) {
      console.log(records.slice(0, 3));
      continue;
    }

    const counts = await writeRecords(db.collection(collectionName), records, key);
    console.log(`${where}: ${Object.entries(counts).map(([label, count]) => `${count} ${label}`).join(', ')}`);
  }
}

async function writeRecords(collection, records, key) {
  const counts = key.length ? { upserted: 0, updated: 0, unchanged: 0 } : { inserted: 0 };

  for (let i = 0; i < records.length; i += batchSize) {
    const batch = records.slice(i, i + batchSize);

    if (key.length) {
      const result = await collection.bulkWrite(
        batch.map(record => ({
          updateOne: { filter: Object.fromEntries(key.map(field => [field, record[field]])), update: { $set: record }, upsert: true },
        })),
        { ordered: false },
      );
      counts.upserted += result.upsertedCount;
      counts.updated += result.modifiedCount;
      counts.unchanged += result.matchedCount - result.modifiedCount;
    } else {
      counts.inserted += (await collection.insertMany(batch, { ordered: false })).insertedCount;
    }
  }
  return counts;
}

// This function is used to indicate an error condition and terminate the program.
function fail(message) {
  console.error(`${message}\n\n${usage}`);
  process.exit(1);
}
