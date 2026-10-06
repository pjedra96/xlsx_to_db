# xlsx_to_database
Generic NodeJS scripts that read Excel (`.xlsx`) workbooks and upload their rows to a MongoDB collection (`index.js`) or a SQL table (`sql.js`, see [Saving to a SQL database](#saving-to-a-sql-database)). No schemas or database names are built in. Field names come from the spreadsheet itself. Both scripts read sheets the same way, using `lib/workbook.js`.

## Getting Started

1. Clone the repository (https://www.github.com/pjedra96/xlsx_to_database) to a directory of your choice.
2. Ensure that you have NodeJS installed (v20.12+) on your system.
3. Run `npm install` to download the required packages.
4. Copy `.env.example` to `.env` and set `MONGODB_URI`. `.env` is git-ignored.
5. Put the workbooks in `./spreadsheets` (or pass file/folder paths) and run `npm start`. Use `npm run dry-run` first to see what will be written without touching the database.

## Sheet layouts

The layout is detected per sheet. Pass `--layout table|matrix` to force one.

**table**: The header row holds the field names. It is detected automatically, skipping any title lines above the table, or can be set with `--header-row n`. Each row below it becomes one document.

| Name  | Code | Opened     |
|-------|------|------------|
| Aldi  | A01  | 01/08/2021 |

→ `{ Name: "Aldi", Code: "A01", Opened: ISODate("2021-08-01") }`

**matrix**: The top-left cell is a label, the header row holds month dates (`01/09/21`, as real dates or as text) and the first column holds day numbers (`1st`, `2nd`, …). Each populated cell becomes one document. Rows such as `Total` and impossible dates such as the 31st of September are skipped.

| House Visitors | 01/08/21 | 01/09/21 |
|------------------|----------|----------|
| 1st              | 2219     | 3061     |

→ `{ label: "House Visitors", date: ISODate("2021-08-01"), value: 2219 }`, …

Matrix documents are upserted on `label` + `date`, so re-running the import updates them instead of duplicating them. The field names can be changed with `--label-field`, `--date-field` and `--value-field`.

## Options

Run `node index.js --help` for the full list. The most useful are:

```
node index.js spreadsheets/NRP_Vehicle.xlsx --db Motorpoint --collection carpark --dry-run
node index.js --db MyDb --key Code            # upsert table rows on the Code column
node index.js --sheet "Aldi" --sheet "Car Park"
```

By default each workbook goes into a collection named after its file. The database is the one named in `MONGODB_URI` unless `--db` is given.

## Saving to a SQL database

`sql.js` takes the same reading and writing options as `index.js`, but writes to PostgreSQL, MySQL/MariaDB, SQL Server or SQLite.

1. Install the driver for your database. Only that one is needed:

   | Database       | `SQL_URL` form                                   | Driver                       |
   |----------------|--------------------------------------------------|------------------------------|
   | PostgreSQL     | `postgres://user:password@host:5432/database`    | `npm i pg`                   |
   | MySQL/MariaDB  | `mysql://user:password@host:3306/database`       | `npm i mysql2`               |
   | SQL Server     | `mssql://user:password@host:1433/database`       | `npm i tedious`              |
   | SQLite         | `sqlite:./path/to/file.db`                       | `npm i better-sqlite3`       |

   Special characters in the password must be URL-encoded (e.g. `@` → `%40`). Extra driver settings can go in the query string, e.g. `mssql://…/db?options={"trustServerCertificate":true}`.
2. Set `SQL_URL` in `.env` (or pass `--url`).
3. Run `npm run sql:dry-run` to see the column types and rows it will write, then `npm run sql`.

Each workbook goes into a table named after its file, unless `--table` is given. The table is created if it doesn't exist, with an auto-increment `id` column plus one column per field. Column types are picked from the data: integer, double, boolean, date, datetime or text. If a later import brings new fields, the missing columns are added.

Rows are inserted. With `--key` (and by default for matrix sheets, on `label` + `date`), each row instead updates the existing row with the same key values, or is inserted if there isn't one. Re-running an import therefore doesn't duplicate rows, and no unique index is needed. Each sheet is written in a single transaction, so a failure leaves that sheet's table unchanged.

```
node sql.js --url sqlite:./imports.db spreadsheets/Test_spreadsheet.xlsx
node sql.js --table shops --key Code
```
