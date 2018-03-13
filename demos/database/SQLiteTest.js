/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require('xlsx');
var assert = require('assert');
var SheetJSSQL = require('./SheetJSSQL');
var Database = require('better-sqlite3');
var db1 = new Database('SheetJS1.db');

/* Sample data table */
var init = [
  "DROP TABLE IF EXISTS pres",
  "CREATE TABLE pres (name TEXT, idx INTEGER)",
  "INSERT INTO pres VALUES ('Barack Obama', 44)",
  "INSERT INTO pres VALUES ('Donald Trump', 45)",
  "DROP TABLE IF EXISTS fmts",
  "CREATE TABLE fmts (ext TEXT, ctr TEXT, multi INTEGER)",
  "INSERT INTO fmts VALUES ('XLSB', 'ZIP', 1)",
  "INSERT INTO fmts VALUES ('XLS',  'CFB', 1)",
  "INSERT INTO fmts VALUES ('XLML', '',    1)",
  "INSERT INTO fmts VALUES ('CSV',  '',    0)",
];
db1.exec(init.join(";"));

/* Export database to XLSX */
var wb = XLSX.utils.book_new();
function book_append_table(wb, db, name) {
  var r = db.prepare('SELECT * FROM ' + name).all();
  var ws = XLSX.utils.json_to_sheet(r);
  XLSX.utils.book_append_sheet(wb, ws, name);
}
book_append_table(wb, db1, "pres");
book_append_table(wb, db1, "fmts");
XLSX.writeFile(wb, "sqlite.xlsx");

/* Import XLSX to database */
var db2 = new Database('SheetJS2.db');
var wb2 = XLSX.readFile("sqlite.xlsx");
var queries = SheetJSSQL.book_to_sql(wb2, "SQLITE");
queries.forEach(function(q) { db2.exec(q); });

/* Compare databases */
var P1 = db1.prepare("SELECT * FROM pres").all();
var P2 = db2.prepare("SELECT * FROM pres").all();
var F1 = db1.prepare("SELECT * FROM fmts").all();
var F2 = db2.prepare("SELECT * FROM fmts").all();
assert.deepEqual(P1, P2);
assert.deepEqual(F1, F2);

console.log(P2);
console.log(F2);

