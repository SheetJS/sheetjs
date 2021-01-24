/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require('xlsx');
var assert = require('assert');
var SheetJSSQL = require('./SheetJSSQL');
var Client = require('pg').Client;

/* Connection options (requires two databases sheetjs and sheetj5) */
var opts = {
  host     : 'localhost',
  user     : 'SheetJS',
  password : 'SheetJS',
};

/* Sample data table */
var init = [
  "DROP TABLE IF EXISTS pres",
  "CREATE TABLE pres (name text, idx smallint)",
  "INSERT INTO pres VALUES ('Barack Obama', 44)",
  "INSERT INTO pres VALUES ('Donald Trump', 45)",
  "DROP TABLE IF EXISTS fmts",
  "CREATE TABLE fmts (ext text, ctr text, multi smallint)",
  "INSERT INTO fmts VALUES ('XLSB', 'ZIP', 1)",
  "INSERT INTO fmts VALUES ('XLS',  'CFB', 1)",
  "INSERT INTO fmts VALUES ('XLML', '',    1)",
  "INSERT INTO fmts VALUES ('CSV',  '',    0)",
];

var conn1 = new Client(Object.assign({}, opts, {database: "sheetjs"}));
var conn2 = new Client(Object.assign({}, opts, {database: "sheetj5"}));
(async () => {
  await conn1.connect();
  for(var i = 0; i < init.length; ++i) await conn1.query(init[i]);

  /* Export table to XLSX */
  var wb = XLSX.utils.book_new();

  async function book_append_table(wb, name) {
    var r_f = await conn1.query('SELECT * FROM ' + name);
    var r = r_f.rows;
    var ws = XLSX.utils.json_to_sheet(r);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  await book_append_table(wb, "pres");
  await book_append_table(wb, "fmts");
  XLSX.writeFile(wb, "pgsql.xlsx");

  /* Capture first database info and close */
  var P1 = (await conn1.query("SELECT * FROM pres")).rows;
  var F1 = (await conn1.query("SELECT * FROM fmts")).rows;
  await conn1.end();

  /* Import XLSX to table */
  await conn2.connect();
  var wb2 = XLSX.readFile("pgsql.xlsx");
  var queries = SheetJSSQL.book_to_sql(wb2, "PGSQL");
  for(i = 0; i < queries.length; ++i) { console.log(queries[i]); await conn2.query(queries[i]); }

  /* Capture first database info and close */
  var P2 = (await conn2.query("SELECT * FROM pres")).rows;
  var F2 = (await conn2.query("SELECT * FROM fmts")).rows;
  await conn2.end();

  /* Compare results */
  assert.deepEqual(P1, P2);
  assert.deepEqual(F1, F2);

  /* Display results */
  console.log(P2);
  console.log(F2);
})();
