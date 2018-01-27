/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require('xlsx');
var assert = require('assert');
var SheetJSSQL = require('./SheetJSSQL');
var mysql = require('mysql2/promise');

/* Connection options (requires two databases sheetjs and sheetj5) */
var opts = {
  host     : 'localhost',
  user     : 'SheetJS',
  password : 'SheetJS',
};

/* Sample data table */
var init = [
  "DROP TABLE IF EXISTS pres",
  "CREATE TABLE pres (name TEXT, idx TINYINT)",
  "INSERT INTO pres VALUES ('Barack Obama', 44)",
  "INSERT INTO pres VALUES ('Donald Trump', 45)",
  "DROP TABLE IF EXISTS fmts",
  "CREATE TABLE fmts (ext TEXT, ctr TEXT, multi TINYINT)",
  "INSERT INTO fmts VALUES ('XLSB', 'ZIP', 1)",
  "INSERT INTO fmts VALUES ('XLS',  'CFB', 1)",
  "INSERT INTO fmts VALUES ('XLML', '',    1)",
  "INSERT INTO fmts VALUES ('CSV',  '',    0)",
];

(async () => {
  const conn1 = await mysql.createConnection({...opts, database: "sheetjs"});
  for(var i = 0; i < init.length; ++i) await conn1.query(init[i]);

  /* Export table to XLSX */
  var wb = XLSX.utils.book_new();

  async function book_append_table(wb, name) {
    var r_f = await conn1.query('SELECT * FROM ' + name);
    var r = r_f[0];
    var ws = XLSX.utils.json_to_sheet(r);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  await book_append_table(wb, "pres");
  await book_append_table(wb, "fmts");
  XLSX.writeFile(wb, "mysql.xlsx");

  /* Capture first database info and close */
  var P1 = (await conn1.query("SELECT * FROM pres"))[0];
  var F1 = (await conn1.query("SELECT * FROM fmts"))[0];
  await conn1.close();

  /* Import XLSX to table */
  const conn2 = await mysql.createConnection({...opts, database: "sheetj5"});
  var wb2 = XLSX.readFile("mysql.xlsx");
  var queries = SheetJSSQL.book_to_sql(wb2, "MYSQL");
  for(i = 0; i < queries.length; ++i) await conn2.query(queries[i]);

  /* Capture first database info and close */
  var P2 = (await conn2.query("SELECT * FROM pres"))[0];
  var F2 = (await conn2.query("SELECT * FROM fmts"))[0];
  await conn2.close();

  /* Compare results */
  assert.deepEqual(P1, P2);
  assert.deepEqual(F1, F2);

  /* Display results */
  console.log(P2);
  console.log(F2);
})();
