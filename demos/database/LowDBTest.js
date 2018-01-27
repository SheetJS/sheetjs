/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var low = require('lowdb');
var SheetJSAdapter = require('./SheetJSLowDB');
var adapter = new SheetJSAdapter();
var db = low(adapter);

db.defaults({ posts: [], user: {}, count: 0 }).write();
db.get('posts').push({ id: 1, title: 'lowdb is awesome'}).write();
db.set('user.name', 'typicode').write();
db.update('count', function(n) { return n + 1; }).write();

adapter.dumpFile('ldb1.xlsx');

var adapter2 = new SheetJSAdapter();
adapter2.loadFile('ldb1.xlsx');
var db2 = low(adapter2);

db2.get('posts').push({ id: 2, title: 'mongodb is not'}).write();
db2.set('user.name', 'sheetjs').write();
db2.update('count', function(n) { return n + 1; }).write();

adapter2.dumpFile('ldb2.xlsx');
