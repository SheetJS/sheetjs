#!/usr/bin/env sjs
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

var XLSX = require('../../xlsx.js');

var io = require('io');
var file = io.open("sheetjs.xlsx", "rb");
var strs = [], str = "";
while((str = file.read()).length > 0) strs.push(str);
var data = (Buffer.concat(strs.map(function(x) { return new Buffer(x); })));

var wb = XLSX.read(data, {type:"buffer"});
console.log(wb.Sheets[wb.SheetNames[0]]);
