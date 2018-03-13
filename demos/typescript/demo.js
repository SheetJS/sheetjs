/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/* eslint-env node */
var readFirstSheet = require("./").readFirstSheet;
console.log(readFirstSheet("../../sheetjs.xlsb", {type:"file", cellDates:true}));