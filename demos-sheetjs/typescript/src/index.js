/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/* eslint-env browser */
/* global require */
var readFirstSheet = require("../").readFirstSheet;
console.log(readFirstSheet("a,b,c\n1,2,3\n4,5,6", {type:"binary"}));