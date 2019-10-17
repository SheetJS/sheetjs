/*! s.js (C) 2019-present SheetJS -- https://sheetjs.com */
/* vim: set ts=2: */

var assert = require("assert");
var XLSX = require("../../../");
var S = require("../");

assert(S != null);
S.set_XLSX(XLSX);
assert(S.get_XLSX() == XLSX);
assert(S.get_XLSX().version);