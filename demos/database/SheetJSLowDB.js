/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */

var XLSX = require('xlsx');
var ObjUtils = require('./ObjUtils');

function SheetJSAdapter() { this.defaultValue = {}; };

SheetJSAdapter.prototype.read = function() { return this.defaultValue; };
SheetJSAdapter.prototype.write = function(/*data*/) {};

SheetJSAdapter.prototype.dumpRaw = function() { return ObjUtils.object_to_workbook(this.defaultValue); };
SheetJSAdapter.prototype.dump = function(options) { XLSX.write(this.dumpRaw(), options); };
SheetJSAdapter.prototype.dumpFile = function(path, options) { XLSX.writeFile(this.dumpRaw(), path, options); };

SheetJSAdapter.prototype.loadRaw = function(wb) { ObjUtils.workbook_set_object(this.defaultValue, wb); };
SheetJSAdapter.prototype.load = function(data, options) { this.loadRaw(XLSX.read(data, options)); };
SheetJSAdapter.prototype.loadFile = function(path, options) { this.loadRaw(XLSX.readFile(path, options)); };

if(typeof module !== 'undefined') module.exports = SheetJSAdapter;
