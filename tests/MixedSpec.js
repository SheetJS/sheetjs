var XLSX = require('../');
var testCommon = require('./Common.js');

var file = 'mixed_sheets.xlsx';

describe(file, function () {
	testCommon(file);
});