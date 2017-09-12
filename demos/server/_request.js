/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var XLSX = require('xlsx'), request = require('request');
var url = 'http://www.freddiemac.com/pmms/2017/historicalweeklydata.xls'
request(url, {encoding: null}, function(err, res, data) {
	if(err || res.statusCode !== 200) return;
	var wb = XLSX.read(data, {type:'buffer'});
	var ws = wb.Sheets[wb.SheetNames[0]];
	console.log(XLSX.utils.sheet_to_csv(ws, {blankrows:false}));
});
