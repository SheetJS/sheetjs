/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
var XLSX = require('xlsx');
var fs = require('fs');

onmessage = function(e) {
	try { switch(e.data.action) {
		case 'write':
			var ws = XLSX.utils.aoa_to_sheet(e.data.data);
			var wb = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
			postMessage({data: XLSX.write(wb, {type:'binary', bookType:e.data.type || e.data.file.match(/\.([^\.]*)$/)[1]})});
			break;
		case 'read':
			var wb;
			if(e.data.file) wb = XLSX.readFile(e.data.file);
			else wb = XLSX.read(e.data.data);
			var ws = wb.Sheets[wb.SheetNames[0]];
			postMessage({data: XLSX.utils.sheet_to_json(ws, {header:1})});
			break;
		default: throw "unknown action";
	}} catch(e) { postMessage({err:e.message || e}); }
};
