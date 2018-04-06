/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*global module, require, XLSX:true */
if(typeof require !== 'undefined' && typeof XLSX === 'undefined') XLSX = require('xlsx');

function generate_random_file(n) {
	if(!n) n = 100;
	var aoo = [];
	var x_ = 0, y_ = 0, xx = 0, xy = 0;
	for(var i = 0; i < n; ++i) {
		var y = Math.fround(2 * i + Math.random());
		aoo.push({x:i, y:y});
		x_ += i / n; y_ += y / n; xx += i*i; xy += i * y;
	}
	var m = Math.fround((xy - n * x_ * y_)/(xx - n * x_ * x_));
	console.log(m, Math.fround(y_ - m * x_), "JS Pre");
	var ws = XLSX.utils.json_to_sheet(aoo);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
	ws = XLSX.utils.aoa_to_sheet([[2, 0]]);
	XLSX.utils.sheet_set_array_formula(ws, "A1:B1", "LINEST(Sheet1!B2:B101,Sheet1!A2:A101)");
	XLSX.utils.book_append_sheet(wb, ws, "Sheet2");

	XLSX.writeFile(wb, "linreg.xlsx");
}
if(typeof module !== 'undefined') module.exports = {
	generate_random_file: generate_random_file
};
