/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require('xlsx');
var pr = require('propel');
var linest = require('./linest');

/* generate linreg.xlsx with 100 random points */
linest.generate_random_file(100);

/* get the first worksheet as an array of arrays, skip the first row */
var wb = XLSX.readFile('linreg.xlsx');
var ws = wb.Sheets[wb.SheetNames[0]];
var aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:true}).slice(1);

/* calculate the coefficients in JS */
(function(aoa) {
	var x_ = 0, y_ = 0, xx = 0, xy = 0, n = aoa.length;
	for(var i = 0; i < n; ++i) {
		x_ += aoa[i][0] / n;
		y_ += aoa[i][1] / n;
		xx += aoa[i][0] * aoa[i][0];
		xy += aoa[i][0] * aoa[i][1];
	}
	var m = Math.fround((xy - n * x_ * y_)/(xx - n * x_ * x_));
	console.log(m, Math.fround(y_ - m * x_), "JS Post");
})(aoa);

/* build X and Y vectors */
var tensor = pr.tensor(aoa).transpose();
var xs = tensor.slice(0, 1);
var ys = tensor.slice(1, 1);

/* compute the coefficient */
var n = xs.size;
var x_ = Math.fround(xs.reduceMean().dataSync()[0]);
var y_ = Math.fround(ys.reduceMean().dataSync()[0]);
var xx = Math.fround(xs.dot(xs.transpose()).dataSync()[0]);
var xy = Math.fround(xs.dot(ys.transpose()).dataSync()[0]);
var m = Math.fround((xy - n * x_ * y_)/(xx - n * x_ * x_));
var b_ = Math.fround(y_ - m * x_);
console.log(m, b_, "Propel");
var yh = xs.mul(m).add(b_);

/* export data to aoa */
var prdata = pr.concat([xs, ys, yh]).transpose();
var shape = prdata.shape;
var prarr = prdata.dataSync();
var praoa = [];
for(var j = 0; j < shape[0]; ++j) {
	praoa[j] = [];
	for(var i = 0; i < shape[1]; ++i) praoa[j][i] = prarr[j * shape[1] + i];
}

/* add headers and export */
praoa.unshift(["x", "y", "pred"]);
var new_ws = XLSX.utils.aoa_to_sheet(praoa);
var new_wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(new_wb, new_ws, "Sheet1");
XLSX.writeFile(new_wb, "propel.xls");
