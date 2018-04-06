/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
var XLSX = require('xlsx');
var tf = require('@tensorflow/tfjs');
var linest = require('./linest');

/* generate linreg.xlsx with 100 random points */
var N = 100;
linest.generate_random_file(N);

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
var tensor = tf.tensor2d(aoa).transpose();
console.log(tensor.shape);
var xs = tensor.slice([0,0], [1,tensor.shape[1]]).flatten();
var ys = tensor.slice([1,0], [1,tensor.shape[1]]).flatten();

/* set up variables with initial guess */
var x_ = xs.mean().dataSync()[0];
var y_ = ys.mean().dataSync()[0];
var a = tf.variable(tf.scalar(y_/x_));
var b = tf.variable(tf.scalar(Math.random()));

/* linear predictor */
function predict(x) { return tf.tidy(function() { return a.mul(x).add(b); }); }
/* mean square scoring */
function loss(yh, y) { return yh.sub(y).square().mean(); }

/* train */
for(var j = 0; j < 5; ++j) {
	var learning_rate = 0.0001 /(j+1), iterations = 1000;
	var optimizer = tf.train.sgd(learning_rate);

	for(var i = 0; i < iterations; ++i) optimizer.minimize(function() {
		var pred = predict(xs);
		var L = loss(pred, ys);
		return L
	});

	/* compute the coefficient */
	var m = a.dataSync()[0], b_ = b.dataSync()[0];
	console.log(m, b_, "TF " + iterations * (j+1));
}

/* export data to aoa */
var yh = predict(xs);
var tfdata = tf.stack([xs, ys, yh]).transpose();
var shape = tfdata.shape;
var tfarr = tfdata.dataSync();
var tfaoa = [];
for(j = 0; j < shape[0]; ++j) {
	tfaoa[j] = [];
	for(i = 0; i < shape[1]; ++i) tfaoa[j][i] = tfarr[j * shape[1] + i];
}

/* add headers and export */
tfaoa.unshift(["x", "y", "pred"]);
var new_ws = XLSX.utils.aoa_to_sheet(tfaoa);
var new_wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(new_wb, new_ws, "Sheet1");
XLSX.writeFile(new_wb, "tfjs.xls");
