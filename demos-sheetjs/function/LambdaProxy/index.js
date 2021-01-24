/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
// base64 sheetjs.xlsb | curl -F "data=@-;filename=test.xlsb" http://localhost:3000/LambdaProxy

'use strict';
var XLSX = require('xlsx');
var Busboy = require('busboy');

exports.handler = function(event, context, callback) {
	/* set up busboy */
	var ctype = event.headers['Content-Type']||event.headers['content-type'];
	var bb = new Busboy({headers:{'content-type':ctype}});

	/* busboy is evented; accumulate the fields and files manually */
	var fields = {}, files = {};
	bb.on('error', function(err) { console.log('err', err); callback(err); });
	bb.on('field', function(fieldname, val) {fields[fieldname] = val });
	bb.on('file', function(fieldname, file, filename) {
		/* concatenate the individual data buffers */
		var buffers = [];
		file.on('data', function(data) { buffers.push(data); });
		file.on('end', function() { files[fieldname] = [Buffer.concat(buffers), filename]; });
	});

	/* on the finish event, all of the fields and files are ready */
	bb.on('finish', function() {
		/* grab the first file */
		var f = files[Object.keys(files)[0]];
		if(!f) callback(new Error("Must submit a file for processing!"));

		/* f[0] is a buffer, convert to string and interpret as Base64 */
		var wb = XLSX.read(f[0].toString(), {type:"base64"});

		/* grab first worksheet and convert to CSV */
		var ws = wb.Sheets[wb.SheetNames[0]];
		callback(null, { body: XLSX.utils.sheet_to_csv(ws) });
	});

	bb.end(event.body);
};
