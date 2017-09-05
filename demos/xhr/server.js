#!/usr/bin/env node
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

var fs = require('fs'), path = require('path');
var express = require('express'), app = express();
var sprintf = require('printj').sprintf;

var port = +process.argv[2] || +process.env.PORT || 7262;
var basepath = process.cwd();

function doit(cb) {
	return function(req, res, next) {
		cb(req, res);
		next();
	};
}

app.use(doit(function(req, res) {
	console.log(sprintf("%s %s %d", req.method, req.url, res.statusCode));
}));
app.use(doit(function(req, res) {
	res.header('Access-Control-Allow-Origin', '*');
}));
app.use(require('express-formidable')());
app.post('/upload', function(req, res) {
	fs.writeFile(req.fields.file, req.fields.data, 'base64', function(err, r) {
		res.end("wrote to " + req.fields.file);
	});
});
app.use(express.static(path.resolve(basepath)));
app.use(require('serve-index')(basepath, {'icons':true}));

app.listen(port, function() { console.log('Serving HTTP on port ' + port); });

