#!/usr/bin/env node
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

var fs = require('fs'), path = require('path');
var express = require('express'), app = express();
var sprintf = require('printj').sprintf;
var logit = require('../server/_logit');
var cors = require('../server/_cors');

var port = +process.argv[2] || +process.env.PORT || 7262;
var basepath = process.cwd();

app.use(logit.mw);
app.use(cors.mw);
app.use(require('express-formidable')());
app.post('/upload', function(req, res) {
	fs.writeFile(req.fields.file, req.fields.data, 'base64', function(err, r) {
		res.end("wrote to " + req.fields.file);
	});
});
app.use(express.static(path.resolve(basepath)));
app.use(require('serve-index')(basepath, {'icons':true}));

app.listen(port, function() { console.log('Serving HTTP on port ' + port); });

