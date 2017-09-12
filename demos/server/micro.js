/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var fs = require('fs');
var URL = require('url');
var child_process = require('child_process');
var micro = require('micro'), formidable = require('formidable');
var logit = require('./_logit'), cors = require('./_cors');
var json2csv = require('json2csv');
var data = "a,b,c\n1,2,3".split("\n").map(function(x) { return x.split(","); });
var xlsx = '../../bin/xlsx.njs';

function get_data(req, res, type) {
	var file = "_tmp." + type;

	/* prepare CSV */
	var csv = json2csv({data:data, hasCSVColumnTitle:false});

	/* write it to a temp file */
	fs.writeFile('tmp.csv', csv, function(err1) {

		/* call xlsx to read the csv and write to another temp file */
		child_process.exec(xlsx+' tmp.csv -o '+ file, function(err, stdout, stderr){
			cors(req, res);
			/* read the new file and send it to the client */
			micro.send(res, 200, fs.readFileSync(file));
		});
	});
}

function get_file(req, res, file) {
	var csv = json2csv({data:data, hasCSVColumnTitle:false});
	fs.writeFile('tmp.csv', csv, function(err1) {
		/* write to specified file */
		child_process.exec(xlsx+' tmp.csv -o '+file, function(err, stdout, stderr) {
			cors(req, res);
			micro.send(res, 200, "wrote to " + file + "\n");
		});
	});
}

function post_data(req, res) {
	var form = new formidable.IncomingForm();
	form.on('file', function(field, file) {
		/* file.path is the location of the file in the system */
		child_process.exec(xlsx+' --arrays ' + file.path, post_cb(req, res));
	});
	form.parse(req);
}

function post_file(req, res, file) {
	child_process.exec(xlsx+' --arrays ' + file, post_cb(req, res));
}

function post_cb(req, res) {
	return function(err, stdout, stderr) {
		cors(req, res);
		/* xlsx --arrays writes JSON to stdout, so parse and assign to data var */
		data = JSON.parse(stdout);
		console.log(data);
		return micro.send(res, 200, "ok\n");
	};
}

function get(req, res) {
	var url = URL.parse(req.url, true);
	if(url.pathname.length > 1) micro.send(res, 404, "File not found");
	else if(url.query.t) get_data(req, res, url.query.t);
	else if(url.query.f) get_file(req, res, url.query.f);
	else micro.send(res, 403, "Forbidden\n");
}

function post(req, res) {
	var url = URL.parse(req.url, true);
	if(url.pathname.length > 1) micro.send(res, 404, "File not found");
	else if(url.query.f) post_file(req, res, url.query.f);
	else post_data(req, res);
}

module.exports = function(req, res) {
	logit(req, res);
	switch(req.method) {
		case 'GET': return get(req, res);
		case 'POST': return post(req, res);
	}
	return micro.send(res, 501, "Unsupported method " + req.method + "\n");
};
