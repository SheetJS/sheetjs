/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var Hapi = require('hapi'), server = new Hapi.Server();
var logit = require('./_logit');
var Worker = require('tiny-worker');
var fs = require('fs');
var data = "a,b,c\n1,2,3".split("\n").map(x => x.split(","));

function get_data(req, res, type) {
	var work = new Worker('worker.js');
	work.onmessage = function(e) {
		if(e.data.err) console.log(e.data.err);
		return res(e.data.data);
	};
	work.postMessage({action:"write", type:type, data:data});
}

function get_file(req, res, file) {
	var work = new Worker('worker.js');
	work.onmessage = function(e) {
		fs.writeFileSync(file, e.data.data, 'binary');
		return res("wrote to " + file + "\n");
	};
	work.postMessage({action:"write", file:file, data:data});
}

function post_file(req, res, file) {
	var work = new Worker('worker.js');
	work.onmessage = function(e) {
		data = e.data.data;
		return res("read from " + file + "\n");
	};
	work.postMessage({action:"read", file:file});
}

function post_data(req, res, type) {
	var keys = Object.keys(req.payload), k = keys[0];
	post_file(req, res, req.payload[k].path);
}

var port = 7262;
server.connection({ host:'localhost', port: port});

server.route({ method: 'GET', path: '/',
handler: function(req, res) {
	logit(req.raw.req, req.raw.res);
	if(req.query.t) get_data(req, res, req.query.t);
	else if(req.query.f) get_file(req, res, req.query.f);
	else res('Forbidden').code(403);
}});
server.route({ method: 'POST', path: '/',
config:{payload:{ output: 'file', parse: true, allow: 'multipart/form-data'}},
handler: function(req, res) {
	logit(req.raw.req, req.raw.res);
	if(req.query.f) return post_file(req, res, req.query.f);
	return post_data(req, res);
}});
server.route({ method: 'POST', path: '/file',
handler: function(req, res) {
	logit(req.raw.req, req.raw.res);
	if(req.query.f) return post_file(req, res, req.query.f);
	return post_data(req, res);
}});
server.start(function(err) {
	if(err) throw err;
	console.log('Serving HTTP on port ' + port);
});
