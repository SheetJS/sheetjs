/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var Hapi = require('hapi'), server = new Hapi.Server();
var logit = require('./_logit');
var Worker = require('webworker-threads').Worker;
var data = "a,b,c\n1,2,3".split("\n").map(x => x.split(","));

function get_data(req, res, type) {
	var work = new Worker(function(){
		var XLSX = require('xlsx');
		this.onmessage = function(e) {
			console.log("get data " + e.data);
			var ws = XLSX.utils.aoa_to_sheet(e.data[1]);
			var wb = XLSX.utils.book_new();
			XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
			console.log("prepared wb");
			postMessage(XLSX.write(wb, {type:'binary', bookType:type}));
			console.log("sent data");
		};
	});
	work.onmessage = function(e) { console.log(e); res(e); };
	work.postMessage([type, data]);
}

var port = 7262;
server.connection({ host:'localhost', port: port});

server.route({ method: 'GET', path: '/', handler: function(req, res) {
	logit(req.raw.req, req.raw.res);
	if(req.query.t) return get_data(req, res, req.query.t);
	else if(req.query.f) return get_file(req, res, req.query.f);
	return res('Forbidden').code(403);
}});
server.route({ method: 'POST', path: '/', handler: function(req, res) {
	logit(req.raw.req, req.raw.res);
	if(req.query.f) return post_file(req, res, req.query.f);
	return post_data(req, res);
}});
server.start(function(err) {
	if(err) throw err;
	console.log('Serving HTTP on port ' + port);
});
