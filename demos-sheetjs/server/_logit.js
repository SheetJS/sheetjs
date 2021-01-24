/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var sprintf = require('printj').sprintf;
var logit = function(req, res) {
	console.log(sprintf("%s %s %d", req.method, req.url, res.statusCode));
};
logit.mw = function(req, res, next) { logit(req, res); next(); }
module.exports = logit;
