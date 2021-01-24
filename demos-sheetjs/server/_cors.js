/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var cors = function(req, res) { res.header('Access-Control-Allow-Origin', '*'); };
cors.mw = function(req, res, next) { cors(req, res); next(); };
module.exports = cors;
