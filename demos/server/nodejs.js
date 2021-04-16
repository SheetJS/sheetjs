/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

var http = require('http');
var XLSX = require('xlsx');
var formidable = require('formidable');
var html = "";
var PORT = 3000;

var extmap = {};

var server = http.createServer(function(req, res) {
	if(req.method !== 'POST') return res.end(html);
	var form = new formidable.IncomingForm();
	form.parse(req, function(err, fields, files) {
		var f = files[Object.keys(files)[0]];
		var wb = XLSX.readFile(f.path);
		var ext = (fields.bookType || "xlsx").toLowerCase();
		res.setHeader('Content-Disposition', 'attachment; filename="download.' + (extmap[ext] || ext) + '";');
		res.end(XLSX.write(wb, {type:"buffer", bookType:ext}));
	});
}).listen(PORT);

html = [
'<pre>',
'<h3><a href="http://sheetjs.com/">SheetJS File Converter</a></h3>',
'Upload a file to convert the contents to another format.',
'',
'<b>Form Fields</b>:',
'- bookType: output format type (defaults to "XLSX")',
'- basename: basename for output file (defaults to "download")',
'',
'<form method="POST" enctype="multipart/form-data" action="/">',
'<input type="file" id="file" name="file"/>',
'<select name="bookType">',
[
	["xlsb",  "XLSB"],
	["xlsx",  "XLSX"],
	["xlsm",  "XLSM"],
	["biff8", "BIFF8 XLS"],
	["biff5", "BIFF5 XLS"],
	["biff2", "BIFF2 XLS"],
	["xlml",  "SSML 2003"],
	["ods",   "ODS"],
	["fods",  "Flat ODS"],
	["csv",   "CSV"],
	["txt",   "Unicode Text"],
	["sylk",  "Symbolic Link"],
	["html",  "HTML"],
	["dif",   "DIF"],
	["dbf",   "DBF"],
	["rtf",   "RTF"],
	["prn",   "Lotus PRN"],
	["eth",   "Ethercalc"],
].map(function(x) { return '  <option value="' + x[0] + '">' + x[1] + '</option>'; }).join("\n"),
'</select>',
'<input type="submit" value="Submit Form">',
'</form>',
'',
'<b>Form code:</b>',
'&lt;form method="POST" enctype="multipart/form-data" action="/"&gt;',
'&lt;input type="file" id="file" name="file"/&gt;',
'&lt;select name="bookType"&gt',
'&lt;!-- options here --&gt;',
'&lt;/select&gt',
'&lt;input type="submit" value="Submit Form"&gt;',
'&lt;/form&gt;',
'',
'<b>fetch Code:</b>',
'var blob = new Blob("1,2,3\\n4,5,6".split("")); // original file',
'var fd = new FormData();',
'fd.set("data", blob, "foo.bar");',
'fd.set("bookType", "xlsb");',
'var res = await fetch("/", {method:"POST", body:fd});',
'var data = await res.arrayBuffer();',
'</pre>'
].join("\n");

extmap = {
	"biff2" : "xls",
	"biff5" : "xls",
	"biff8" : "xls",
	"xlml"  : "xls"
};
console.log('listening on port ' + PORT);
