#!/usr/bin/env node
var fs = require('fs');
var paths = fs.readFileSync('tests/fixtures.lst','utf-8').replace(/\r/g,"").split("\n");
var aux = [
	'multiformat.lst',
	'./misc/ssf.json',
	'./test_files/biff5/number_format_greek.xls'
]
var fullpaths = paths.concat(aux);
fs.writeFileSync('tests/fixtures.js',
  fullpaths.map(function(x) {
    return [x, fs.existsSync(x) ? fs.readFileSync(x).toString('base64') : ""]
  }).map(function(w) {
    return "fs['" + w[0] + "'] = '" + w[1] + "';\n";
  }).join("")
);
