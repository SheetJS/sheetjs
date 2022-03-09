/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
/* eslint no-console:0 */
/*jshint node:true */
var X = require('ssf');
function run() {
  var argv = process.argv.slice(2);
  if (argv.length < 2 || argv[0] == "-h" || argv[0] == "--help") {
    console.error("usage: ssf-cli <format> <value>");
    console.error("output: format_as_string|format_as_number|");
    process.exit(0);
  }
  console.log(X.format(argv[0], argv[1]) + "|" + X.format(argv[0], +(argv[1])) + "|");
}

module.exports = run;
