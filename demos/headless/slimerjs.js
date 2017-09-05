/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var fs = require('fs');
var xlsx = require('../../dist/xlsx.full.min');
var page = require('webpage').create();

page.open('http://www.google.com', function(status) {

  var data = fs.read('sheetjs.xlsx', {mode: 'rb', charset: 'utf8'});
  var workbook = xlsx.read(data, {type: 'binary'});
  data = xlsx.utils.sheet_to_csv(workbook.Sheets['SheetJS']);
  console.log("Data: " + data);

  phantom.exit();
});

