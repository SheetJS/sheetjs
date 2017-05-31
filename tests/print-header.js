var XLSX = require('../.');

var JSZip = require('jszip');
var fs = require('fs');
var cheerio = require('cheerio');

var assert = require('assert');

function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var defaultCellStyle = { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};



describe('repeats header', function () {
  it ('repeats header', function() {


    var workbook = {
      SheetNames: ["Sheet1"],
      Sheets: {
        "Sheet1": {
          "!ref":"A1:Z99",
          "!printHeader":[1,1],
          "!printColumns":["A","C"]
        }
      }
    }

    "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split('').forEach(function(c) {
      for (var i=1; i<100; i++) {
        var address = c + i;

        workbook.Sheets.Sheet1[address] = {v: address};
      }
    })
    var OUTFILE = '/tmp/header.xlsx';
    var OUTFILE = __dirname + '/../lab/headers/header.xlsx';


    // write the file and read it back...
    XLSX.writeFile(workbook, OUTFILE, {bookType: 'xlsx', bookSST: false});
    console.log("open \""+OUTFILE+"\"")
  });
});

