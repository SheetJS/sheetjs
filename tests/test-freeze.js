var XLSX = require('../.');

var JSZip = require('jszip');
var fs = require('fs');
var cheerio = require('cheerio');

var assert = require('assert');

function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var defaultCellStyle = { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};
//"A1": {v: "Header", s: { border: { top: { style: 'medium', color: { rgb: "FFFFAA00"}}, left: { style: 'medium', color: { rgb: "FFFFAA00"}} }}},


var workbook = {
  SheetNames : ["Sheet1"],
  Sheets: {
    "Sheet1": {
      "A1": {v: "Header"},
      "A2": {v: "Anchorage"},
      "A3": {v: "Anchorage"},
      "A4": {v: "Boston"},
      "A5": {v: "Chicago"},
      "A6": {v: "Dayton"},
      "A7": {v: "East Lansing"},
      "A8": {v: "Fargo"},
      "A9": {v: "Galena"},
      "A10": {v: "Iowa City"},
      "A11": {v: "Jacksonville"},
      "A12": {v: "Jacksonville"},
      "A13": {v: "Jacksonville"},
      "A14": {v: "Jacksonville"},
      "A15": {v: "Jacksonville"},
      "A16": {v: "Jacksonville"},
      "A17": {v: "Jacksonville"},
      "A18": {v: "Jacksonville"},
      "A19": {v: "Jacksonville"},
      "A20": {v: "Jacksonville"},
      "A21": {v: "Jacksonville"},
      "A22": {v: "Jacksonville"},
      "A23": {v: "Jacksonville"},
      "A24": {v: "Jacksonville"},
      "A25": {v: "Jacksonville"},
      "A26": {v: "Jacksonville"},
      "A27": {v: "Jacksonville"},
      "A28": {v: "Jacksonville"},
      "A29": {v: "Jacksonville"},
      "A30": {v: "Jacksonville"},
      "A31": {v: "Jacksonville"},
      "A32": {v: "Jacksonville"},
      "A33": {v: "Jacksonville"},
      "A34": {v: "Jacksonville"},
      "A35": {v: "Jacksonville"},
      "A36": {v: "Jacksonville"},
      "A37": {v: "Jacksonville"},
      "A38": {v: "Jacksonville"},
      "A39": {v: "Jacksonville"},
      "A40": {v: "Jacksonville"},
      "A41": {v: "Jacksonville"},
      "A42": {v: "Jacksonville"},
      "A43": {v: "Jacksonville"},
      "A44": {v: "Jacksonville"},
      "A45": {v: "Jacksonville"},
      "A46": {v: "Jacksonville"},
      "A47": {v: "Jacksonville"},
      "A48": {v: "Jacksonville"},
      "A49": {v: "Jacksonville"},
      "A50": {v: "Jacksonville"},
      "A51": {v: "Jacksonville"},
      "A52": {v: "Jacksonville"},
      "A53": {v: "Jacksonville"},
      "A54": {v: "Jacksonville"},
      "A55": {v: "Jacksonville"},
      "A56": {v: "Jacksonville"},
      "A57": {v: "Jacksonville"},
      "A58": {v: "Jacksonville"},
      "A59": {v: "Jacksonville"},
      "!ref":"A1:A59",
      "!printHeader":[1,1],
      "!freeze":{
        xSplit: "1",
        ySplit: "1",
        topLeftCell: "B2",
        activePane: "bottomRight",
        state: "frozen",
      }
    }
  }
};

describe('repeats header', function () {
  it ('repeats header', function() {


    var OUTFILE = '/tmp/freeze.xlsx';
    var OUTFILE = './lab/freeze/freeze.xlsx'

    // write the file and read it back...
    XLSX.writeFile(workbook, OUTFILE, {bookType: 'xlsx', bookSST: false});
    console.log("open \""+OUTFILE+"\"")
  });
});

