#target illustrator
﻿var thisFile = new File($.fileName);
var basePath = thisFile.path;

#include "xlsx.extendscript.js";

var filename = "/sheetjs.xlsx";

/* Read file from disk */
var workbook = XLSX.readFile(basePath + filename);

/* Display first worksheet */
var first_sheet_name = workbook.SheetNames[0], first_worksheet = workbook.Sheets[first_sheet_name];
var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});
alert(data);

var outfmts = [
  ["xlsb",  "testw.xlsb"],
  ["biff8", "testw.xls"],
  ["xlml",  "testw.xml"],
  ["fods",  "testw.fods"],
  ["csv",   "testw.csv"],
  ["txt",   "testw.txt"],
  ["slk",   "testw.slk"],
  ["eth",   "testw.eth"],
  ["htm",   "testw.htm"],
  ["dif",   "testw.dif"],
  ["ods",   "testw.ods"],
  ["xlsx",  "testw.xlsx"]
];
for(var i = 0; i < outfmts.length; ++i) {
  alert(outfmts[i][0]);
  var fname = basePath + "/" + outfmts[i][1];

  /* Write file to disk */
  XLSX.writeFile(workbook, fname);

  /* Read new file */
  var wb = XLSX.readFile(fname, {cellDates:true});

  /* Display first worksheet */
  var f_sheet_name = wb.SheetNames[0], f_worksheet = wb.Sheets[f_sheet_name];
  var data = XLSX.utils.sheet_to_json(f_worksheet, {header:1, cellDates:true});
  alert(data);
}
