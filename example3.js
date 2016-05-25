var XLSX = require('./');

var wb = {};
wb.Sheets = {};
wb.Props = {};
wb.SSF = {};
wb.SheetNames = [];

var ws = {
  "!cols": []
};

var range = {
  s: {c: 0, r: 0},
  e: {c: 0, r: 0}
};

var cell;
for (var r = 0; r < 6; r++) {
  ws["!cols"].push({wch: 6});
  if (range.e.r < r + 1) range.e.r = r + 1;
  for (var c = 0; c < 6; c++) {
    if (range.e.c < c) range.e.c = c;
    cell_ref = XLSX.utils.encode_cell({c: c, r: r});
    cell = {v: cell_ref};
    ws[cell_ref] = cell;
  }
}

ws["!ref"] = XLSX.utils.encode_range(range);
wb.SheetNames.push("Sheet1");
wb.Sheets["Sheet1"] = ws;
wb.SheetNames.push("Sheet2");
wb.Sheets["Sheet2"] = JSON.parse(JSON.stringify(ws))
// workbook options
var wopts = {bookType: "xlsx"};


//console.log(JSON.stringify(wb, null,4))
var OUTFILE = '/tmp/example.xlsx';
XLSX.writeFile(wb, OUTFILE, wopts);
console.log("Results written to " + OUTFILE)
