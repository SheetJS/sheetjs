var X = require('./');
var opts = { cellNF: true,
  type: 'file',
  cellHTML: true,
  cellFormula: true,
  cellStyles: false,
  cellDates: false,
  sheetStubs: false,
  sheetRows: 0,
  bookDeps: false,
  bookSheets: false,
  bookProps: false,
  bookFiles: false,
  bookVBA: false,
  WTF: false }
;
var FILENAME = './test_files/number_format_entities-2.xlsx';
wb = X.read(X.write(X.readFile(FILENAME,opts), {type:"buffer", bookType:'xlsx'}), {WTF:true, cellNF: true})

X.writeFile(wb,'/tmp/wb3.xlsx');
