var XLSX = require('./');
var sheetName = 'test &amp; debug';


  var ATTRIBUTE_VALUE_STYLE={
      font: {
      name:   "Arial",
      sz:     10
      }
  };

  var Workbook = function(){
      this.SheetNames = [];
      this.Sheets = {};
  };

  var range = {s: {c:10000000, r:10000000}, e:{c:0, r:0}};

  function updateRange(row, col) {
      if (range.s.r > row) { range.s.r = row;}
  if (range.s.c > col) { range.s.c = col; }
  if (range.e.r < row) { range.e.r = row; }
  if (range.e.c < col) { range.e.c = col; }
    }

    function addCell(wb, ws, value, type, row, col, styles) {

        updateRange(row, col);

        var cell = {t: type, v: value, s:styles};

    // i use d to recognize that the format is a date, and if it is, i use z attribute to format it
    if (cell.t === 'd') {
        cell.t = 'n';
        cell.z = XLSX.SSF._table[14];
        }

    var cell_ref = XLSX.utils.encode_cell({c: col, r:row});

    ws[cell_ref] = cell;
    }

    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
        }

    function datenum(v, date1904) {
        if(date1904) v+=1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        }

    var wb = new Workbook();
    var ws = {};

    /* Here i add a cell with format number with value 10*/
    addCell(wb, ws, 10, 'n', 0, 0, ATTRIBUTE_VALUE_STYLE);

/* Here i add a cell with format number that is current date*/
addCell(wb, ws, datenum(new Date()), 'd', 0, 1, ATTRIBUTE_VALUE_STYLE);

/* Here i add a cell with format number with value 10, again another number*/
addCell(wb, ws, 10, 'd', 0, 2, ATTRIBUTE_VALUE_STYLE);

/* Here i add a cell with format number with value 10, again another number*/
addCell(wb, ws, "Hello null\u0000world", 's', 0, 3);


ws['!ref'] = XLSX.utils.encode_range(range);

wb.SheetNames.push(sheetName);
wb.Sheets[sheetName] = ws;


ws['!rowBreaks'] = [12,24];
ws['!colBreaks'] = [3,6];
ws['!pageSetup'] = {scale: '140'};

/* bookType can be 'xszlsx' or 'xlsm' or 'xlsb' */
var defaultCellStyle =  { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}};
var wopts = { bookType:'xlsx', bookSST:false, type:'binary', defaultCellStyle: defaultCellStyle, showGridLines: false};

//console.log(JSON.stringify(wb, null,4))
var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(wb, OUTFILE, wopts);
console.log("Results written to " + OUTFILE)
