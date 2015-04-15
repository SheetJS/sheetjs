var XLSX = require('./');
var Workbook = require('../workbook');

///http://daveaddey.com/?p=40
function JSDateToExcelDate(inDate) {
  return 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
}

var workbook = new Workbook(XLSX)
    .addRowsToSheet("Main", [
      [
        {
          v: "This is a submerged cell",
          s:{
            border: {
              left: {style: 'thick', color: {auto: 1}},
              top: {style: 'thick', color: {auto: 1}},
              bottom: {style: 'thick', color: {auto: 1}}
            }
          }
        },
        {
          v: "Pirate ship",
          s:{
            border: {
              top: {style: 'thick', color: {auto: 1}},
              bottom: {style: 'thick', color: {auto: 1}}
            }
          }
        },
        {
          v: "Sunken treasure",
          s:{
            border: {
              right: {style: 'thick', color: {auto: 1}},
              top: {style: 'thick', color: {auto: 1}},
              bottom: {style: 'thick', color: {auto: 1}}
            }
          }
        }],
      [
        {"v": "Blank"},
        {"v": "Red", "s": {fill: { fgColor: { rgb: "FFFF0000"}}}},
        {"v": "Green", "s": {fill: { fgColor: { rgb: "FF00FF00"}}}},
        {"v": "Blue", "s": {fill: { fgColor: { rgb: "FF0000FF"}}}}
      ],
      [
        {"v": "Default"},
        {"v": "Arial", "s": {font: {name: "Arial", sz: 24, color: {theme: "5"}}}},
        {"v": "Times New Roman", "s": {font: {name: "Times New Roman", sz: 16, color: {rgb: "FF2222FF"}}}},
        {"v": "Courier New", "s": {font: {name: "Courier New", sz: 14}}}
      ],
      [
        0.618033989,
        {"v": 0.618033989},
        {"v": 0.618033989, "t": "n"},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.0%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.00%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.000%"}},
        {"v": 0.618033989, "t": "n", "s": { "numFmt": "0.0000%"}},
        {"v": 0, "t": "n", "s": { numFmt: "0.00%;\\(0.00%\\);\\-;@"}, fill: { fgColor: { rgb: "FFFFCC00"}}}
      ],
      [
        {v: (new Date()).toLocaleString()},
        {v: JSDateToExcelDate(new Date()), t: 'd'},
        {v: JSDateToExcelDate(new Date()),  s: {numFmt: 'd-mmm-yy'}}
      ]
      ,
      [
        {v: "left", "s": { alignment: {horizontal: "left"}}},
        {v: "left", "s": { alignment: {horizontal: "center"}}},
        {v: "left", "s": { alignment: {horizontal: "right"}}}
      ],[
        {v: "vertical", "s": { alignment: {vertical: "top"}}},
        {v: "vertical", "s": { alignment: {vertical: "center"}}},
        {v: "vertical", "s": { alignment: {vertical: "bottom"}}}
      ],[
        {v: "indent", "s": { alignment: {indent: "1"}}},
        {v: "indent", "s": { alignment: {indent: "2"}}},
        {v: "indent", "s": { alignment: {indent: "3"}}}
      ],
        [{
          v: "In publishing and graphic design, lorem ipsum is a filler text commonly used to demonstrate the graphic elements of a document or visual presentation. ",
          s: { alignment: { wrapText: 1, alignment: 'right', vertical: 'center', indent: 1}}
         }
        ],
        [
          {v: 41684.35264774306, s: {numFmt: 'm/d/yy'}},
          {v: 41684.35264774306, s: {numFmt: 'd-mmm-yy'}},
          {v: 41684.35264774306, s: {numFmt: 'h:mm:ss AM/PM'}},
          {v: JSDateToExcelDate(new Date()),  s: {numFmt: 'm/d/yy'}},
          {v: 42065.02247239584,  s: {numFmt: 'm/d/yy'}},
          {v: JSDateToExcelDate(new Date()),  s: {numFmt: 'm/d/yy h:mm:ss AM/PM'}}
        ],
        [
          {v: "Apple", s: {border: {top: { style: "thin"}, left: { style: "thin"}, right: { style: "thin"}, bottom: { style: "thin"}}}},
          {},
          {
            v: "Apple",
            s: {
              border: {
                diagonalUp: 1, diagonalDown: 1,
                top: { style: "dashed", color: {auto: 1}},
                right: { style: "medium", color: {theme: "5"}},
                bottom: { style: "hair", color: {theme: 5, tint: "-0.3"}},
                left: { style: "thin", color: {rgb: "FFFFAA00"}},
                diagonal: {style: "dotted", color: {auto: 1}}
              }
            }
          },
          {},
          {
              v: "Pear",
              s: {
                border: {
                  diagonalUp: 1, diagonalDown: 1,
                  top: { style: "dashed", color: {auto: 1}},
                  right: { style: "dotted", color: {theme: "5"}},
                  bottom: { style: "mediumDashed", color: {theme: 5, tint: "-0.3"}},
                  left: { style: "double", color: {rgb: "FFFFAA00"}},
                  diagonal: {style: "hair", color: {auto: 1}}
                }
            }
          }
        ],
        [
          {v: "Up 90", s: {alignment: {textRotation: 90}}},
          {v: "Up 45", s: {alignment: {textRotation: 45}}},
          {v: "Horizontal", s: {alignment: {textRotation: 0}}},
          {v: "Down 45", s: {alignment: {textRotation: 135}}},
          {v: "Down 90", s: {alignment: {textRotation: 180}}},
          {v: "Vertical", s: {alignment: {textRotation: 255}}}
        ],
      [
        {v: "Font color test", s: { font: {fgColor: {rgb: "FFC6EFCE"}}}}
      ]
    ]).mergeCells("Main", {
      "s": {"c": 0, "r": 0 },
      "e": {"c": 2, "r": 0 }
    }).finalize();


var OUTFILE = '/tmp/wb.xlsx';
XLSX.writeFile(workbook, OUTFILE, {defaultCellStyle: { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}}});
console.log("Results written to " + OUTFILE)
