var XLSX = require('./');

wbin = XLSX.readFile('/tmp/wb.xlsx', {type: "xlsx"});

XLSX.writeFile(wbin, '/tmp/wb2.xlsx', {
  defaultCellStyle: { font: { name: "Verdana", sz: 11, color: "FF00FF88"}, fill: {fgColor: {rgb: "FFFFAA00"}}}
});

