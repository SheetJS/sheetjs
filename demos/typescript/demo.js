var readFirstSheet = require("./").readFirstSheet;
console.log(readFirstSheet("../../sheetjs.xlsb", {type:"file", cellDates:true}));