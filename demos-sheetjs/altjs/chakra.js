/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
var wb = XLSX.read(payload, {type:'base64'});
console.log(XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]));
