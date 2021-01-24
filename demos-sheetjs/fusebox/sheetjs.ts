import * as XLSX from 'xlsx';

console.log(XLSX.version);

const bookType: string = "xlsb";
const fn: string = "sheetjsfbox." + bookType
const sn: string = "SheetJSFBox";
const aoa: any[][] = [ ["Sheet", "JS"], ["Fuse", "Box"], [72, 62] ];


var wb: XLSX.WorkBook = XLSX.utils.book_new();
var ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(aoa);
XLSX.utils.book_append_sheet(wb, ws, sn);

var payload: string = "";
var w2: XLSX.WorkBook;
if(typeof process != 'undefined' && process.versions && process.versions.node) {
	/* server */
	XLSX.writeFile(wb, fn);
	w2 = XLSX.readFile(fn)
} else {
	/* client */
	payload = XLSX.write(wb, {bookType: "xlsb", type:"binary"});
	w2 = XLSX.read(payload, {type:"binary"});
}

var s2: XLSX.WorkSheet = w2.Sheets[sn];
console.log(XLSX.utils.sheet_to_csv(s2));
