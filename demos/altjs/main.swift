/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */

let sheetjs = try SheetJSCore();

try print(sheetjs.version());

let filenames: [[String]] = [
	["xlsx", "xlsx"],
	["xlsb", "xlsb"],
	["biff8.xls", "xls"],
	["xml.xls", "xlml"]
];

for fn in filenames {
	let wb: SJSWorkbook = try sheetjs.readFile(file: "sheetjs." + fn[0]);
	let ws: SJSWorksheet = try wb.getSheetAtIndex(idx: 0);
	let csv: String = try ws.toCSV();
    print(csv);
    let wbout: String = try wb.writeBStr(bookType: fn[1]);
    try wbout.write(toFile: "sheetjsswift." + fn[0], atomically: false, encoding: String.Encoding.isoLatin1);
}
