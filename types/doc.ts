import * as XLSX from 'xlsx';
import * as fs from 'fs';

const version: string = XLSX.version;

const SSF = XLSX.SSF;

let read_opts: XLSX.ParsingOptions = {
	type: "buffer",
	raw: false,
	cellFormula: false,
	cellHTML: false,
	cellNF: false,
	cellStyles: false,
	cellText: false,
	cellDates: false,
	dateNF: "yyyy-mm-dd",
	sheetStubs: false,
	sheetRows: 3,
	bookDeps: false,
	bookFiles: false,
	bookProps: false,
	bookSheets: false,
	bookVBA: false,
	password: "",
	WTF: false
};

let write_opts: XLSX.WritingOptions = {
	type: "buffer",
	cellDates: false,
	bookSST: false,
	bookType: "xlsx",
	sheet: "Sheet1",
	compression: false,
	Props: {
		Author: "Someone",
		Company: "SheetJS LLC"
	}
};

const wb1 = XLSX.readFile("sheetjs.xls", read_opts);
XLSX.writeFile(wb1, "sheetjs.new.xlsx", write_opts);

read_opts.type = "binary";
const wb2 = XLSX.read("1,2,3\n4,5,6", read_opts);
write_opts.type = "binary";
const out2 = XLSX.write(wb2, write_opts);

read_opts.type = "buffer";
const wb3 = XLSX.read(fs.readFileSync("sheetjs.xlsx"), read_opts);
write_opts.type = "base64";
const out3 = XLSX.write(wb3, write_opts);

const ws1 = XLSX.utils.aoa_to_sheet([
    "SheetJS".split(""),
    [1,2,3,4,5,6,7],
    [2,3,4,5,6,7,8]
], {
	dateNF: "yyyy-mm-dd",
	cellDates: true,
	sheetStubs: false
});

const ws2 = XLSX.utils.json_to_sheet([
    {S:1,h:2,e:3,e_1:4,t:5,J:6,S_1:7},
    {S:2,h:3,e:4,e_1:5,t:6,J:7,S_1:8}
], {
	header:["S","h","e","e_1","t","J","S_1"],
	cellDates: true,
	dateNF: "yyyy-mm-dd"
});

const tbl = {}; /* document.getElementById('table'); */
const ws3 = XLSX.utils.table_to_sheet(tbl, {
	raw: true,
	cellDates: true,
	dateNF: "yyyy-mm-dd"
});

const obj1 = XLSX.utils.sheet_to_formulae(ws1);

const str1: string = XLSX.utils.sheet_to_csv(ws2, {
	FS: "\t",
	RS: "|",
	dateNF: "yyyy-mm-dd",
	strip: true,
	blankrows: true,
	skipHidden: true
});

const html1: string = XLSX.utils.sheet_to_html(ws3, {
	editable: false
});

const arr1: object[] = XLSX.utils.sheet_to_json(ws1, {
	raw: true,
	range: 1,
	header: "A",
	dateNF: "yyyy-mm-dd",
	defval: 0,
	blankrows: true
});

const arr2: any[][] = XLSX.utils.sheet_to_json<any[][]>(ws2, {
	header: 1
});

const arr3: any[] = XLSX.utils.sheet_to_json(ws3, {
	header: ["Sheet", "JS", "Rocks"]
});
