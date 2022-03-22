/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
// @deno-types="https://unpkg.com/xlsx/types/index.d.ts"
import * as XLSX from 'https://unpkg.com/xlsx/xlsx.mjs';
import * as cptable from 'https://unpkg.com/xlsx/dist/cpexcel.full.mjs';
XLSX.set_cptable(cptable);

const filename = Deno.args[0];
if(!filename) {
	console.error("usage: sheet2csv <filename> [sheetname]");
	Deno.exit(1);
}

const workbook = XLSX.readFile(filename);
const sheetname = Deno.args[1] || workbook.SheetNames[0];

if(!workbook.Sheets[sheetname]) {
	console.error(`error: workbook missing sheet ${sheetname}`);
	Deno.exit(1);
}

console.log(XLSX.utils.sheet_to_csv(workbook.Sheets[sheetname]));
