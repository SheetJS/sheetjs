/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */

import * as XLSX from 'xlsx';

const { read, utils: { sheet_to_json } } = XLSX;

export function readFirstSheet(data:any, options:XLSX.ParsingOptions): any[][] {
	const wb: XLSX.WorkBook = read(data, options);
	const ws: XLSX.WorkSheet = wb.Sheets[wb.SheetNames[0]];
	return sheet_to_json(ws, {header:1, raw:true});
};
