/**
 * Test TypeScript defs
 * (not compiled into js, only used to test def resolution)
 */
import * as XLSX from "xlsx-js-style";

//const wb:XLSX.WorkBook = []
const wb = XLSX.utils.book_new();
let row1: XLSX.CellObject[] = ["a", "b", "c"];
let row2: XLSX.CellObject[] = [1, 2, 3];

const ws = XLSX.utils.aoa_to_sheet([row1, row2]);
XLSX.utils.book_append_sheet(wb, ws, "SheetJS");

/* generate XLSX file and send to client */
XLSX.writeFile(wb, "sheetjs.xlsx");
