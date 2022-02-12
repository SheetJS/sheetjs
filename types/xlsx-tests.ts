import XLSX = require('xlsx');

const options: XLSX.ParsingOptions = {
    cellDates: true
};

const workbook: XLSX.WorkBook = XLSX.readFile('test.xlsx', options);
const otherworkbook: XLSX.WorkBook = XLSX.readFile('test.xlsx', {type: 'file'});

const author: string = workbook.Props.Author;

const firstsheet: string = workbook.SheetNames[0];
const firstworksheet: XLSX.WorkSheet = workbook.Sheets[firstsheet];
const WB1A1: XLSX.CellObject = (firstworksheet["A1"]);

interface Tester {
    name: string;
    age: number;
}

const jsonvalues: Tester[] = XLSX.utils.sheet_to_json<Tester>(firstworksheet);
const csv: string = XLSX.utils.sheet_to_csv(firstworksheet);
const txt: string = XLSX.utils.sheet_to_txt(firstworksheet);
const dif: string = XLSX.utils.sheet_to_dif(firstworksheet);
const slk: string = XLSX.utils.sheet_to_slk(firstworksheet);
const eth: string = XLSX.utils.sheet_to_eth(firstworksheet);
const formulae: string[] = XLSX.utils.sheet_to_formulae(firstworksheet);
const aoa: any[][] = XLSX.utils.sheet_to_json<any[]>(firstworksheet, {raw:true, header:1});

const aoa2: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet<number>([
    [1,2,3,4,5,6,7],
    [2,3,4,5,6,7,8]
]);

const js2ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet<Tester>([
    {name:"Sheet", age: 12},
    {name:"JS", age: 24}
]);

const WBProps = workbook.Workbook;
const WBSheets = WBProps.Sheets;
const WBSheet0 = WBSheets[0];
console.log(WBSheet0.Hidden);

const fmt14 = XLSX.SSF.get_table()[14];
XLSX.SSF.load('"This is a custom format "0.000');

const newwb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newwb, aoa2, "AOA");
XLSX.utils.book_append_sheet(newwb, js2ws, "JSON");
const bstrxlsx: string = XLSX.write(newwb, {type: "binary", bookType: "xlsx" });

const wb_1: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "base64", bookType: "xlsx" }), {type: "base64"});
const wb_2: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "binary", bookType: "xlsx" }), {type: "binary"});
const wb_3: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "buffer", bookType: "xlsx" }), {type: "buffer"});
const wb_4: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "file", bookType: "xlsx" }), {type: "file"});
const wb_5: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "array", bookType: "xlsx" }), {type: "array"});
const wb_6: XLSX.WorkBook = XLSX.read(XLSX.write(newwb, {type: "string", bookType: "xlsx" }), {type: "string"});

function get_header_row(sheet: XLSX.WorkSheet) {
    const headers: string[] = [];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const R: number = range.s.r;
    for(let C = range.s.c; C <= range.e.c; ++C) {
        const cell: XLSX.CellObject = sheet[XLSX.utils.encode_cell({c:C, r:R})];
        let hdr = "UNKNOWN " + C;
        if(cell && cell.t) hdr = XLSX.utils.format_cell(cell);
        headers.push(hdr);
    }
    return headers;
}

const headers: string[] = get_header_row(aoa2);

const CFB = XLSX.CFB;
const vbawb = XLSX.readFile("test.xlsm", {bookVBA:true});
if(vbawb.vbaraw) {
    const cfb: any /* XLSX.CFB.CFB$Container */ = CFB.read(vbawb.vbaraw, {type: "buffer"});
}
