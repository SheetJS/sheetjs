import XLSX = require('xlsx');

const options: XLSX.ParsingOptions = {
    cellDates: true
};

const workbook = XLSX.readFile('test.xlsx', options);
const otherworkbook = XLSX.readFile('test.xlsx', {type: 'file'});

console.log(workbook.Props.Author);

const firstsheet: string = workbook.SheetNames[0];

const firstworksheet = workbook.Sheets[firstsheet];

console.log(firstworksheet["A1"]);

interface Tester {
    name: string;
    age: number;
}

const jsonvalues: Tester[] = XLSX.utils.sheet_to_json<Tester>(firstworksheet);
const csv = XLSX.utils.sheet_to_csv(firstworksheet);
const formulae = XLSX.utils.sheet_to_formulae(firstworksheet);
