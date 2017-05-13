import xlsx = require('xlsx');

const options: xlsx.ParsingOptions = {
    cellDates: true
};

const workbook = xlsx.readFile('test.xlsx', options);
const otherworkbook = xlsx.readFile('test.xlsx', {type: 'file'});

console.log(workbook.Props.Author);

const firstsheet: string = workbook.SheetNames[0];

const firstworksheet = workbook.Sheets[firstsheet];

console.log(firstworksheet["A1"]);

interface Tester {
    name: string;
    age: number;
}

const jsonvalues: Tester[] = xlsx.utils.sheet_to_json<Tester>(firstworksheet);
const csv = xlsx.utils.sheet_to_csv(firstworksheet);
const formulae = xlsx.utils.sheet_to_formulae(firstworksheet);
