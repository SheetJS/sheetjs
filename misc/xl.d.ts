interface Cell {
  v;
  t: string;
  ixfe: number;
}

interface Worksheet {
  [key: string]: Cell;
}

interface Worksheets {
  [key: string]: Worksheet;
}

interface Workbook {
  SheetNames: string[];
  Sheets: Worksheets;
}

interface XLSX {
  verbose: Number;
  readFile(filename: string): Workbook; 
  utils: {
    get_formulae(worksheet: Worksheet): string[];
    make_csv(worksheet: Worksheet): string;
  };
}
