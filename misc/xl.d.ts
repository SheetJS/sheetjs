interface Cell {
  v;
  w?: string;
  t?: string;
  f?: string;
  r?: string;
  h?: string;
  c?: any;
  z?: string;
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
