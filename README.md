# xlsx

Currently a parser for XLSX files.  Cleanroom implementation from the ISO 29500  Office Open XML specifications.

This has been tested on some very basic XLSX files generated from Excel 2011.

*THIS WAS WHIPPED UP VERY QUICKLY TO SATISFY A VERY SPECIFIC NEED*.  If you need something that is not currently supported, file an issue and attach a sample file.  I will get to it :)

## Installation

In node:

    npm install xlsx

In the browser:

    <script lang="javascript" src="/path/to/jszip.js"></script>
    <script lang="javascript" src="/path/to/xlsx.js"></script>

## Usage

See `xlsx2csv.njs` in the bin directory for usage in node.

See http://niggler.github.com/js-xlsx/ for a browser example.

## Notes 

`.SheetNames` is an ordered list of the sheets in the workbook

`.Sheets[sheetname]` returns a data structure representing the sheet.  Each key
that does not start with `!` corresponds to a cell (using `A-1` notation).  

`.Sheets[sheetname][address].v` returns the value of the cell and `.Sheets[sheetname][address].t` returns the type of the cell (constrained to the enumeration `ST_CellType` as documented in page 4215 of ISO/IEC 29500-1:2012(E) ) 

Simple usage:

    var XLSX = require('xlsx')
    var xlsx = XLSX.readFile('test.xlsx');
    var sheet_name_list = xlsx.SheetNames;
    xlsx.SheetNames.forEach(function(y) {
      for (z in zip.Sheets[y]) {
        if(z[0] === '!') continue;
        console.log(y + "!" + z + "=" + JSON.stringify(zip.Sheets[y][z].v));
      }
    });

## License

Please consult the attached LICENSE file for details.  All rights not explicitly granted by the MIT license are reserved by the Original Author.

## References

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

