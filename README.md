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

## Tested Environments

 - Node >= 0.8.14 (untested in versions before 0.8.14)
 - IE 6/7/8/9/10 using Base64 mode
 - FF 18 using Base64 or HTML5 mode
 - Chrome 24 using Base64 or HTML5 mode

## Usage

See `xlsx2csv.njs` in the bin directory for usage in node.

See http://niggler.github.com/js-xlsx/ for a browser example. 

Note that IE does not support HTML5 File API, so the base64 mode is provided for testing.  On OSX you can get the base64 encoding by running:

    $ <target_file.xlsx base64 | pbcopy # the pbcopy puts the content in the clipboard

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
      for (z in xlsx.Sheets[y]) {
        if(z[0] === '!') continue;
        console.log(y + "!" + z + "=" + JSON.stringify(xlsx.Sheets[y][z].v));
      }
    });

## License

Please consult the attached LICENSE file for details.  All rights not explicitly granted by the MIT license are reserved by the Original Author.

## XLS Support

XLS is not supported in this module.  Due to Licensing issues [that are discussed in more detail elsewhere](https://github.com/Niggler/js-xls/issues/1#issuecomment-13852286), the implementation cannot be released in a GPL or MIT-style license.  If you need XLS support, consult [my js-xls project](https://github.com/Niggler/js-xls).

## References

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

