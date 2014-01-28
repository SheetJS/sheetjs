# xlsx

Currently a parser for XLSX/XLSM/XLSB files.  Cleanroom implementation from the 
ISO 29500  Office Open XML specifications, [MS-XLSB], and related documents.

## Installation

In node:

    npm install xlsx

In the browser:

    <!-- JSZip must be included before xlsx.js -->
    <script lang="javascript" src="/path/to/jszip.js"></script>
    <script lang="javascript" src="/path/to/xlsx.js"></script>

## Usage

The node version installs a binary `xlsx2csv` which can read XLSX/XLSM/XLSB files and output the contents in various formats.  The source is available at `xlsx2csv.njs` in the bin directory.

See <http://oss.sheetjs.com/js-xlsx/> for a browser example. 

Note that older versions of IE does not support HTML5 File API, so the base64 mode is provided for testing.  On OSX you can get the base64 encoding by running:

    $ <target_file.xlsx base64 | pbcopy # the pbcopy puts the content in the clipboard

Simple usage (walks through every cell of every sheet and dumps the values):

    var XLSX = require('xlsx')
    var xlsx = XLSX.readFile('test.xlsx');
    var sheet_name_list = xlsx.SheetNames;
    xlsx.SheetNames.forEach(function(y) {
      for (z in xlsx.Sheets[y]) {
        if(z[0] === '!') continue;
        console.log(y + "!" + z + "=" + JSON.stringify(xlsx.Sheets[y][z].v));
      }
    });

Some helper functions in `XLSX.utils` generate different views of the sheets:

- `XLSX.utils.sheet_to_csv` generates CSV 
- `XLSX.utils.sheet_to_row_object_array` interprets sheets as tables with a header column and generates an array of objects
- `XLSX.utils.get_formulae` generates a list of formulae

## Notes 

`.SheetNames` is an ordered list of the sheets in the workbook

`.Sheets[sheetname]` returns a data structure representing the sheet.  Each key
that does not start with `!` corresponds to a cell (using `A-1` notation).  

`.Sheets[sheetname][address].v` returns the value of the specified cell and `.Sheets[sheetname][address].t` returns the type of the cell (constrained to the enumeration `ST_CellType` as documented in page 4215 of ISO/IEC 29500-1:2012(E) ) 

For more details:

- `bin/xlsx2csv.njs` is a tool for node
- `index.html` is the live demo
- `bits/90_utils.js` contains the logic for generating CSV and JSON from sheets

## Tested Environments

 - Node 0.8.14, 0.10.1
 - IE 6/7/8/9/10 using Base64 mode (IE10 using HTML5 mode)
 - FF 18 using Base64 or HTML5 mode
 - Chrome 24 using Base64 or HTML5 mode

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xlsx> for XLSX module in node
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS* modules
 - <https://saucelabs.com/u/sheetjs> for XLS* modules using Sauce Labs 

## Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

## XLS Support

XLS is available in [js-xls](https://github.com/SheetJS/js-xls).

## License

Please consult the attached LICENSE file for details.  All rights not explicitly granted by the Apache 2.0 license are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of the Microsoft Open Specifications Promise, falling under the same terms as OpenOffice (which is governed by the Apache License v2).  Given the vagaries of the promise, the Original Author makes no legal claim that in fact end users are protected from future actions.  It is highly recommended that, for commercial uses, you consult a lawyer before proceeding.

## References

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

OSP-covered specifications:

 - [MS-XLSX]: Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File Format
 - [MS-XLSB]: Excel (.xlsb) Binary File Format

## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-xlsx.png?branch=master)](https://travis-ci.org/SheetJS/js-xlsx)

[![Coverage Status](https://coveralls.io/repos/SheetJS/js-xlsx/badge.png?branch=master)](https://coveralls.io/r/SheetJS/js-xlsx?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/ed5bb2c4c4346a474fef270f847f3f78 "githalytics.com")](http://githalytics.com/SheetJS/js-xlsx)

