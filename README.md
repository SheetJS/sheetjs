# xlsx

Parser and writer for XLSX/XLSM/XLSB files.  Cleanroom implementation from the
ISO 29500  Office Open XML specifications, [MS-XLSB], and related documents.

## Installation

In [nodejs](https://www.npmjs.org/package/xlsx):

    npm install xlsx

In the browser:

    <!-- This is the only file you need (includes xlsx.js and jszip) -->
    <script lang="javascript" src="dist/xlsx.core.min.js"></script>

In [bower](http://bower.io/search/?q=js-xlsx):

    bower install js-xlsx

CDNjs automatically pulls the latest version and makes all versions available at
<http://cdnjs.com/libraries/xlsx>

Older versions of this README recommended a more explicit approach:

    <!-- JSZip must be included before xlsx.js -->
    <script lang="javascript" src="/path/to/jszip.js"></script>
    <script lang="javascript" src="/path/to/xlsx.js"></script>

## Optional Modules

The nodejs version automatically requires modules for additional features.  Some
of these modules are rather large in size and are only needed in special
circumstances, so they do not ship with the core.  For browser use, they must
be included directly:

    <!-- international support from https://github.com/sheetjs/js-codepage -->
    <script src="dist/cpexcel.js"></script>

An appropriate version for each dependency is included in the dist/ directory.

The complete single-file version is generated at `dist/xlsx.full.min.js`

## Parsing Workbooks

For parsing, the first step is to read the file.

- nodejs:

```
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

- ajax:

```
/* set up XMLHttpRequest */
var url = "test_files/formula_stress_test_ajax.xlsx";
var oReq = new XMLHttpRequest();
oReq.open("GET", url, true);
oReq.responseType = "arraybuffer";

oReq.onload = function(e) {
  var arraybuffer = oReq.response;

  /* convert data to binary string */
  var data = new Uint8Array(arraybuffer);
  var arr = new Array();
  for(var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  var bstr = arr.join("");

  /* Call XLSX */
  var workbook = XLSX.read(bstr, {type:"binary"});

  /* DO SOMETHING WITH workbook HERE */
}

oReq.send();
```

- html5 drag-and-drop using readAsBinaryString:

```
/* set up drag-and-drop event */
function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      /* if binary string, read with type 'binary' */
      var wb = XLSX.read(data, {type: 'binary'});

      /* DO SOMETHING WITH workbook HERE */
    };
    reader.readAsBinaryString(f);
  }
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

This example walks through every cell of every sheet and dumps the values:

```
var sheet_name_list = workbook.SheetNames;
sheet_name_list.forEach(function(y) {
  var worksheet = workbook.Sheets[y];
  for (z in worksheet) {
    if(z[0] === '!') continue;
    console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
  }
});
```

Complete examples:

- <http://oss.sheetjs.com/js-xlsx/> HTML5 File API / Base64 Text / Web Workers

Note that older versions of IE does not support HTML5 File API, so the base64
mode is provided for testing.  On OSX you can get the base64 encoding with:

    $ <target_file.xlsx base64 | pbcopy

- <http://oss.sheetjs.com/js-xlsx/ajax.html> XMLHttpRequest

- <https://github.com/SheetJS/js-xlsx/blob/master/bin/xlsx.njs> nodejs

The nodejs version installs a binary `xlsx` which can read XLSX/XLSM/XLSB
files and output the contents in various formats.  The source is available at
`xlsx.njs` in the bin directory.

Some helper functions in `XLSX.utils` generate different views of the sheets:

- `XLSX.utils.sheet_to_csv` generates CSV
- `XLSX.utils.sheet_to_json` generates an array of objects
- `XLSX.utils.get_formulae` generates a list of formulae

## Writing Workbooks

Assuming `workbook` is a workbook object, just call write:

- nodejs write to file:

```
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsx');
```

- write to binary string (using FileSaver.js)

```
/* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
var wopts = { bookType:'xlsx', bookSST:true, type:'binary' };

var wbout = XLSX.write(workbook,wopts);

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

saveAs(new Blob([s2ab(wbout)],{type:""}), "test.xlsx")
```

Complete examples:

- <http://sheetjs.com/demos/writexlsx.html> generates a simple file
- <http://git.io/WEK88Q> writing an array of arrays in nodejs

## Interface

`XLSX` is the exposed variable in the browser and the exported nodejs variable


`XLSX.read(data, read_opts)` attempts to parse `data`.

`XLSX.readFile(filename, read_opts)` attempts to read `filename` and parse.

`XLSX.write(wb, write_opts)` attempts to write the workbook `wb`

`XLSX.writeFile(wb, filename, write_opts)` attempts to write `wb` to `filename`

## Cell Object Description

js-xlsx conforms to the Common Spreadsheet Format (CSF):

`.SheetNames` is an ordered list of the sheets in the workbook

`.Sheets[sheetname]` returns a data structure representing the sheet.  Each key
that does not start with `!` corresponds to a cell (using `A-1` notation).

`.Sheets[sheetname][address]` returns the specified cell:

- `.v` : the raw value of the cell
- `.w` : the formatted text of the cell (if applicable)
- `.t` : the type of the cell (constrained to the enumeration `ST_CellType` as
  documented in page 4215 of ISO/IEC 29500-1:2012(E) )
- `.f` : the formula of the cell (if applicable)
- `.r` : the rich text encoding of a cell text (if applicable)
- `.h` : an HTML rendering of the rich text (if applicable)
- `.c` : comments associated with the cell
- `.z` : the number format string associated with the cell (if requested)
- `.l` : the hyperlink of the cell (.Target holds link, .tooltip is tooltip)

For dates, `.v` holds the raw date code from the sheet and `.w` holds the text

## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| cellFormula | true    | Save formulae to the .f field |
| cellHTML    | true    | Parse rich text and save HTML to the .h field |
| cellNF      | false   | Save number format string to the .z field |
| sheetStubs  | false   | Create cell objects for stub cells |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows ** |
| bookDeps    | false   | If true, parse calculation chains |
| bookFiles   | false   | If true, add raw files to book object ** |
| bookProps   | false   | If true, only parse enough to get book metadata ** |
| bookSheets  | false   | If true, only parse enough to get the sheet names |
| bookVBA     | false   | If true, expose vbaProject.bin to `vbaraw` field ** |

- Even if `cellNF` is false, formatted text (.w) will be generated
- In some cases, sheets may be parsed even if `bookSheets` is false.
- `bookSheets` and `bookProps` combine to give both sets of information
- `Deps` will be an empty object if `bookDeps` is falsy
- `bookFiles` adds a `keys` array (paths in the ZIP) and a `files` hash (whose
  keys are paths and values are objects representing the files)
- `sheetRows-1` rows will be generated when looking at the JSON object output
  (since the header row is counted as a row when parsing the data)
- `bookVBA` merely exposes the raw vba object.  It does not parse the data.

The defaults are enumerated in bits/84_defaults.js

## Writing Options

The exported `write` and `writeFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| bookSST     | false   | Generate Shared String Table ** |
| bookType    | 'xlsx'  | Type of Workbook ("xlsx" or "xlsm" or "xlsb") |

- `bookSST` is slower and more memory intensive, but has better compatibility
  with iOS Numbers
- `bookType = 'xlsb'` is stubbed and far from complete
- The raw data is the only thing guaranteed to be saved.  Formulae, formatting,
  and other niceties may not be serialized (pending CSF standardization)

## Tested Environments

 - NodeJS 0.8, 0.10 (latest release)
 - IE 6/7/8/9/10 using Base64 mode (IE10/11 using HTML5 mode)
 - FF 18 using Base64 or HTML5 mode
 - Chrome 24 using Base64 or HTML5 mode

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xlsx> for XLSX module in nodejs
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS* modules
 - <https://saucelabs.com/u/sheetjs> for XLS* modules using Sauce Labs

## Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.

## Testing

`make test` will run the nodejs-based tests.  To run the in-browser tests, clone
[the oss.sheetjs.com repo](https://github.com/SheetJS/SheetJS.github.io) and
replace the xlsx.js file (then fire up the browser and go to `stress.html`):

```
$ cp xlsx.js ../SheetJS.github.io
$ cd ../SheetJS.github.io
$ simplehttpserver # or "python -mSimpleHTTPServer" or "serve"
$ open -a Chromium.app http://localhost:8000/stress.html
```

For a much smaller test, run `make test_misc`.

## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

## XLS Support

XLS is available in [js-xls](http://git.io/xls).

## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 license are reserved by the Original Author.

It is the opinion of the Original Author that this code conforms to the terms of
the Microsoft Open Specifications Promise, falling under the same terms as
OpenOffice (which is governed by the Apache License v2).  Given the vagaries of
the promise, the Original Author makes no legal claim that in fact end users are
protected from future actions.  It is highly recommended that, for commercial
uses, you consult a lawyer before proceeding.

## References

ISO/IEC 29500:2012(E) "Information technology — Document description and processing languages — Office Open XML File Formats"

OSP-covered specifications:

 - [MS-XLSB]: Excel (.xlsb) Binary File Format
 - [MS-XLSX]: Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File Format
 - [MS-OE376]: Office Implementation Information for ECMA-376 Standards Support
 - [MS-XLDM]: Spreadsheet Data Model File Format

## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-xlsx.png?branch=master)](https://travis-ci.org/SheetJS/js-xlsx)

[![Coverage Status](https://coveralls.io/repos/SheetJS/js-xlsx/badge.png?branch=master)](https://coveralls.io/r/SheetJS/js-xlsx?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/ed5bb2c4c4346a474fef270f847f3f78 "githalytics.com")](http://githalytics.com/SheetJS/js-xlsx)

