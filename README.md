# xlsx

Parser and writer for various spreadsheet formats.  Pure-JS cleanroom
implementation from official specifications and related documents.

Supported read formats:

- Excel 2007+ XML Formats (XLSX/XLSM)
- Excel 2007+ Binary Format (XLSB)
- Excel 2003-2004 XML Format (XML "SpreadsheetML")
- Excel 97-2004 (XLS BIFF8)
- Excel 5.0/95 (XLS BIFF5)
- OpenDocument Spreadsheet (ODS)

Supported write formats:

- XLSX
- CSV (and general DSV)
- JSON and JS objects (various styles)

Demo: <http://oss.sheetjs.com/js-xlsx>

Source: <http://git.io/xlsx>

## Installation

With [npm](https://www.npmjs.org/package/xlsx):

    npm install xlsx

In the browser:

    <script lang="javascript" src="dist/xlsx.core.min.js"></script>

With [bower](http://bower.io/search/?q=js-xlsx):

    bower install js-xlsx

CDNjs automatically pulls the latest version and makes all versions available at
<http://cdnjs.com/libraries/xlsx>

## Optional Modules

The node version automatically requires modules for additional features.  Some
of these modules are rather large in size and are only needed in special
circumstances, so they do not ship with the core.  For browser use, they must
be included directly:

    <!-- international support from https://github.com/sheetjs/js-codepage -->
    <script src="dist/cpexcel.js"></script>
    <!-- ODS support -->
    <script src="dist/ods.js"></script>

An appropriate version for each dependency is included in the dist/ directory.

The complete single-file version is generated at `dist/xlsx.full.min.js`

## ECMAScript 5 Compatibility

Since xlsx.js uses ES5 functions like `Array#forEach`, older browsers require
[Polyfills](http://git.io/QVh77g).  This repo and the gh-pages branch include
[a shim](https://github.com/SheetJS/js-xlsx/blob/master/shim.js)

To use the shim, add the shim before the script tag that loads xlsx.js:

    <script type="text/javascript" src="/path/to/shim.js"></script>

## Parsing Workbooks

For parsing, the first step is to read the file.  This involves acquiring the
data and feeding it into the library.  Here are a few common scenarios:

- node readFile:

```javascript
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

- ajax (for a more complete example that works in older browsers, check the demo
  at <http://oss.sheetjs.com/js-xlsx/ajax.html>):

```javascript
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

- HTML5 drag-and-drop using readAsBinaryString:

```javascript
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
      var workbook = XLSX.read(data, {type: 'binary'});

      /* DO SOMETHING WITH workbook HERE */
    };
    reader.readAsBinaryString(f);
  }
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

- HTML5 input file element using readAsBinaryString:

```javascript
function handleFile(e) {
  var files = e.target.files;
  var i,f;
  for (i = 0, f = files[i]; i != files.length; ++i) {
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      var workbook = XLSX.read(data, {type: 'binary'});

      /* DO SOMETHING WITH workbook HERE */
    };
    reader.readAsBinaryString(f);
  }
}
input_dom_element.addEventListener('change', handleFile, false);
```

## Working with the Workbook

The full object format is described later in this README.

This example extracts the value stored in cell A1 from the first worksheet:

```javascript
var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'A1';

/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

/* Find desired cell */
var desired_cell = worksheet[address_of_cell];

/* Get the value */
var desired_value = desired_cell.v;
```

This example iterates through every nonempty of every sheet and dumps values:

```javascript
var sheet_name_list = workbook.SheetNames;
sheet_name_list.forEach(function(y) { /* iterate through sheets */
  var worksheet = workbook.Sheets[y];
  for (z in worksheet) {
    /* all keys that do not begin with "!" correspond to cell addresses */
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

- <https://github.com/SheetJS/js-xlsx/blob/master/bin/xlsx.njs> node

The node version installs a command line tool `xlsx` which can read spreadsheet
files and output the contents in various formats.  The source is available at
`xlsx.njs` in the bin directory.

Some helper functions in `XLSX.utils` generate different views of the sheets:

- `XLSX.utils.sheet_to_csv` generates CSV
- `XLSX.utils.sheet_to_json` generates an array of objects
- `XLSX.utils.sheet_to_formulae` generates a list of formulae

## Writing Workbooks

For writing, the first step is to generate output data.  The helper functions
`write` and `writeFile` will produce the data in various formats suitable for
dissemination.  The second step is to actual share the data with the end point.
Assuming `workbook` is a workbook object:

- nodejs write to file:

```javascript
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsx');
/* at this point, out.xlsx is a file that you can distribute */
```

- write to binary string (using FileSaver.js):

```javascript
/* bookType can be 'xlsx' or 'xlsm' or 'xlsb' */
var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

var wbout = XLSX.write(workbook,wopts);

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

/* the saveAs call downloads a file on the local machine */
saveAs(new Blob([s2ab(wbout)],{type:""}), "test.xlsx")
```

Complete examples:

- <http://sheetjs.com/demos/writexlsx.html> generates a simple file
- <http://git.io/WEK88Q> writing an array of arrays in nodejs
- <http://sheetjs.com/demos/table.html> exporting an HTML table

## Interface

`XLSX` is the exposed variable in the browser and the exported node variable

`XLSX.version` is the version of the library (added by the build script).

`XLSX.SSF` is an embedded version of the [format library](http://git.io/ssf).

### Parsing functions

`XLSX.read(data, read_opts)` attempts to parse `data`.

`XLSX.readFile(filename, read_opts)` attempts to read `filename` and parse.

### Writing functions

`XLSX.write(wb, write_opts)` attempts to write the workbook `wb`

`XLSX.writeFile(wb, filename, write_opts)` attempts to write `wb` to `filename`

### Utilities

Utilities are available in the `XLSX.utils` object:

Exporting:

- `sheet_to_json` converts a workbook object to an array of JSON objects.
- `sheet_to_csv` generates delimiter-separated-values output
- `sheet_to_formulae` generates a list of the formulae (with value fallbacks)

Cell and cell address manipulation:

- `format_cell` generates the text value for a cell (using number formats)
- `{en,de}code_{row,col}` convert between 0-indexed rows/cols and A1 forms.
- `{en,de}code_cell` converts cell addresses
- `{en,de}code_range` converts cell ranges

## Workbook / Worksheet / Cell Object Description

js-xlsx conforms to the Common Spreadsheet Format (CSF):

### General Structures

Cell address objects are stored as `{c:C, r:R}` where `C` and `R` are 0-indexed
column and row numbers, respectively.  For example, the cell address `B5` is
represented by the object `{c:1, r:4}`.

Cell range objects are stored as `{s:S, e:E}` where `S` is the first cell and
`E` is the last cell in the range.  The ranges are inclusive.  For example, the
range `A3:B7` is represented by the object `{s:{c:0, r:2}, e:{c:1, r:6}}`. Utils
use the following pattern to walk each of the cells in a range:

```javascript
for(var R = range.s.r; R <= range.e.r; ++R) {
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
  }
}
```

### Cell Object

| Key | Description |
| --- | ----------- |
| `v` | raw value (see Data Types section for more info) |
| `w` | formatted text (if applicable) |
| `t` | cell type: `b` Boolean, `n` Number, `e` error, `s` String, `d` Date |
| `f` | cell formula (if applicable) |
| `r` | rich text encoding (if applicable) |
| `h` | HTML rendering of the rich text (if applicable) |
| `c` | comments associated with the cell ** |
| `z` | number format string associated with the cell (if requested) |
| `l` | cell hyperlink object (.Target holds link, .tooltip is tooltip) |
| `s` | the style/theme of the cell (if applicable) |

Built-in export utilities (such as the CSV exporter) will use the `w` text if it
is available.  To change a value, be sure to delete `cell.w` (or set it to
`undefined`) before attempting to export.  The utilities will regenerate the `w`
text from the number format (`cell.z`) and the raw value if possible.

### Data Types

The raw value is stored in the `v` field, interpreted based on the `t` field.

Type `b` is the Boolean type.  `v` is interpreted according to JS truth tables

Type `e` is the Error type. `v` holds the number and `w` holds the common name:

| Value | Error Meaning |
| ----: | :------------ |
|  0x00 | #NULL!        |
|  0x07 | #DIV/0!       |
|  0x0F | #VALUE!       |
|  0x17 | #REF!         |
|  0x1D | #NAME?        |
|  0x24 | #NUM!         |
|  0x2A | #N/A          |
|  0x2B | #GETTING_DATA |

Type `n` is the Number type. This includes all forms of data that Excel stores
as numbers, such as dates/times and Boolean fields.  Excel exclusively uses data
that can be fit in an IEEE754 floating point number, just like JS Number, so the
`v` field holds the raw number.  The `w` field holds formatted text.

Type `d` is the Date type, generated only when the option `cellDates` is passed.
Since JSON does not have a natural Date type, parsers are generally expected to
store ISO 8601 Date strings like you would get from `date.toISOString()`.  On
the other hand, writers and exporters should be able to handle date strings and
JS Date objects.  Note that Excel disregards the timezone modifier and treats all
dates in the local timezone.  js-xlsx does not correct for this error.

Type `s` is the String type.  `v` should be explicitly stored as a string to
avoid possible confusion.


### Worksheet Object

Each key that does not start with `!` maps to a cell (using `A-1` notation)

`worksheet[address]` returns the cell object for the specified address.

Special worksheet keys (accessible as `worksheet[key]`, each starting with `!`):

- `ws['!ref']`: A-1 based range representing the worksheet range. Functions that
  work with sheets should use this parameter to determine the range.  Cells that
  are assigned outside of the range are not processed.  In particular, when
  writing a worksheet by hand, be sure to update the range.  For a longer
  discussion, see <http://git.io/KIaNKQ>

  Functions that handle worksheets should test for the presence of `!ref` field.
  If the `!ref` is omitted or is not a valid range, functions are free to treat
  the sheet as empty or attempt to guess the range.  The standard utilities that
  ship with this library treat sheets as empty (for example, the CSV output is an
  empty string).

  When reading a worksheet with the `sheetRows` property set, the ref parameter
  will use the restricted range.  The original range is set at `ws['!fullref']`

- `ws['!cols']`: array of column properties objects.  Column widths are actually
  stored in files in a normalized manner, measured in terms of the "Maximum
  Digit Width" (the largest width of the rendered digits 0-9, in pixels).  When
  parsed, the column objects store the pixel width in the `wpx` field, character
  width in the `wch` field, and the maximum digit width in the `MDW` field.

- `ws['!merges']`: array of range objects corresponding to the merged cells in
  the worksheet.  Plaintext utilities are unaware of merge cells.  CSV export
  will write all cells in the merge range if they exist, so be sure that only
  the first cell (upper-left) in the range is set.

### Workbook Object

`workbook.SheetNames` is an ordered list of the sheets in the workbook

`wb.Sheets[sheetname]` returns an object representing the worksheet.

`wb.Props` is an object storing the standard properties.  `wb.Custprops` stores
custom properties.  Since the XLS standard properties deviate from the XLSX
standard, XLS parsing stores core properties in both places.  .


## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| cellFormula | true    | Save formulae to the .f field ** |
| cellHTML    | true    | Parse rich text and save HTML to the .h field |
| cellNF      | false   | Save number format string to the .z field |
| cellStyles  | false   | Save style/theme info to the .s field |
| cellDates   | false   | Store dates as type `d` (default is `n`) ** |
| sheetStubs  | false   | Create cell objects for stub cells |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows ** |
| bookDeps    | false   | If true, parse calculation chains |
| bookFiles   | false   | If true, add raw files to book object ** |
| bookProps   | false   | If true, only parse enough to get book metadata ** |
| bookSheets  | false   | If true, only parse enough to get the sheet names |
| bookVBA     | false   | If true, expose vbaProject.bin to `vbaraw` field ** |
| password    | ""      | If defined and file is encrypted, use password ** |

- `cellFormula` option only applies to formats that require extra processing to
  parse formulae (XLS/XLSB).
- Even if `cellNF` is false, formatted text will be generated and saved to `.w`
- In some cases, sheets may be parsed even if `bookSheets` is false.
- `bookSheets` and `bookProps` combine to give both sets of information
- `Deps` will be an empty object if `bookDeps` is falsy
- `bookFiles` behavior depends on file type:
    * `keys` array (paths in the ZIP) for ZIP-based formats
    * `files` hash (mapping paths to objects representing the files) for ZIP
    * `cfb` object for formats using CFB containers
- `sheetRows-1` rows will be generated when looking at the JSON object output
  (since the header row is counted as a row when parsing the data)
- `bookVBA` merely exposes the raw vba object.  It does not parse the data.
- `cellDates` currently does not convert numerical dates to JS dates.
- Currently only XOR encryption is supported.  Unsupported error will be thrown
  for files employing other encryption methods.

The defaults are enumerated in bits/84_defaults.js

## Writing Options

The exported `write` and `writeFile` functions accept an options argument:

| Option Name | Default | Description |
| :---------- | ------: | :---------- |
| cellDates   | false   | Store dates as type `d` (default is `n`) |
| bookSST     | false   | Generate Shared String Table ** |
| bookType    | 'xlsx'  | Type of Workbook ("xlsx" or "xlsm" or "xlsb") |

- `bookSST` is slower and more memory intensive, but has better compatibility
  with older versions of iOS Numbers
- `bookType = 'xlsb'` is stubbed and far from complete
- The raw data is the only thing guaranteed to be saved.  Formulae, formatting,
  and other niceties may not be serialized (pending CSF standardization)
- `cellDates` only applies to XLSX output and is not guaranteed to work with
  third-party readers.  Excel itself does not usually write cells with type `d`
  so non-Excel tools may ignore the data or blow up in the presence of dates.

## Tested Environments

 - NodeJS 0.8, 0.10 (latest release), 0.11.14 (unstable), io.js
 - IE 6/7/8/9/10/11 using Base64 mode (IE10/11 using HTML5 mode)
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

`make test` will run the node-based tests.  To run the in-browser tests, clone
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

The xlsx.js file is constructed from the files in the `bits` subdirectory. The
build script (run `make`) will concatenate the individual bits to produce the
script.  Before submitting a contribution, ensure that running make will produce
the xlsx.js file exactly.  The simplest way to test is to move the script:

```
$ mv xlsx.js xlsx.new.js
$ make
$ diff xlsx.js xlsx.new.js
```

To produce the dist files, run `make dist`.  The dist files are updated in each
version release and should not be committed between versions.

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
 - [MS-CFB]: Compound File Binary File Format
 - [MS-XLS]: Excel Binary File Format (.xls) Structure Specification
 - [MS-ODATA]: Open Data Protocol (OData)
 - [MS-OFFCRYPTO]: Office Document Cryptography Structure
 - [MS-OLEDS]: Object Linking and Embedding (OLE) Data Structures
 - [MS-OLEPS]: Object Linking and Embedding (OLE) Property Set Data Structures
 - [MS-OSHARED]: Office Common Data Types and Objects Structures
 - [MS-OVBA]: Office VBA File Format Structure
 - [MS-CTXLS]: Excel Custom Toolbar Binary File Format
 - [MS-XLDM]: Spreadsheet Data Model File Format
 - [MS-EXSPXML3]: Excel Calculation Version 2 Web Service XML Schema
 - [XLS]: Microsoft Office Excel 97-2007 Binary File Format Specification

Open Document Format for Office Applications Version 1.2 (29 September 2011)


## Badges

[![Build Status](https://travis-ci.org/SheetJS/js-xlsx.svg?branch=master)](https://travis-ci.org/SheetJS/js-xlsx)

[![Coverage Status](http://img.shields.io/coveralls/SheetJS/js-xlsx/master.svg)](https://coveralls.io/r/SheetJS/js-xlsx?branch=master)

[![githalytics.com alpha](https://cruel-carlota.pagodabox.com/ed5bb2c4c4346a474fef270f847f3f78 "githalytics.com")](http://githalytics.com/SheetJS/js-xlsx)
