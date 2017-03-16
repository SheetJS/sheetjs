# xlsx

Parser and writer for various spreadsheet formats.  Pure-JS cleanroom
implementation from official specifications, related documents, and test files.
Emphasis on parsing and writing robustness, cross-format feature compatibility
with a unified JS representation, and ES3/ES5 browser compatibility back to IE6.

[**In-Browser Demo**](http://oss.sheetjs.com/js-xlsx)

[**Source Code**](http://git.io/xlsx)

[**Commercial Support**](http://sheetjs.com/reinforcements)

[**File format support for known spreadsheet data formats:**](#file-formats)

![circo graph of format support](formats.png)



## Table of Contents

<!-- toc -->

- [Installation](#installation)
  * [JS Ecosystem Demos](#js-ecosystem-demos)
  * [Optional Modules](#optional-modules)
  * [ECMAScript 5 Compatibility](#ecmascript-5-compatibility)
- [Parsing Workbooks](#parsing-workbooks)
- [Working with the Workbook](#working-with-the-workbook)
- [Writing Workbooks](#writing-workbooks)
- [Interface](#interface)
  * [Parsing functions](#parsing-functions)
  * [Writing functions](#writing-functions)
  * [Utilities](#utilities)
- [Workbook / Worksheet / Cell Object Description](#workbook--worksheet--cell-object-description)
  * [General Structures](#general-structures)
  * [Cell Object](#cell-object)
  * [Data Types](#data-types)
  * [Formulae](#formulae)
  * [Worksheet Object](#worksheet-object)
  * [Workbook Object](#workbook-object)
- [Parsing Options](#parsing-options)
  * [Input Type](#input-type)
  * [Guessing File Type](#guessing-file-type)
- [Writing Options](#writing-options)
  * [Supported Output Formats](#supported-output-formats)
  * [Output Type](#output-type)
- [Utility Functions](#utility-functions)
  * [Formulae Output](#formulae-output)
  * [CSV and general DSV Output](#csv-and-general-dsv-output)
  * [JSON](#json)
- [File Formats](#file-formats)
  * [Excel 2007+ XML (XLSX/XLSM)](#excel-2007-xml-xlsxxlsm)
  * [Excel 2.0-95 (BIFF2/BIFF3/BIFF4/BIFF5)](#excel-20-95-biff2biff3biff4biff5)
  * [Excel 97-2004 Binary (BIFF8)](#excel-97-2004-binary-biff8)
  * [Excel 2003-2004 (SpreadsheetML)](#excel-2003-2004-spreadsheetml)
  * [Excel 2007+ Binary (XLSB, BIFF12)](#excel-2007-binary-xlsb-biff12)
  * [OpenDocument Spreadsheet (ODS/FODS) and Uniform Office Spreadsheet (UOS1/2)](#opendocument-spreadsheet-odsfods-and-uniform-office-spreadsheet-uos12)
  * [Comma-Separated Values](#comma-separated-values)
  * [HTML](#html)
- [Testing](#testing)
  * [Tested Environments](#tested-environments)
  * [Test Files](#test-files)
- [Contributing](#contributing)
- [License](#license)
- [References](#references)
- [Badges](#badges)

<!-- tocstop -->

## Installation

With [npm](https://www.npmjs.org/package/xlsx):

```bash
$ npm install xlsx
```

In the browser:

```html
<script lang="javascript" src="dist/xlsx.core.min.js"></script>
```

With [bower](http://bower.io/search/?q=js-xlsx):

```bash
$ bower install js-xlsx
```

CDNjs automatically pulls the latest version and makes all versions available at
<http://cdnjs.com/libraries/xlsx>

### JS Ecosystem Demos

The `demos` directory includes sample projects for:

- [`browserify`](http://browserify.org/)
- [`requirejs`](http://requirejs.org/)
- [`webpack`](https://webpack.js.org/)

### Optional Modules

The node version automatically requires modules for additional features.  Some
of these modules are rather large in size and are only needed in special
circumstances, so they do not ship with the core.  For browser use, they must
be included directly:

```html
<!-- international support from js-codepage -->
<script src="dist/cpexcel.js"></script>
```

An appropriate version for each dependency is included in the dist/ directory.

The complete single-file version is generated at `dist/xlsx.full.min.js`

Webpack and browserify builds include optional modules by default.  Webpack can
be configured to remove support with `resolve.alias`:

```js
  /* uncomment the lines below to remove support */
  resolve: {
    alias: { "./dist/cpexcel.js": "" } // <-- omit international support
  }
```


### ECMAScript 5 Compatibility

Since xlsx.js uses ES5 functions like `Array#forEach`, older browsers require
[Polyfills](http://git.io/QVh77g).  This repo and the gh-pages branch include
[a shim](https://github.com/SheetJS/js-xlsx/blob/master/shim.js)

To use the shim, add the shim before the script tag that loads xlsx.js:

```html
<script type="text/javascript" src="/path/to/shim.js"></script>
```

## Parsing Workbooks

For parsing, the first step is to read the file.  This involves acquiring the
data and feeding it into the library.  Here are a few common scenarios:

- node readFile:

```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

- ajax (for a more complete example that works in older browsers, check the demo
  at <http://oss.sheetjs.com/js-xlsx/ajax.html>):

```js
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

- HTML5 drag-and-drop using readAsBinaryString or readAsArrayBuffer:
  note: readAsBinaryString and readAsArrayBuffer may not be available in every
  browser.  Use dynamic feature tests to determine which method to use.

```js
/* processing array buffers, only required for readAsArrayBuffer */
function fixdata(data) {
  var o = "", l = 0, w = 10240;
  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
  o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
  return o;
}

var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
/* set up drag-and-drop event */
function handleDrop(e) {
  e.stopPropagation();
  e.preventDefault();
  var files = e.dataTransfer.files;
  var i,f;
  for (i = 0; i != files.length; ++i) {
    f = files[i];
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      var workbook;
      if(rABS) {
        /* if binary string, read with type 'binary' */
        workbook = XLSX.read(data, {type: 'binary'});
      } else {
        /* if array buffer, convert to base64 */
        var arr = fixdata(data);
        workbook = XLSX.read(btoa(arr), {type: 'base64'});
      }

      /* DO SOMETHING WITH workbook HERE */
    };
    if(rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  }
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

- HTML5 input file element using readAsBinaryString or readAsArrayBuffer:

```js
/* fixdata and rABS are defined in the drag and drop example */
function handleFile(e) {
  var files = e.target.files;
  var i,f;
  for (i = 0; i != files.length; ++i) {
    f = files[i];
    var reader = new FileReader();
    var name = f.name;
    reader.onload = function(e) {
      var data = e.target.result;

      var workbook;
      if(rABS) {
        /* if binary string, read with type 'binary' */
        workbook = XLSX.read(data, {type: 'binary'});
      } else {
        /* if array buffer, convert to base64 */
        var arr = fixdata(data);
        workbook = XLSX.read(btoa(arr), {type: 'base64'});
      }

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

```js
var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'A1';

/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

/* Find desired cell */
var desired_cell = worksheet[address_of_cell];

/* Get the value */
var desired_value = (desired_cell ? desired_cell.v : undefined);
```

**Complete examples:**

- <http://oss.sheetjs.com/js-xlsx/> HTML5 File API / Base64 Text / Web Workers

Note that older versions of IE do not support HTML5 File API, so the base64 mode
is used for testing.  On OSX you can get the base64 encoding with:

```bash
$ <target_file base64 | pbcopy
```

On Windows XP and up you can get the base64 encoding using `certutil`:

```cmd
> certutil -encode target_file target_file.b64
```

(note: You have to open the file and remove the header and footer lines)

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

```js
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsx');
/* at this point, out.xlsx is a file that you can distribute */
```

- browser generate binary blob and "download" to client
  (using [FileSaver.js](https://github.com/eligrey/FileSaver.js/) for download):

```js
/* bookType can be 'xlsx' or 'xlsm' or 'xlsb' or 'ods' */
var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

var wbout = XLSX.write(workbook,wopts);

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

/* the saveAs call downloads a file on the local machine */
saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx");
```

**Complete examples:**

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

Parse options are described in the [Parsing Options](#parsing-options) section.

### Writing functions

`XLSX.write(wb, write_opts)` attempts to write the workbook `wb`

`XLSX.writeFile(wb, filename, write_opts)` attempts to write `wb` to `filename`

Write options are described in the [Writing Options](#writing-options) section.

### Utilities

Utilities are available in the `XLSX.utils` object:

**Exporting:**

- `sheet_to_json` converts a worksheet object to an array of JSON objects.
  `sheet_to_row_object_array` is an alias that will be removed in the future.
- `sheet_to_csv` generates delimiter-separated-values output.
- `sheet_to_formulae` generates a list of the formulae (with value fallbacks).

Exporters are described in the [Utility Functions](#utility-functions) section.


**Cell and cell address manipulation:**

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

```js
for(var R = range.s.r; R <= range.e.r; ++R) {
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
  }
}
```

### Cell Object

| Key | Description                                                            |
| --- | ---------------------------------------------------------------------- |
| `v` | raw value (see Data Types section for more info)                       |
| `w` | formatted text (if applicable)                                         |
| `t` | cell type: `b` Boolean, `n` Number, `e` error, `s` String, `d` Date    |
| `f` | cell formula encoded as an A1-style string (if applicable)             |
| `F` | range of enclosing array if formula is array formula (if applicable)   |
| `r` | rich text encoding (if applicable)                                     |
| `h` | HTML rendering of the rich text (if applicable)                        |
| `c` | comments associated with the cell                                      |
| `z` | number format string associated with the cell (if requested)           |
| `l` | cell hyperlink object (.Target holds link, .tooltip is tooltip)        |
| `s` | the style/theme of the cell (if applicable)                            |

Built-in export utilities (such as the CSV exporter) will use the `w` text if it
is available.  To change a value, be sure to delete `cell.w` (or set it to
`undefined`) before attempting to export.  The utilities will regenerate the `w`
text from the number format (`cell.z`) and the raw value if possible.

The actual array formula is stored in the `f` field of the first cell in the
array range.  Other cells in the range will omit the `f` field.

### Data Types

The raw value is stored in the `v` field, interpreted based on the `t` field.

Type `b` is the Boolean type.  `v` is interpreted according to JS truth tables.

Type `e` is the Error type. `v` holds the number and `w` holds the common name:

|  Value | Error Meaning   |
| -----: | :-------------- |
| `0x00` | `#NULL!`        |
| `0x07` | `#DIV/0!`       |
| `0x0F` | `#VALUE!`       |
| `0x17` | `#REF!`         |
| `0x1D` | `#NAME?`        |
| `0x24` | `#NUM!`         |
| `0x2A` | `#N/A`          |
| `0x2B` | `#GETTING_DATA` |

Type `n` is the Number type. This includes all forms of data that Excel stores
as numbers, such as dates/times and Boolean fields.  Excel exclusively uses data
that can be fit in an IEEE754 floating point number, just like JS Number, so the
`v` field holds the raw number.  The `w` field holds formatted text.

Type `d` is the Date type, generated only when the option `cellDates` is passed.
Since JSON does not have a natural Date type, parsers are generally expected to
store ISO 8601 Date strings like you would get from `date.toISOString()`.  On
the other hand, writers and exporters should be able to handle date strings and
JS Date objects.  Note that Excel disregards timezone modifiers and treats all
dates in the local timezone.  js-xlsx does not correct for this error.

Type `s` is the String type.  `v` should be explicitly stored as a string to
avoid possible confusion.

Type `z` represents blank stub cells.  These do not have any data or type, and
are not processed by any of the core library functions.  By default these cells
will not be generated; the parser `sheetStubs` option must be set to `true`.

### Formulae

The A1-style formula string is stored in the `f` field.  Even though different
file formats store the formulae in different ways, the formats are converted.

Shared formulae are decompressed and each cell has the correct formula.

Array formulae are stored in the top-left cell of the array block.  All cells
of an array formula have a `F` field corresponding to the range.  A single-cell
formula can be distinguished from a plain formula by the presence of `F` field.

The `sheet_to_formulae` method generates one line per formula or array formula.
Array formulae are rendered in the form `range=formula` while plain cells are
rendered in the form `cell=formula or value`.

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
  ship with this library treat sheets as empty (for example, the CSV output is
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

| Option Name | Default | Description                                          |
| :---------- | ------: | :--------------------------------------------------- |
| type        |         | Input data encoding (see Input Type below)           |
| cellFormula | true    | Save formulae to the .f field **                     |
| cellHTML    | true    | Parse rich text and save HTML to the .h field        |
| cellNF      | false   | Save number format string to the .z field            |
| cellStyles  | false   | Save style/theme info to the .s field                |
| cellDates   | false   | Store dates as type `d` (default is `n`) **          |
| sheetStubs  | false   | Create cell objects of type `z` for stub cells       |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows **            |
| bookDeps    | false   | If true, parse calculation chains                    |
| bookFiles   | false   | If true, add raw files to book object **             |
| bookProps   | false   | If true, only parse enough to get book metadata **   |
| bookSheets  | false   | If true, only parse enough to get the sheet names    |
| bookVBA     | false   | If true, expose vbaProject.bin to `vbaraw` field **  |
| password    | ""      | If defined and file is encrypted, use password **    |
| WTF         | false   | If true, throw errors on unexpected file features ** |

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
- WTF is mainly for development.  By default, the parser will suppress read
  errors on single worksheets, allowing you to read from the worksheets that do
  parse properly. Setting `WTF:1` forces those errors to be thrown.

The defaults are enumerated in bits/84\_defaults.js

### Input Type

Strings can be interpreted in multiple ways.  The `type` parameter for `read`
tells the library how to parse the data argument:

| `type`     | expected input                                                  |
|------------|-----------------------------------------------------------------|
| `"base64"` | string: base64 encoding of the file                             |
| `"binary"` | string:  binary string (`n`-th byte is `data.charCodeAt(n)`)    |
| `"buffer"` | nodejs Buffer                                                   |
| `"array"`  | array: array of 8-bit unsigned int (`n`-th byte is `data[n]`)   |
| `"file"`   | string: filename that will be read and processed (nodejs only)  |

### Guessing File Type

Excel and other spreadsheet tools read the first few bytes and apply other
heuristics to determine a file type.  This enables file type punning: renaming
files with the `.xls` extension will tell your computer to use Excel to open the
file but Excel will know how to handle it.  This library applies similar logic:

| Byte 0 | Raw File Type | Spreadsheet Types                                   |
|:-------|:--------------|:----------------------------------------------------|
| `0xD0` | CFB Container | BIFF 5/8 or password-protected XLSX/XLSB            |
| `0x09` | BIFF Stream   | BIFF 2/3/4/5                                        |
| `0x3C` | XML/HTML      | SpreadsheetML or Flat ODS or UOS1 or HTML           |
| `0x50` | ZIP Archive   | XLSB or XLSX/M or ODS or UOS2                       |
| `0xFE` | UTF8 Text     | SpreadsheetML or Flat ODS or UOS1                   |

## Writing Options

The exported `write` and `writeFile` functions accept an options argument:

| Option Name |  Default | Description                                         |
| :---------- | -------: | :-------------------------------------------------- |
| type        |          | Output data encoding (see Output Type below)        |
| cellDates   |  `false` | Store dates as type `d` (default is `n`)            |
| bookSST     |  `false` | Generate Shared String Table **                     |
| bookType    | `"xlsx"` | Type of Workbook (see below for supported formats)  |
| sheet       |     `""` | Name of Worksheet for single-sheet formats **       |
| compression |  `false` | Use ZIP compression for ZIP-based formats **        |

- `bookSST` is slower and more memory intensive, but has better compatibility
  with older versions of iOS Numbers
- The raw data is the only thing guaranteed to be saved.  Formulae, formatting,
  and other niceties may not be serialized (pending CSF standardization)
- `cellDates` only applies to XLSX output and is not guaranteed to work with
  third-party readers.  Excel itself does not usually write cells with type `d`
  so non-Excel tools may ignore the data or blow up in the presence of dates.

### Supported Output Formats

For broad compatibility with third-party tools, this library supports many
output formats.  The specific file type is controlled with `bookType` option:

| bookType | file ext | container | sheets | Description                       |
| :------- | -------: | :-------: | :----- |:--------------------------------- |
| `xlsx`   | `.xlsx`  |    ZIP    | multi  | Excel 2007+ XML Format            |
| `xlsm`   | `.xlsm`  |    ZIP    | multi  | Excel 2007+ Macro XML Format      |
| `xlsb`   | `.xlsb`  |    ZIP    | multi  | Excel 2007+ Binary Format         |
| `ods`    | `.ods`   |    ZIP    | multi  | OpenDocument Spreadsheet          |
| `biff2`  | `.xls`   |   none    | single | Excel 2.0 Worksheet format        |
| `fods`   | `.fods`  |   none    | multi  | Flat OpenDocument Spreadsheet     |
| `csv`    | `.csv`   |   none    | single | Comma Separated Values            |

- `compression` only applies to formats with ZIP containers.
- Formats that only support a single sheet require a `sheet` option specifying
  the worksheet.  If the string is empty, the first worksheet is used.

### Output Type

The `type` argument for `write` mirrors the `type` argument for `read`:

| `type`     | output                                                          |
|------------|-----------------------------------------------------------------|
| `"base64"` | string: base64 encoding of the file                             |
| `"binary"` | string:  binary string (`n`-th byte is `data.charCodeAt(n)`)    |
| `"buffer"` | nodejs Buffer                                                   |
| `"file"`   | string: name of file to be written (nodejs only)                |


## Utility Functions

The `sheet_to_*` functions accept a worksheet and an optional options object.

The examples are based on the following worksheet:

```
XXX| A | B | C | D | E | F | G |
---+---+---+---+---+---+---+---+
 1 | S | h | e | e | t | J | S |
 2 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |
 3 | 2 | 3 | 4 | 5 | 6 | 7 | 8 |
```

### Formulae Output

`XLSX.utils.sheet_to_formulae` generates an array of commands that represent
how a person would enter data into an application.  Each entry is of the form
`A1-cell-address=formula-or-value`.  String literals are prefixed with a `'` in
accordance with Excel.  For the example sheet:

```js
> var o = XLSX.utils.sheet_to_formulae(ws);
> o.filter(function(v, i) { return i % 5 === 0; });
[ 'A1=\'S', 'F1=\'J', 'D2=4', 'B3=3', 'G3=8' ]
```

### CSV and general DSV Output

As an alternative to the `writeFile` CSV type, `XLSX.utils.sheet_to_csv` also
produces CSV output.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
| FS          |  `","`   | "Field Separator"  delimiter between fields         |
| RS          |  `"\n"`  | "Record Separator" delimiter between rows           |

For the example sheet:

```js
> console.log(XLSX.utils.sheet_to_csv(ws));
S,h,e,e,t,J,S
1,2,3,4,5,6,7
2,3,4,5,6,7,8
> console.log(XLSX.utils.sheet_to_csv(ws, {FS:"\t"}));
S	h	e	e	t	J	S
1	2	3	4	5	6	7
2	3	4	5	6	7	8
> console.log(X.utils.sheet_to_csv(_ws,{FS:":",RS:"|"}));
S:h:e:e:t:J:S|1:2:3:4:5:6:7|2:3:4:5:6:7:8|
```

### JSON

`XLSX.utils.sheet_to_json` and the alias `XLSX.utils.sheet_to_row_object_array`
generate different types of JS objects.  The function takes an options argument:

| Option Name |  Default | Description                                         |
| :---------- | :------: | :-------------------------------------------------- |
| raw         | `false`  | Use raw values (true) or formatted strings (false)  |
| range       | from WS  | Override Range (see table below)                    |
| header      |          | Control output format (see table below)             |

- `raw` only affects cells which have a format code (`.z`) field or a formatted
  text (`.w`) field.
- If `header` is specified, the first row is considered a data row; if `header`
  is not specified, the first row is the header row and not considered data.
- When `header` is not specified, the conversion will automatically disambiguate
  header entries by affixing `_` and a count starting at `1`.  For example, if
  three columns have header `foo` the output fields are `foo`, `foo_1`, `foo_2`

`range` is expected to be one of:

| `range`          | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| (number)         | Use worksheet range but set starting row to the value     |
| (string)         | Use specified range (A1-style bounded range string)       |
| (default)        | Use worksheet range (`ws['!ref']`)                        |

`header` is expected to be one of:

| `header`         | Description                                               |
| :--------------- | :-------------------------------------------------------- |
| `1`              | Generate an array of arrays                               |
| `"A"`            | Row object keys are literal column labels                 |
| array of strings | Use specified strings as keys in row objects              |
| (default)        | Read and disambiguate first row as keys                   |

For the example sheet:

```js
> console.log(X.utils.sheet_to_json(_ws));
[ { S: 1, h: 2, e: 3, e_1: 4, t: 5, J: 6, S_1: 7 },
  { S: 2, h: 3, e: 4, e_1: 5, t: 6, J: 7, S_1: 8 } ]

> console.log(X.utils.sheet_to_json(_ws, {header:1}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],
  [ 2, 3, 4, 5, 6, 7, 8 ] ]

> console.log(X.utils.sheet_to_json(_ws, {header:"A"}));
[ { A: 'S', B: 'h', C: 'e', D: 'e', E: 't', F: 'J', G: 'S' },
  { A: 1, B: 2, C: 3, D: 4, E: 5, F: 6, G: 7 },
  { A: 2, B: 3, C: 4, D: 5, E: 6, F: 7, G: 8 } ]
> console.log(X.utils.sheet_to_json(_ws, {header:["A","E","I","O","U","6","9"]}));
[ { '6': 'J', '9': 'S', A: 'S', E: 'h', I: 'e', O: 'e', U: 't' },
  { '6': 6, '9': 7, A: 1, E: 2, I: 3, O: 4, U: 5 },
  { '6': 7, '9': 8, A: 2, E: 3, I: 4, O: 5, U: 6 } ]
```

Example showing the effect of `raw`:

```js
> _ws['A2'].w = "1";                         // set A2 formatted string value
> console.log(X.utils.sheet_to_json(_ws, {header:1}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ '1', 2, 3, 4, 5, 6, 7 ],                 // <-- A2 uses the formatted string
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
> console.log(X.utils.sheet_to_json(_ws, {header:1, raw:true}));
[ [ 'S', 'h', 'e', 'e', 't', 'J', 'S' ],
  [ 1, 2, 3, 4, 5, 6, 7 ],                   // <-- A2 uses the raw value
  [ 2, 3, 4, 5, 6, 7, 8 ] ]
```

## File Formats

Despite the library name `xlsx`, it supports numerous spreadsheet file formats:

| Format                                                       | Read  | Write |
|:-------------------------------------------------------------|:-----:|:-----:|
| **Excel Worksheet/Workbook Formats**                         |:-----:|:-----:|
| Excel 2007+ XML Formats (XLSX/XLSM)                          |  :o:  |  :o:  |
| Excel 2007+ Binary Format (XLSB BIFF12)                      |  :o:  |  :o:  |
| Excel 2003-2004 XML Format (XML "SpreadsheetML")             |  :o:  |  :o:  |
| Excel 97-2004 (XLS BIFF8)                                    |  :o:  |       |
| Excel 5.0/95 (XLS BIFF5)                                     |  :o:  |       |
| Excel 4.0 (XLS/XLW BIFF4)                                    |  :o:  |       |
| Excel 3.0 (XLS BIFF3)                                        |  :o:  |       |
| Excel 2.0/2.1 (XLS BIFF2)                                    |  :o:  |  :o:  |
| **Excel Supported Text Formats**                             |:-----:|:-----:|
| Delimiter-Separated Values (CSV/TSV/DSV)                     |       |  :o:  |
| **Other Workbook/Worksheet Formats**                         |:-----:|:-----:|
| OpenDocument Spreadsheet (ODS)                               |  :o:  |  :o:  |
| Flat XML ODF Spreadsheet (FODS)                              |  :o:  |  :o:  |
| Uniform Office Format Spreadsheet (标文通 UOS1/UOS2)         |  :o:  |       |
| **Other Common Spreadsheet Output Formats**                  |:-----:|:-----:|
| HTML Tables                                                  |  :o:  |       |

### Excel 2007+ XML (XLSX/XLSM)

XLSX and XLSM files are ZIP containers containing a series of XML files in
accordance with the Open Packaging Conventions (OPC).  The XLSM filetype, almost
identical to XLSX, is used for files containing macros.

The format is standardized in ECMA-376 and later in ISO/IEC 29500.  Excel does
not follow the specification, and there are additional documents discussing how
Excel deviates from the specification.

### Excel 2.0-95 (BIFF2/BIFF3/BIFF4/BIFF5)

BIFF 2/3 XLS are single-sheet streams of binary records.  Excel 4 introduced
the concept of a workbook (`XLW` files) but also had single-sheet `XLS` format.
The structure is largely similar to the Lotus 1-2-3 file formats.  BIFF5/8/12
extended the format in various ways but largely stuck to the same record format.

There is no official specification for any of these formats.  Excel 95 can write
files in these formats, so record lengths and fields were backsolved by writing
in all of the supported formats and comparing files.  Excel 2016 can generate
BIFF5 files, enabling a full suite of file tests starting from XLSX or BIFF2.

### Excel 97-2004 Binary (BIFF8)

BIFF8 exclusively uses the Compound File Binary container format, splitting some
content into streams within the file.  At its core, it still uses an extended
version of the binary record format from older versions of BIFF.

The `MS-XLS` specification covers the basics of the file format, and other
specifications expand on serialization of features like properties.

### Excel 2003-2004 (SpreadsheetML)

Predating XLSX, SpreadsheetML files are simple XML files.  There is no official
and comprehensive specification, although MS has released whitepapers on the
format.  Since Excel 2016 can generate SpreadsheetML files, backsolving is
pretty straightforward.

### Excel 2007+ Binary (XLSB, BIFF12)

Introduced in parallel with XLSX, the XLSB filetype combines BIFF architecture
with the content separation and ZIP container of XLSX.  For the most part nodes
in an XLSX sub-file can be mapped to XLSB records in a corresponding sub-file.

The `MS-XLSB` specification covers the basics of the file format, and other
specifications expand on serialization of features like properties.

### OpenDocument Spreadsheet (ODS/FODS) and Uniform Office Spreadsheet (UOS1/2)

ODS is an XML-in-ZIP format akin to XLSX while FODS is an XML format akin to
SpreadsheetML.  Both are detailed in the OASIS standard, but tools like LO/OO
add undocumented extensions.

UOS is a very similar format, and it comes in 2 varieties corresponding to ODS
and FODS respectively.  For the most part, the difference between the formats
lies in the names of tags and attributes.

### Comma-Separated Values

Excel CSV deviates from RFC4180 in a number of important ways.  The generated
CSV files should generally work in Excel although they may not work in RFC4180
compatible readers.

### HTML

Excel HTML worksheets include special metadata encoded in styles.  For example,
`mso-number-format` is a localized string containing the number format.  Despite
the metadata the output is valid HTML, although it does accept bare `&` symbols.

## Testing

`make test` will run the node-based tests.  By default it runs tests on files in
every supported format.  To test a specific file type, set `FMTS` to the format
you want to test.  Feature-specific tests are avaialble with `make test_misc`

```bash
$ make test        # run full tests
$ make test_xls    # only use the XLS test files
$ make test_xlsx   # only use the XLSX test files
$ make test_xlsb   # only use the XLSB test files
$ make test_xml    # only use the XLSB test files
$ make test_ods    # only use the XLSB test files
```

To enable all errors, set the environment variable `WTF=1`:

```bash
$ make test        # run full tests
$ WTF=1 make test  # enable all error messages
```

Flow and JSHint/JSCS checks are available:

```bash
$ make lint        # JSHint and JSCS checks
$ make flow        # make lint + Flow checking
```

To run the in-browser tests, clone the repo for
[oss.sheetjs.com](https://github.com/SheetJS/SheetJS.github.io) and replace
the xlsx.js file (then fire up the browser and go to `stress.html`):

```bash
$ cp xlsx.js ../SheetJS.github.io
$ cd ../SheetJS.github.io
$ simplehttpserver # or "python -mSimpleHTTPServer" or "serve"
$ open -a Chromium.app http://localhost:8000/stress.html
```
### Tested Environments

 - NodeJS 0.8, 0.9, 0.10, 0.11, 0.12, 4.x, 5.x, 6.x, 7.x
 - IE 6/7/8/9/10/11 (IE6-9 browsers require shims for interacting with client)
 - Chrome 24+
 - Safari 6+
 - FF 18+

Tests utilize the mocha testing framework.  Travis-CI and Sauce Labs links:

 - <https://travis-ci.org/SheetJS/js-xlsx> for XLSX module in nodejs
 - <https://travis-ci.org/SheetJS/SheetJS.github.io> for XLS\* modules
 - <https://saucelabs.com/u/sheetjs> for XLS\* modules using Sauce Labs

### Test Files

Test files are housed in [another repo](https://github.com/SheetJS/test_files).

Running `make init` will refresh the `test_files` submodule and get the files.



## Contributing

Due to the precarious nature of the Open Specifications Promise, it is very
important to ensure code is cleanroom.  Consult CONTRIBUTING.md

The xlsx.js file is constructed from the files in the `bits` subdirectory. The
build script (run `make`) will concatenate the individual bits to produce the
script.  Before submitting a contribution, ensure that running make will produce
the xlsx.js file exactly.  The simplest way to test is to add the script:

```bash
$ git add xlsx.js
$ make clean
$ make
$ git diff xlsx.js
```

To produce the dist files, run `make dist`.  The dist files are updated in each
version release and *should not be committed between versions*.


## License

Please consult the attached LICENSE file for details.  All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.


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
 - [MS-ODRAW]: Office Drawing Binary File Format
 - [MS-ODRAWXML]: Office Drawing Extensions to Office Open XML Structure
 - [MS-OVBA]: Office VBA File Format Structure
 - [MS-CTXLS]: Excel Custom Toolbar Binary File Format
 - [MS-XLDM]: Spreadsheet Data Model File Format
 - [MS-EXSPXML3]: Excel Calculation Version 2 Web Service XML Schema
 - [XLS]: Microsoft Office Excel 97-2007 Binary File Format Specification

Open Document Format for Office Applications Version 1.2 (29 September 2011)

Worksheet File Format (From Lotus) December 1984


## Badges

[![Build Status](https://saucelabs.com/browser-matrix/xlsx.svg)](https://saucelabs.com/u/xlsx)

[![Build Status](https://travis-ci.org/SheetJS/js-xlsx.svg?branch=master)](https://travis-ci.org/SheetJS/js-xlsx)

[![Coverage Status](http://img.shields.io/coveralls/SheetJS/js-xlsx/master.svg)](https://coveralls.io/r/SheetJS/js-xlsx?branch=master)

[![NPM Downloads](https://img.shields.io/npm/dt/xlsx.svg)](https://npmjs.org/package/xlsx)

[![Dependencies Status](https://david-dm.org/sheetjs/js-xlsx/status.svg)](https://david-dm.org/sheetjs/js-xlsx)

[![ghit.me](https://ghit.me/badge.svg?repo=sheetjs/js-xlsx)](https://ghit.me/repo/sheetjs/js-xlsx)

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
