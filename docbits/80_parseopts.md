## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description                                          |
| :---------- | ------: | :--------------------------------------------------- |
|`type`       |         | Input data encoding (see Input Type below)           |
|`raw`        | false   | If true, plain text parsing will not parse values ** |
|`codepage`   |         | If specified, use code page when appropriate **      |
|`cellFormula`| true    | Save formulae to the .f field                        |
|`cellHTML`   | true    | Parse rich text and save HTML to the `.h` field      |
|`cellNF`     | false   | Save number format string to the `.z` field          |
|`cellStyles` | false   | Save style/theme info to the `.s` field              |
|`cellText`   | true    | Generated formatted text to the `.w` field           |
|`cellDates`  | false   | Store dates as type `d` (default is `n`)             |
|`dateNF`     |         | If specified, use the string for date code 14 **     |
|`sheetStubs` | false   | Create cell objects of type `z` for stub cells       |
|`sheetRows`  | 0       | If >0, read the first `sheetRows` rows **            |
|`bookDeps`   | false   | If true, parse calculation chains                    |
|`bookFiles`  | false   | If true, add raw files to book object **             |
|`bookProps`  | false   | If true, only parse enough to get book metadata **   |
|`bookSheets` | false   | If true, only parse enough to get the sheet names    |
|`bookVBA`    | false   | If true, copy VBA blob to `vbaraw` field **          |
|`password`   | ""      | If defined and file is encrypted, use password **    |
|`WTF`        | false   | If true, throw errors on unexpected file features ** |
|`sheets`     |         | If specified, only parse specified sheets **         |
|`PRN`        | false   | If true, allow parsing of PRN files **               |
|`xlfn`       | false   | If true, preserve `_xlfn.` prefixes in formulae **   |

- Even if `cellNF` is false, formatted text will be generated and saved to `.w`
- In some cases, sheets may be parsed even if `bookSheets` is false.
- Excel aggressively tries to interpret values from CSV and other plain text.
  This leads to surprising behavior! The `raw` option suppresses value parsing.
- `bookSheets` and `bookProps` combine to give both sets of information
- `Deps` will be an empty object if `bookDeps` is false
- `bookFiles` behavior depends on file type:
    * `keys` array (paths in the ZIP) for ZIP-based formats
    * `files` hash (mapping paths to objects representing the files) for ZIP
    * `cfb` object for formats using CFB containers
- `sheetRows-1` rows will be generated when looking at the JSON object output
  (since the header row is counted as a row when parsing the data)
- By default all worksheets are parsed.  `sheets` restricts based on input type:
    * number: zero-based index of worksheet to parse (`0` is first worksheet)
    * string: name of worksheet to parse (case insensitive)
    * array of numbers and strings to select multiple worksheets.
- `bookVBA` merely exposes the raw VBA CFB object.  It does not parse the data.
  XLSM and XLSB store the VBA CFB object in `xl/vbaProject.bin`. BIFF8 XLS mixes
  the VBA entries alongside the core Workbook entry, so the library generates a
  new XLSB-compatible blob from the XLS CFB container.
- `codepage` is applied to BIFF2 - BIFF5 files without `CodePage` records and to
  CSV files without BOM in `type:"binary"`.  BIFF8 XLS always defaults to 1200.
- `PRN` affects parsing of text files without a common delimiter character.
- Currently only XOR encryption is supported.  Unsupported error will be thrown
  for files employing other encryption methods.
- Newer Excel functions are serialized with the `_xlfn.` prefix, hidden from the
  user. SheetJS will strip `_xlfn.` normally. The `xlfn` option preserves them.
- WTF is mainly for development.  By default, the parser will suppress read
  errors on single worksheets, allowing you to read from the worksheets that do
  parse properly. Setting `WTF:true` forces those errors to be thrown.

### Input Type

Strings can be interpreted in multiple ways.  The `type` parameter for `read`
tells the library how to parse the data argument:

| `type`     | expected input                                                  |
|------------|-----------------------------------------------------------------|
| `"base64"` | string: Base64 encoding of the file                             |
| `"binary"` | string: binary string (byte `n` is `data.charCodeAt(n)`)        |
| `"string"` | string: JS string (characters interpreted as UTF8)              |
| `"buffer"` | nodejs Buffer                                                   |
| `"array"`  | array: array of 8-bit unsigned int (byte `n` is `data[n]`)      |
| `"file"`   | string: path of file that will be read (nodejs only)            |

### Guessing File Type

<details>
  <summary><b>Implementation Details</b> (click to show)</summary>

Excel and other spreadsheet tools read the first few bytes and apply other
heuristics to determine a file type.  This enables file type punning: renaming
files with the `.xls` extension will tell your computer to use Excel to open the
file but Excel will know how to handle it.  This library applies similar logic:

| Byte 0 | Raw File Type | Spreadsheet Types                                   |
|:-------|:--------------|:----------------------------------------------------|
| `0xD0` | CFB Container | BIFF 5/8 or password-protected XLSX/XLSB or WQ3/QPW |
| `0x09` | BIFF Stream   | BIFF 2/3/4/5                                        |
| `0x3C` | XML/HTML      | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x50` | ZIP Archive   | XLSB or XLSX/M or ODS or UOS2 or plain text         |
| `0x49` | Plain Text    | SYLK or plain text                                  |
| `0x54` | Plain Text    | DIF or plain text                                   |
| `0xEF` | UTF8 Encoded  | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0xFF` | UTF16 Encoded | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x00` | Record Stream | Lotus WK\* or Quattro Pro or plain text             |
| `0x7B` | Plain text    | RTF or plain text                                   |
| `0x0A` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x0D` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |
| `0x20` | Plain text    | SpreadsheetML / Flat ODS / UOS1 / HTML / plain text |

DBF files are detected based on the first byte as well as the third and fourth
bytes (corresponding to month and day of the file date)

Plain text format guessing follows the priority order:

| Format | Test                                                                |
|:-------|:--------------------------------------------------------------------|
| XML    | `<?xml` appears in the first 1024 characters                        |
| HTML   | starts with `<` and HTML tags appear in the first 1024 characters * |
| XML    | starts with `<`                                                     |
| RTF    | starts with `{\rt`                                                  |
| DSV    | starts with `/sep=.$/`, separator is the specified character        |
| DSV    | more unquoted `";"` chars than `"\t"` or `","` in the first 1024    |
| TSV    | more unquoted `"\t"` chars than `","` chars in the first 1024       |
| CSV    | one of the first 1024 characters is a comma `","`                   |
| ETH    | starts with `socialcalc:version:`                                   |
| PRN    | (default)                                                           |

- HTML tags include: `html`, `table`, `head`, `meta`, `script`, `style`, `div`

</details>

<details>
  <summary><b>Why are random text files valid?</b> (click to show)</summary>

Excel is extremely aggressive in reading files.  Adding an XLS extension to any
display text file  (where the only characters are ANSI display chars) tricks
Excel into thinking that the file is potentially a CSV or TSV file, even if it
is only one column!  This library attempts to replicate that behavior.

The best approach is to validate the desired worksheet and ensure it has the
expected number of rows or columns.  Extracting the range is extremely simple:

```js
var range = XLSX.utils.decode_range(worksheet['!ref']);
var ncols = range.e.c - range.s.c + 1, nrows = range.e.r - range.s.r + 1;
```

</details>

