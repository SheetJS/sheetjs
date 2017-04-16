## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description                                          |
| :---------- | ------: | :--------------------------------------------------- |
| type        |         | Input data encoding (see Input Type below)           |
| cellFormula | true    | Save formulae to the .f field                        |
| cellHTML    | true    | Parse rich text and save HTML to the .h field        |
| cellNF      | false   | Save number format string to the .z field            |
| cellStyles  | false   | Save style/theme info to the .s field                |
| cellDates   | false   | Store dates as type `d` (default is `n`)             |
| sheetStubs  | false   | Create cell objects of type `z` for stub cells       |
| sheetRows   | 0       | If >0, read the first `sheetRows` rows **            |
| bookDeps    | false   | If true, parse calculation chains                    |
| bookFiles   | false   | If true, add raw files to book object **             |
| bookProps   | false   | If true, only parse enough to get book metadata **   |
| bookSheets  | false   | If true, only parse enough to get the sheet names    |
| bookVBA     | false   | If true, expose vbaProject.bin to `vbaraw` field **  |
| password    | ""      | If defined and file is encrypted, use password **    |
| WTF         | false   | If true, throw errors on unexpected file features ** |

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
| `0xD0` | CFB Container | BIFF 5/8 or password-protected XLSX/XLSB or WQ3/QPW |
| `0x09` | BIFF Stream   | BIFF 2/3/4/5                                        |
| `0x3C` | XML/HTML      | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0x50` | ZIP Archive   | XLSB or XLSX/M or ODS or UOS2 or plaintext          |
| `0x49` | Plain Text    | SYLK or plaintext                                   |
| `0x54` | Plain Text    | DIF or plaintext                                    |
| `0xFE` | UTF16 Encoded | SpreadsheetML or Flat ODS or UOS1 or plaintext      |
| `0x00` | Record Stream | Lotus WK\* or Quattro Pro or plaintext              |

DBF files are detected based on the first byte as well as the third and fourth
bytes (corresponding to month and day of the file date)

Plaintext format guessing follows the priority order:

| Format | Test                                                                |
|:-------|:--------------------------------------------------------------------|
| HTML   | starts with \<html                                                  |
| XML    | starts with \<                                                      |
| DSV    | starts with `/sep=.$/`, separator is the specified character        |
| TSV    | one of the first 1024 characters is a tab char `"\t"`               |
| CSV    | one of the first 1024 characters is a comma char `","`              |
| PRN    | (default)                                                           |

