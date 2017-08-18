## Parsing Options

The exported `read` and `readFile` functions accept an options argument:

| Option Name | Default | Description                                          |
| :---------- | ------: | :--------------------------------------------------- |
| type        |         | Input data encoding (see Input Type below)           |
| raw         | false   | If true, plaintext parsing will not parse values **  |
| cellFormula | true    | Save formulae to the .f field                        |
| cellHTML    | true    | Parse rich text and save HTML to the `.h` field      |
| cellNF      | false   | Save number format string to the `.z` field          |
| cellStyles  | false   | Save style/theme info to the `.s` field              |
| cellText    | true    | Generated formatted text to the `.w` field           |
| cellDates   | false   | Store dates as type `d` (default is `n`)             |
| dateNF      |         | If specified, use the string for date code 14 **     |
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
- Excel aggressively tries to interpret values from CSV and other plaintext.
  This leads to surprising behavior! The `raw` option suppresses value parsing.
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
| `0x3C` | XML/HTML      | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0x50` | ZIP Archive   | XLSB or XLSX/M or ODS or UOS2 or plaintext          |
| `0x49` | Plain Text    | SYLK or plaintext                                   |
| `0x54` | Plain Text    | DIF or plaintext                                    |
| `0xEF` | UTF8 Encoded  | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0xFF` | UTF16 Encoded | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0x00` | Record Stream | Lotus WK\* or Quattro Pro or plaintext              |
| `0x0A` | Plaintext     | RTF or plaintext                                    |
| `0x0A` | Plaintext     | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0x0D` | Plaintext     | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |
| `0x20` | Plaintext     | SpreadsheetML / Flat ODS / UOS1 / HTML / plaintext  |

DBF files are detected based on the first byte as well as the third and fourth
bytes (corresponding to month and day of the file date)

Plaintext format guessing follows the priority order:

| Format | Test                                                                |
|:-------|:--------------------------------------------------------------------|
| XML    | `<?xml` appears in the first 1024 characters                        |
| HTML   | starts with `<` and HTML tags appear in the first 1024 characters * |
| XML    | starts with `<`                                                     |
| RTF    | starts with `{\rt`                                                  |
| DSV    | starts with `/sep=.$/`, separator is the specified character        |
| CSV    | more unquoted `","` characters than `"\t"` chars in the first 1024  |
| TSV    | one of the first 1024 characters is a tab char `"\t"`               |
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
var ncols = range.e.c - range.r.c + 1, nrows = range.e.r - range.s.r + 1;
```

</details>

