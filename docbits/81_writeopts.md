## Writing Options

The exported `write` and `writeFile` functions accept an options argument:

| Option Name |  Default | Description                                         |
| :---------- | -------: | :-------------------------------------------------- |
|`type`       |          | Output data encoding (see Output Type below)        |
|`cellDates`  |  `false` | Store dates as type `d` (default is `n`)            |
|`bookSST`    |  `false` | Generate Shared String Table **                     |
|`bookType`   | `"xlsx"` | Type of Workbook (see below for supported formats)  |
|`sheet`      |     `""` | Name of Worksheet for single-sheet formats **       |
|`compression`|  `false` | Use ZIP compression for ZIP-based formats **        |
|`Props`      |          | Override workbook properties when writing **        |
|`themeXLSX`  |          | Override theme XML when writing XLSX/XLSB/XLSM **   |

- `bookSST` is slower and more memory intensive, but has better compatibility
  with older versions of iOS Numbers
- The raw data is the only thing guaranteed to be saved.  Features not described
  in this README may not be serialized.
- `cellDates` only applies to XLSX output and is not guaranteed to work with
  third-party readers.  Excel itself does not usually write cells with type `d`
  so non-Excel tools may ignore the data or error in the presence of dates.
- `Props` is an object mirroring the workbook `Props` field.  See the table from
  the [Workbook File Properties](#workbook-file-properties) section.
- if specified, the string from `themeXLSX` will be saved as the primary theme
  for XLSX/XLSB/XLSM files (to `xl/theme/theme1.xml` in the ZIP)

### Supported Output Formats

For broad compatibility with third-party tools, this library supports many
output formats.  The specific file type is controlled with `bookType` option:

| `bookType` | file ext | container | sheets | Description                     |
| :--------- | -------: | :-------: | :----- |:------------------------------- |
| `xlsx`     | `.xlsx`  |    ZIP    | multi  | Excel 2007+ XML Format          |
| `xlsm`     | `.xlsm`  |    ZIP    | multi  | Excel 2007+ Macro XML Format    |
| `xlsb`     | `.xlsb`  |    ZIP    | multi  | Excel 2007+ Binary Format       |
| `biff8`    | `.xls`   |    CFB    | multi  | Excel 97-2004 Workbook Format   |
| `biff5`    | `.xls`   |    CFB    | multi  | Excel 5.0/95 Workbook Format    |
| `biff2`    | `.xls`   |   none    | single | Excel 2.0 Worksheet Format      |
| `xlml`     | `.xls`   |   none    | multi  | Excel 2003-2004 (SpreadsheetML) |
| `ods`      | `.ods`   |    ZIP    | multi  | OpenDocument Spreadsheet        |
| `fods`     | `.fods`  |   none    | multi  | Flat OpenDocument Spreadsheet   |
| `csv`      | `.csv`   |   none    | single | Comma Separated Values          |
| `txt`      | `.txt`   |   none    | single | UTF-16 Unicode Text (TXT)       |
| `sylk`     | `.sylk`  |   none    | single | Symbolic Link (SYLK)            |
| `html`     | `.html`  |   none    | single | HTML Document                   |
| `dif`      | `.dif`   |   none    | single | Data Interchange Format (DIF)   |
| `dbf`      | `.dbf`   |   none    | single | dBASE II + VFP Extensions (DBF) |
| `rtf`      | `.rtf`   |   none    | single | Rich Text Format (RTF)          |
| `prn`      | `.prn`   |   none    | single | Lotus Formatted Text            |

- `compression` only applies to formats with ZIP containers.
- Formats that only support a single sheet require a `sheet` option specifying
  the worksheet.  If the string is empty, the first worksheet is used.
- `writeFile` will automatically guess the output file format based on the file
  extension if `bookType` is not specified.  It will choose the first format in
  the aforementioned table that matches the extension.

### Output Type

The `type` argument for `write` mirrors the `type` argument for `read`:

| `type`     | output                                                          |
|------------|-----------------------------------------------------------------|
| `"base64"` | string: Base64 encoding of the file                             |
| `"binary"` | string: binary string (byte `n` is `data.charCodeAt(n)`)        |
| `"string"` | string: JS string (characters interpreted as UTF8)              |
| `"buffer"` | nodejs Buffer                                                   |
| `"file"`   | string: path of file that will be created (nodejs only)         |

