## File Formats

Despite the library name `xlsx`, it supports numerous spreadsheet file formats:

| Format                                                       | Read  | Write |
|:-------------------------------------------------------------|:-----:|:-----:|
| **Excel Worksheet/Workbook Formats**                         |:-----:|:-----:|
| Excel 2007+ XML Formats (XLSX/XLSM)                          |  :o:  |  :o:  |
| Excel 2007+ Binary Format (XLSB BIFF12)                      |  :o:  |  :o:  |
| Excel 2003-2004 XML Format (XML "SpreadsheetML")             |  :o:  |  :o:  |
| Excel 97-2004 (XLS BIFF8)                                    |  :o:  |  :o:  |
| Excel 5.0/95 (XLS BIFF5)                                     |  :o:  |       |
| Excel 4.0 (XLS/XLW BIFF4)                                    |  :o:  |       |
| Excel 3.0 (XLS BIFF3)                                        |  :o:  |       |
| Excel 2.0/2.1 (XLS BIFF2)                                    |  :o:  |  :o:  |
| **Excel Supported Text Formats**                             |:-----:|:-----:|
| Delimiter-Separated Values (CSV/TXT)                         |  :o:  |  :o:  |
| Data Interchange Format (DIF)                                |  :o:  |  :o:  |
| Symbolic Link (SYLK/SLK)                                     |  :o:  |  :o:  |
| Lotus Formatted Text (PRN)                                   |  :o:  |  :o:  |
| UTF-16 Unicode Text (TXT)                                    |  :o:  |  :o:  |
| **Other Workbook/Worksheet Formats**                         |:-----:|:-----:|
| OpenDocument Spreadsheet (ODS)                               |  :o:  |  :o:  |
| Flat XML ODF Spreadsheet (FODS)                              |  :o:  |  :o:  |
| Uniform Office Format Spreadsheet (标文通 UOS1/UOS2)         |  :o:  |       |
| dBASE II/III/IV / Visual FoxPro (DBF)                        |  :o:  |       |
| Lotus 1-2-3 (WKS/WK1/WK2/WK3/WK4/123)                        |  :o:  |       |
| Quattro Pro Spreadsheet (WQ1/WQ2/WB1/WB2/WB3/QPW)            |  :o:  |       |
| **Other Common Spreadsheet Output Formats**                  |:-----:|:-----:|
| HTML Tables                                                  |  :o:  |  :o:  |

### Excel 2007+ XML (XLSX/XLSM)

<details>
	<summary>(click to show)</summary>

XLSX and XLSM files are ZIP containers containing a series of XML files in
accordance with the Open Packaging Conventions (OPC).  The XLSM filetype, almost
identical to XLSX, is used for files containing macros.

The format is standardized in ECMA-376 and later in ISO/IEC 29500.  Excel does
not follow the specification, and there are additional documents discussing how
Excel deviates from the specification.

</details>

### Excel 2.0-95 (BIFF2/BIFF3/BIFF4/BIFF5)

<details>
	<summary>(click to show)</summary>

BIFF 2/3 XLS are single-sheet streams of binary records.  Excel 4 introduced
the concept of a workbook (`XLW` files) but also had single-sheet `XLS` format.
The structure is largely similar to the Lotus 1-2-3 file formats.  BIFF5/8/12
extended the format in various ways but largely stuck to the same record format.

There is no official specification for any of these formats.  Excel 95 can write
files in these formats, so record lengths and fields were backsolved by writing
in all of the supported formats and comparing files.  Excel 2016 can generate
BIFF5 files, enabling a full suite of file tests starting from XLSX or BIFF2.

</details>

### Excel 97-2004 Binary (BIFF8)

<details>
	<summary>(click to show)</summary>

BIFF8 exclusively uses the Compound File Binary container format, splitting some
content into streams within the file.  At its core, it still uses an extended
version of the binary record format from older versions of BIFF.

The `MS-XLS` specification covers the basics of the file format, and other
specifications expand on serialization of features like properties.

</details>

### Excel 2003-2004 (SpreadsheetML)

<details>
	<summary>(click to show)</summary>

Predating XLSX, SpreadsheetML files are simple XML files.  There is no official
and comprehensive specification, although MS has released whitepapers on the
format.  Since Excel 2016 can generate SpreadsheetML files, backsolving is
pretty straightforward.

</details>

### Excel 2007+ Binary (XLSB, BIFF12)

<details>
	<summary>(click to show)</summary>

Introduced in parallel with XLSX, the XLSB filetype combines BIFF architecture
with the content separation and ZIP container of XLSX.  For the most part nodes
in an XLSX sub-file can be mapped to XLSB records in a corresponding sub-file.

The `MS-XLSB` specification covers the basics of the file format, and other
specifications expand on serialization of features like properties.

</details>

### Delimiter-Separated Values (CSV/TXT)

<details>
	<summary>(click to show)</summary>

Excel CSV deviates from RFC4180 in a number of important ways.  The generated
CSV files should generally work in Excel although they may not work in RFC4180
compatible readers.  The parser should generally understand Excel CSV. The
writer proactively generates cells for formulae if values are unavailable.

Excel TXT uses tab as the delimiter and codepage 1200.

Notes:

- Like in Excel, files starting with `0x49 0x44 ("ID")` are treated as Symbolic
  Link files.  Unlike Excel, if the file does not have a valid SYLK header, it
  will be proactively reinterpreted as CSV.  There are some files with semicolon
  delimiter that align with a valid SYLK file.  For the broadest compatibility,
  all cells with the value of `ID` are automatically wrapped in double-quotes.

</details>

### Other Workbook Formats

<details>
	<summary>(click to show)</summary>

Support for other formats is generally far XLS/XLSB/XLSX support, due in large
part to a lack of publicly available documentation.  Test files were produced in
the respective apps and compared to their XLS exports to determine structure.
The main focus is data extraction.

</details>

#### Lotus 1-2-3 (WKS/WK1/WK2/WK3/WK4/123)

<details>
	<summary>(click to show)</summary>

The Lotus formats consist of binary records similar to the BIFF structure. Lotus
did release a whitepaper decades ago covering the original WK1 format.  Other
features were deduced by producing files and comparing to Excel support.

</details>

#### Quattro Pro (WQ1/WQ2/WB1/WB2/WB3/QPW)

<details>
	<summary>(click to show)</summary>

The Quattro Pro formats use binary records in the same way as BIFF and Lotus.
Some of the newer formats (namely WB3 and QPW) use a CFB enclosure just like
BIFF8 XLS.

</details>

#### OpenDocument Spreadsheet (ODS/FODS)

<details>
	<summary>(click to show)</summary>

ODS is an XML-in-ZIP format akin to XLSX while FODS is an XML format akin to
SpreadsheetML.  Both are detailed in the OASIS standard, but tools like LO/OO
add undocumented extensions.  The parsers and writers do not implement the full
standard, instead focusing on parts necessary to extract and store raw data.

</details>

#### Uniform Office Spreadsheet (UOS1/2)

<details>
	<summary>(click to show)</summary>

UOS is a very similar format, and it comes in 2 varieties corresponding to ODS
and FODS respectively.  For the most part, the difference between the formats
is in the names of tags and attributes.

</details>

### Other Single-Worksheet Formats

Many older formats supported only one worksheet:

#### dBASE and Visual FoxPro (DBF)

<details>
	<summary>(click to show)</summary>

DBF is really a typed table format: each column can only hold one data type and
each record omits type information.  The parser generates a header row and
inserts records starting at the second row of the worksheet.

Multi-file extensions like external memos and tables are currently unsupported,
limited by the general ability to read arbitrary files in the web browser.

</details>

#### Symbolic Link (SYLK)

<details>
	<summary>(click to show)</summary>

There is no real documentation.  All knowledge was gathered by saving files in
various versions of Excel to deduce the meaning of fields.  Notes:

- Plain formulae are stored in the RC form.
- Column widths are rounded to integral characters.

</details>

#### Lotus Formatted Text (PRN)

<details>
	<summary>(click to show)</summary>

There is no real documentation, and in fact Excel treats PRN as an output-only
file format.  Nevertheless we can guess the column widths and reverse-engineer
the original layout.  Excel's 240-character width limitation is not enforced.

</details>

#### Data Interchange Format (DIF)

<details>
	<summary>(click to show)</summary>

There is no unified definition.  Visicalc DIF differs from Lotus DIF, and both
differ from Excel DIF.  Where ambiguous, the parser/writer follows the expected
behavior from Excel.  In particular, Excel extends DIF in incompatible ways:

- Since Excel automatically converts numbers-as-strings to numbers, numeric
  string constants are converted to formulae: `"0.3" -> "=""0.3""`
- DIF technically expects numeric cells to hold the raw numeric data, but Excel
  permits formatted numbers (including dates)
- DIF technically has no support for formulae, but Excel will automatically
  convert plain formulae.  Array formulae are not preserved.

</details>

#### HTML

<details>
	<summary>(click to show)</summary>

Excel HTML worksheets include special metadata encoded in styles.  For example,
`mso-number-format` is a localized string containing the number format.  Despite
the metadata the output is valid HTML, although it does accept bare `&` symbols.

</details>

