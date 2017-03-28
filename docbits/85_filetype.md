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
| dBASE II/III/IV / Visual FoxPro (DBF)                        |  :o:  |       |
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

### dBASE and Visual FoxPro (DBF)

DBF is really a typed table format: each column can only hold one data type and
each record omits type information.  The parser generates a header row and
inserts records starting at the second row of the worksheet.

Multi-file extensions like external memos and tables are currently unsupported,
limited by the general ability to read arbitrary files in the web browser.

### Comma-Separated Values

Excel CSV deviates from RFC4180 in a number of important ways.  The generated
CSV files should generally work in Excel although they may not work in RFC4180
compatible readers.

### HTML

Excel HTML worksheets include special metadata encoded in styles.  For example,
`mso-number-format` is a localized string containing the number format.  Despite
the metadata the output is valid HTML, although it does accept bare `&` symbols.

