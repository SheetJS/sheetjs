### Usage

Most scenarios involving spreadsheets and data can be broken into 5 parts:

1) **Acquire Data**:  Data may be stored anywhere: local or remote files,
   databases, HTML TABLE, or even generated programmatically in the web browser.

2) **Extract Data**:  For spreadsheet files, this involves parsing raw bytes to
   read the cell data. For general JS data, this involves reshaping the data.

3) **Process Data**:  From generating summary statistics to cleaning data
   records, this step is the heart of the problem.

4) **Package Data**:  This can involve making a new spreadsheet or serializing
   with `JSON.stringify` or writing XML or simply flattening data for UI tools.

5) **Release Data**:  Spreadsheet files can be uploaded to a server or written
   locally.  Data can be presented to users in an HTML TABLE or data grid.

A common problem involves generating a valid spreadsheet export from data stored
in an HTML table.  In this example, an HTML TABLE on the page will be scraped,
a row will be added to the bottom with the date of the report, and a new file
will be generated and downloaded locally. `XLSX.writeFile` takes care of
packaging the data and attempting a local download:

```js
// Acquire Data (reference to the HTML table)
var table_elt = document.getElementById("my-table-id");

// Extract Data (create a workbook object from the table)
var workbook = XLSX.utils.table_to_book(table_elt);

// Process Data (add a new row)
var worksheet = workbook.Sheets["Sheet1"];
XLSX.utils.sheet_add_aoa([["Created "+new Date().toISOString()}]], {origin:-1});

// Package and Release Data (`writeFile` tries to write and save an XLSB file)
XLSX.writeFile(workbook, "Report.xlsb");
```

This library tries to simplify steps 2 and 4 with functions to extract useful
data from spreadsheet files (`read` / `readFile`) and generate new spreadsheet
files from data (`write` / `writeFile`).

This documentation and various demo projects cover a number of common scenarios
and approaches for steps 1 and 5.

Utility functions help with step 3.


#### The Zen of SheetJS


_File formats are implementation details_

The parser covers a wide gamut of common spreadsheet file formats to ensure that
"HTML-saved-as-XLS" files work as well as actual XLS or XLSX files.

The writer supports a number of common output formats for broad compatibility
with the data ecosystem.


_Data processing should fit in any workflow_

The library does not impose a separate lifecycle.  It fits nicely in websites
and apps built using any framework.  The plain JS data objects play nice with
Web Workers and future APIs.

["Parsing Workbooks"](#parsing-workbooks) describes solutions for common data
import scenarios involving actual spreadsheet files.

["Writing Workbooks"](#writing-workbooks) describes solutions for common data
export scenarios involving actual spreadsheet files.

["Utility Functions"](#utility-functions) details utility functions for
translating JSON Arrays and other common JS structures into worksheet objects.


_JavaScript is a powerful language for data processing_

The ["Common Spreadsheet Format"](#common-spreadsheet-format) is a simple object
representation of the core concepts of a workbook.  The various functions in the
library provide low-level tools for working with the object.

For friendly JS processing, there are utility functions for converting parts of
a worksheet to/from an Array of Arrays.  For example, summing columns from an
array of arrays can be implemented in a single Array reduce operation:

```js
var aoa = XLSX.utils.sheet_to_json(worksheet, {header: 1});
var sum_of_column_B = aoa.reduce((acc, row) => acc + (+row[1]||0), 0);
```


