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
var ws = workbook.Sheets["Sheet1"];
XLSX.utils.sheet_add_aoa(ws, [["Created "+new Date().toISOString()]], {origin:-1});

// Package and Release Data (`writeFile` tries to write and save an XLSB file)
XLSX.writeFile(workbook, "Report.xlsb");
```

This library tries to simplify steps 2 and 4 with functions to extract useful
data from spreadsheet files (`read` / `readFile`) and generate new spreadsheet
files from data (`write` / `writeFile`).  Additional utility functions like
`table_to_book` work with other common data sources like HTML tables.

This documentation and various demo projects cover a number of common scenarios
and approaches for steps 1 and 5.

Utility functions help with step 3.

["Acquiring and Extracting Data"](#acquiring-and-extracting-data) describes
solutions for common data import scenarios.

["Packaging and Releasing Data"](#packaging-and-releasing-data) describes
solutions for common data export scenarios.

["Processing Data"](#packaging-and-releasing-data) describes solutions for
common workbook processing and manipulation scenarios.

["Utility Functions"](#utility-functions) details utility functions for
translating JSON Arrays and other common JS structures into worksheet objects.

### The Zen of SheetJS

_Data processing should fit in any workflow_

The library does not impose a separate lifecycle.  It fits nicely in websites
and apps built using any framework.  The plain JS data objects play nice with
Web Workers and future APIs.

_JavaScript is a powerful language for data processing_

The ["Common Spreadsheet Format"](#common-spreadsheet-format) is a simple object
representation of the core concepts of a workbook.  The various functions in the
library provide low-level tools for working with the object.

For friendly JS processing, there are utility functions for converting parts of
a worksheet to/from an Array of Arrays.  The following example combines powerful
JS Array methods with a network request library to download data, select the
information we want and create a workbook file:

<details>
  <summary><b>Get Data from a JSON Endpoint and Generate a Workbook</b> (click to show)</summary>

The goal is to generate a XLSB workbook of US President names and birthdays.

**Acquire Data**

_Raw Data_

<https://theunitedstates.io/congress-legislators/executive.json> has the desired
data.  For example, John Adams:

```js
{
  "id": { /* (data omitted) */ },
  "name": {
    "first": "John",          // <-- first name
    "last": "Adams"           // <-- last name
  },
  "bio": {
    "birthday": "1735-10-19", // <-- birthday
    "gender": "M"
  },
  "terms": [
    { "type": "viceprez", /* (other fields omitted) */ },
    { "type": "viceprez", /* (other fields omitted) */ },
    { "type": "prez", /* (other fields omitted) */ } // <-- look for "prez"
  ]
}
```

_Filtering for Presidents_

The dataset includes Aaron Burr, a Vice President who was never President!

`Array#filter` creates a new array with the desired rows.  A President served
at least one term with `type` set to `"prez"`.  To test if a particular row has
at least one `"prez"` term, `Array#some` is another native JS function.  The
complete filter would be:

```js
const prez = raw_data.filter(row => row.terms.some(term => term.type === "prez"));
```

_Lining up the data_

For this example, the name will be the first name combined with the last name
(`row.name.first + " " + row.name.last`) and the birthday will be the subfield
`row.bio.birthday`.  Using `Array#map`, the dataset can be massaged in one call:

```js
const rows = prez.map(row => ({
  name: row.name.first + " " + row.name.last,
  birthday: row.bio.birthday
}));
```

The result is an array of "simple" objects with no nesting:

```js
[
  { name: "George Washington", birthday: "1732-02-22" },
  { name: "John Adams", birthday: "1735-10-19" },
  // ... one row per President
]
```

**Extract Data**

With the cleaned dataset, `XLSX.utils.json_to_sheet` generates a worksheet:

```js
const worksheet = XLSX.utils.json_to_sheet(rows);
```

`XLSX.utils.book_new` creates a new workbook and `XLSX.utils.book_append_sheet`
appends a worksheet to the workbook. The new worksheet will be called "Dates":

```js
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");
```

**Process Data**

_Fixing headers_

By default, `json_to_sheet` creates a worksheet with a header row. In this case,
the headers come from the JS object keys: "name" and "birthday".

The headers are in cells A1 and B1.  `XLSX.utils.sheet_add_aoa` can write text
values to the existing worksheet starting at cell A1:

```js
XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });
```

_Fixing Column Widths_

Some of the names are longer than the default column width.  Column widths are
set by [setting the `"!cols"` worksheet property](#row-and-column-properties).

The following line sets the width of column A to approximately 10 characters:

```js
worksheet["!cols"] = [ { wch: 10 } ]; // set column A width to 10 characters
```

One `Array#reduce` call over `rows` can calculate the maximum width:

```js
const max_width = rows.reduce((w, r) => Math.max(w, r.name.length), 10);
worksheet["!cols"] = [ { wch: max_width } ];
```

Note: If the starting point was a file or HTML table, `XLSX.utils.sheet_to_json`
will generate an array of JS objects.

**Package and Release Data**

`XLSX.writeFile` creates a spreadsheet file and tries to write it to the system.
In the browser, it will try to prompt the user to download the file.  In NodeJS,
it will write to the local directory.

```js
XLSX.writeFile(workbook, "Presidents.xlsx");
```

**Complete Example**

```js
// Uncomment the next line for use in NodeJS:
// const XLSX = require("xlsx"), axios = require("axios");

(async() => {
  /* fetch JSON data and parse */
  const url = "https://theunitedstates.io/congress-legislators/executive.json";
  const raw_data = (await axios(url, {responseType: "json"})).data;

  /* filter for the Presidents */
  const prez = raw_data.filter(row => row.terms.some(term => term.type === "prez"));

  /* flatten objects */
  const rows = prez.map(row => ({
    name: row.name.first + " " + row.name.last,
    birthday: row.bio.birthday
  }));

  /* generate worksheet and workbook */
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");

  /* fix headers */
  XLSX.utils.sheet_add_aoa(worksheet, [["Name", "Birthday"]], { origin: "A1" });

  /* calculate column width */
  const max_width = rows.reduce((w, r) => Math.max(w, r.name.length), 10);
  worksheet["!cols"] = [ { wch: max_width } ];

  /* create an XLSX file and try to save to Presidents.xlsx */
  XLSX.writeFile(workbook, "Presidents.xlsx");
})();
```

For use in the web browser, assuming the snippet is saved to `snippet.js`,
script tags should be used to include the `axios` and `xlsx` standalone builds:

```html
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
<script src="https://unpkg.com/axios/dist/axios.min.js"></script>
<script src="snippet.js"></script>
```


</details>

_File formats are implementation details_

The parser covers a wide gamut of common spreadsheet file formats to ensure that
"HTML-saved-as-XLS" files work as well as actual XLS or XLSX files.

The writer supports a number of common output formats for broad compatibility
with the data ecosystem.

To the greatest extent possible, data processing code should not have to worry
about the specific file formats involved.


