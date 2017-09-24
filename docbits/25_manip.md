## Working with the Workbook

The full object format is described later in this README.

<details>
  <summary><b>Reading a specific cell </b> (click to show)</summary>

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

</details>

<details>
  <summary><b>Adding a new worksheet to a workbook</b> (click to show)</summary>

This example uses [`XLSX.utils.aoa_to_sheet`](#array-of-arrays-input) to make a
worksheet and appends the new worksheet to the workbook:

```js
var new_ws_name = "SheetJS";

/* make worksheet */
var ws_data = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
  [  1 ,  2 ,  3 ,  4 ,  5 ]
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the sheet name to the list */
wb.SheetNames.push(ws_name);

/* Load the worksheet object */
wb.Sheets[ws_name] = ws;

```

</details>


### Parsing and Writing Examples

- <http://sheetjs.com/demos/modify.html> read + modify + write files

- <https://github.com/SheetJS/js-xlsx/blob/master/bin/xlsx.njs> node

The node version installs a command line tool `xlsx` which can read spreadsheet
files and output the contents in various formats.  The source is available at
`xlsx.njs` in the bin directory.

Some helper functions in `XLSX.utils` generate different views of the sheets:

- `XLSX.utils.sheet_to_csv` generates CSV
- `XLSX.utils.sheet_to_html` generates HTML
- `XLSX.utils.sheet_to_json` generates an array of objects
- `XLSX.utils.sheet_to_formulae` generates a list of formulae

