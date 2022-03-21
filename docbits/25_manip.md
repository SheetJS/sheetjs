## Processing Data

The ["Common Spreadsheet Format"](#common-spreadsheet-format) is a simple object
representation of the core concepts of a workbook.  The utility functions work
with the object representation and are intended to handle common use cases.

### Modifying Workbook Structure

**API**

_Append a Worksheet to a Workbook_

```js
XLSX.utils.book_append_sheet(workbook, worksheet, sheet_name);
```

The `book_append_sheet` utility function appends a worksheet to the workbook.
The third argument specifies the desired worksheet name. Multiple worksheets can
be added to a workbook by calling the function multiple times.  If the worksheet
name is already used in the workbook, it will throw an error.

_Append a Worksheet to a Workbook and find a unique name_

```js
var new_name = XLSX.utils.book_append_sheet(workbook, worksheet, name, true);
```

If the fourth argument is `true`, the function will start with the specified
worksheet name.  If the sheet name exists in the workbook, a new worksheet name
will be chosen by finding the name stem and incrementing the counter:

```js
XLSX.utils.book_append_sheet(workbook, sheetA, "Sheet2", true); // Sheet2
XLSX.utils.book_append_sheet(workbook, sheetB, "Sheet2", true); // Sheet3
XLSX.utils.book_append_sheet(workbook, sheetC, "Sheet2", true); // Sheet4
XLSX.utils.book_append_sheet(workbook, sheetD, "Sheet2", true); // Sheet5
```

_List the Worksheet names in tab order_

```js
var wsnames = workbook.SheetNames;
```

The `SheetNames` property of the workbook object is a list of the worksheet
names in "tab order".  API functions will look at this array.

_Replace a Worksheet in place_

```js
workbook.Sheets[sheet_name] = new_worksheet;
```

The `Sheets` property of the workbook object is an object whose keys are names
and whose values are worksheet objects.  By reassigning to a property of the
`Sheets` object, the worksheet object can be changed without disrupting the
rest of the worksheet structure.

**Examples**

<details>
  <summary><b>Add a new worksheet to a workbook</b> (click to show)</summary>

This example uses [`XLSX.utils.aoa_to_sheet`](#array-of-arrays-input).

```js
var ws_name = "SheetJS";

/* Create worksheet */
var ws_data = [
  [ "S", "h", "e", "e", "t", "J", "S" ],
  [  1 ,  2 ,  3 ,  4 ,  5 ]
];
var ws = XLSX.utils.aoa_to_sheet(ws_data);

/* Add the worksheet to the workbook */
XLSX.utils.book_append_sheet(wb, ws, ws_name);
```

</details>

### Modifying Cell Values

**API**

_Modify a single cell value in a worksheet_

```js
XLSX.utils.sheet_add_aoa(worksheet, [[new_value]], { origin: address });
```

_Modify multiple cell values in a worksheet_

```js
XLSX.utils.sheet_add_aoa(worksheet, aoa, opts);
```

The `sheet_add_aoa` utility function modifies cell values in a worksheet.  The
first argument is the worksheet object.  The second argument is an array of
arrays of values.  The `origin` key of the third argument controls where cells
will be written.  The following snippet sets `B3=1` and `E5="abc"`:

```js
XLSX.utils.sheet_add_aoa(worksheet, [
  [1],                             // <-- Write 1 to cell B3
  ,                                // <-- Do nothing in row 4
  [/*B5*/, /*C5*/, /*D5*/, "abc"]  // <-- Write "abc" to cell E5
], { origin: "B3" });
```

["Array of Arrays Input"](#array-of-arrays-input) describes the function and the
optional `opts` argument in more detail.

**Examples**

<details>
  <summary><b>Appending rows to a worksheet</b> (click to show)</summary>

The special origin value `-1` instructs `sheet_add_aoa` to start in column A of
the row after the last row in the range, appending the data:

```js
XLSX.utils.sheet_add_aoa(worksheet, [
  ["first row after data", 1],
  ["second row after data", 2]
], { origin: -1 });
```

</details>


### Modifying Other Worksheet / Workbook / Cell Properties

The ["Common Spreadsheet Format"](#common-spreadsheet-format) section describes
the object structures in greater detail.

