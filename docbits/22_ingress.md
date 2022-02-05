### Processing JSON and JS Data

JSON and JS data tend to represent single worksheets.  This section will use a
few utility functions to generate workbooks:

_Create a new Worksheet_

```js
var workbook = XLSX.utils.book_new();
```

The `book_new` utility function creates an empty workbook with no worksheets.


_Append a Worksheet to a Workbook_

```js
XLSX.utils.book_append_sheet(workbook, worksheet, sheet_name);
```

The `book_append_sheet` utility function appends a worksheet to the workbook.
The third argument specifies the desired worksheet name. Multiple worksheets can
be added to a workbook by calling the function multiple times.


#### API

_Create a worksheet from an array of arrays of JS values_

```js
var worksheet = XLSX.utils.aoa_to_sheet(aoa, opts);
```

The `aoa_to_sheet` utility function walks an "array of arrays" in row-major
order, generating a worksheet object.  The following snippet generates a sheet
with cell `A1` set to the string `A1`, cell `B1` set to `B2`, etc:

```js
var worksheet = XLSX.utils.aoa_to_sheet([
  ["A1", "B1", "C1"],
  ["A2", "B2", "C2"],
  ["A3", "B3", "C3"]
])
```

["Array of Arrays Input"](#array-of-arrays-input) describes the function and the
optional `opts` argument in more detail.


_Create a worksheet from an array of JS objects_

```js
var worksheet = XLSX.utils.json_to_sheet(jsa, opts);
```

The `json_to_sheet` utility function walks an array of JS objects in order,
generating a worksheet object.  By default, it will generate a header row and
one row per object in the array.  The optional `opts` argument has settings to
control the column order and header output.

["Array of Objects Input"](#array-of-arrays-input) describes the function and
the optional `opts` argument in more detail.

#### Examples

["Zen of SheetJS"](#the-zen-of-sheetjs) contains a detailed example "Get Data
from a JSON Endpoint and Generate a Workbook"

The [`database` demo](/demos/database/) includes examples of working with
databases and query results.

### Processing HTML Tables

#### API

_Create a worksheet by scraping an HTML TABLE in the page_

```js
var worksheet = XLSX.utils.table_to_sheet(dom_element, opts);
```

The `table_to_sheet` utility function takes a DOM TABLE element and iterates
through the rows to generate a worksheet.  The `opts` argument is optional.
["HTML Table Input"](#html-table-input) describes the function in more detail.



_Create a workbook by scraping an HTML TABLE in the page_

```js
var workbook = XLSX.utils.table_to_book(dom_element, opts);
```

The `table_to_book` utility function follows the same logic as `table_to_sheet`.
After generating a worksheet, it creates a blank workbook and appends the
spreadsheet.

The options argument supports the same options as `table_to_sheet`, with the
addition of a `sheet` property to control the worksheet name.  If the property
is missing or no options are specified, the default name `Sheet1` is used.

#### Examples

Here are a few common scenarios (click on each subtitle to see the code):

<details>
  <summary><b>HTML TABLE element in a webpage</b> (click to show)</summary>

```html
<!-- include the standalone script and shim.  this uses the UNPKG CDN -->
<script src="https://unpkg.com/xlsx/dist/shim.min.js"></script>
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>

<!-- example table with id attribute -->
<table id="tableau">
  <tr><td>Sheet</td><td>JS</td></tr>
  <tr><td>12345</td><td>67</td></tr>
</table>

<!-- this block should appear after the table HTML and the standalone script -->
<script type="text/javascript">
  var workbook = XLSX.utils.table_to_book(document.getElementById("tableau"));

  /* DO SOMETHING WITH workbook HERE */
</script>
```

Multiple tables on a web page can be converted to individual worksheets:

```js
/* create new workbook */
var workbook = XLSX.utils.book_new();

/* convert table "table1" to worksheet named "Sheet1" */
var sheet1 = XLSX.utils.table_to_sheet(document.getElementById("table1"));
XLSX.utils.book_append_sheet(workbook, sheet1, "Sheet1");

/* convert table "table2" to worksheet named "Sheet2" */
var sheet2 = XLSX.utils.table_to_sheet(document.getElementById("table2"));
XLSX.utils.book_append_sheet(workbook, sheet2, "Sheet2");

/* workbook now has 2 worksheets */
```

Alternatively, the HTML code can be extracted and parsed:

```js
var htmlstr = document.getElementById("tableau").outerHTML;
var workbook = XLSX.read(htmlstr, {type:"string"});
```

</details>

<details>
  <summary><b>Chrome/Chromium Extension</b> (click to show)</summary>

The [`chrome` demo](demos/chrome/) shows a complete example and details the
required permissions and other settings.

In an extension, it is recommended to generate the workbook in a content script
and pass the object back to the extension:

```js
/* in the worker script */
chrome.runtime.onMessage.addListener(function(msg, sender, cb) {
  /* pass a message like { sheetjs: true } from the extension to scrape */
  if(!msg || !msg.sheetjs) return;
  /* create a new workbook */
  var workbook = XLSX.utils.book_new();
  /* loop through each table element */
  var tables = document.getElementsByTagName("table")
  for(var i = 0; i < tables.length; ++i) {
    var worksheet = XLSX.utils.table_to_sheet(tables[i]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Table" + i);
  }
  /* pass back to the extension */
  return cb(workbook);
});
```

</details>

