### Processing JSON and JS Data

JSON and JS data tend to represent single worksheets.  This section will use a
few utility functions to generate workbooks.

_Create a new Workbook_

```js
var workbook = XLSX.utils.book_new();
```

The `book_new` utility function creates an empty workbook with no worksheets.

Spreadsheet software generally require at least one worksheet and enforce the
requirement in the user interface.  This library enforces the requirement at
write time, throwing errors if an empty workbook is passed to write functions.


**API**

_Create a worksheet from an array of arrays of JS values_

```js
var worksheet = XLSX.utils.aoa_to_sheet(aoa, opts);
```

The `aoa_to_sheet` utility function walks an "array of arrays" in row-major
order, generating a worksheet object.  The following snippet generates a sheet
with cell `A1` set to the string `A1`, cell `B1` set to `B1`, etc:

```js
var worksheet = XLSX.utils.aoa_to_sheet([
  ["A1", "B1", "C1"],
  ["A2", "B2", "C2"],
  ["A3", "B3", "C3"]
]);
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

**Examples**

["Zen of SheetJS"](#the-zen-of-sheetjs) contains a detailed example "Get Data
from a JSON Endpoint and Generate a Workbook"


[`x-spreadsheet`](https://github.com/myliang/x-spreadsheet) is an interactive
data grid for previewing and modifying structured data in the web browser.  The
[`xspreadsheet` demo](/demos/xspreadsheet) includes a sample script with the
`xtos` function for converting from x-spreadsheet data object to a workbook.
<https://oss.sheetjs.com/sheetjs/x-spreadsheet> is a live demo.

<details>
  <summary><b>Records from a database query (SQL or no-SQL)</b> (click to show)</summary>

The [`database` demo](/demos/database/) includes examples of working with
databases and query results.

</details>


<details>
  <summary><b>Numerical Computations with TensorFlow.js</b> (click to show)</summary>

[`@tensorflow/tfjs`](@tensorflow/tfjs) and other libraries expect data in simple
arrays, well-suited for worksheets where each column is a data vector.  That is
the transpose of how most people use spreadsheets, where each row is a vector.

When recovering data from `tfjs`, the returned data points are stored in a typed
array.  An array of arrays can be constructed with loops. `Array#unshift` can
prepend a title row before the conversion:

```js
const XLSX = require("xlsx");
const tf = require('@tensorflow/tfjs');

/* suppose xs and ys are vectors (1D tensors) -> tfarr will be a typed array */
const tfdata = tf.stack([xs, ys]).transpose();
const shape = tfdata.shape;
const tfarr = tfdata.dataSync();

/* construct the array of arrays */
const aoa = [];
for(let j = 0; j < shape[0]; ++j) {
  aoa[j] = [];
  for(let i = 0; i < shape[1]; ++i) aoa[j][i] = tfarr[j * shape[1] + i];
}
/* add headers to the top */
aoa.unshift(["x", "y"]);

/* generate worksheet */
const worksheet = XLSX.utils.aoa_to_sheet(aoa);
```

The [`array` demo](demos/array/) shows a complete example.

</details>


### Processing HTML Tables

**API**

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

**Examples**

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

<details>
  <summary><b>Server-Side HTML Tables with Headless Chrome</b> (click to show)</summary>

The [`headless` demo](demos/headless/) includes a complete demo to convert HTML
files to XLSB workbooks.  The core idea is to add the script to the page, parse
the table in the page context, generate a `base64` workbook and send it back
for further processing:

```js
const XLSX = require("xlsx");
const { readFileSync } = require("fs"), puppeteer = require("puppeteer");

const url = `https://sheetjs.com/demos/table`;

/* get the standalone build source (node_modules/xlsx/dist/xlsx.full.min.js) */
const lib = readFileSync(require.resolve("xlsx/dist/xlsx.full.min.js"), "utf8");

(async() => {
  /* start browser and go to web page */
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto(url, {waitUntil: "networkidle2"});

  /* inject library */
  await page.addScriptTag({content: lib});

  /* this function `s5s` will be called by the script below, receiving the Base64-encoded file */
  await page.exposeFunction("s5s", async(b64) => {
    const workbook = XLSX.read(b64, {type: "base64" });

    /* DO SOMETHING WITH workbook HERE */
  });

  /* generate XLSB file in webpage context and send back result */
  await page.addScriptTag({content: `
    /* call table_to_book on first table */
    var workbook = XLSX.utils.table_to_book(document.querySelector("TABLE"));

    /* generate XLSX file */
    var b64 = XLSX.write(workbook, {type: "base64", bookType: "xlsb"});

    /* call "s5s" hook exposed from the node process */
    window.s5s(b64);
  `});

  /* cleanup */
  await browser.close();
})();
```

</details>

<details>
  <summary><b>Server-Side HTML Tables with Headless WebKit</b> (click to show)</summary>

The [`headless` demo](demos/headless/) includes a complete demo to convert HTML
files to XLSB workbooks using [PhantomJS](https://phantomjs.org/). The core idea
is to add the script to the page, parse the table in the page context, generate
a `binary` workbook and send it back for further processing:

```js
var XLSX = require('xlsx');
var page = require('webpage').create();

/* this code will be run in the page */
var code = [ "function(){",
  /* call table_to_book on first table */
  "var wb = XLSX.utils.table_to_book(document.body.getElementsByTagName('table')[0]);",

  /* generate XLSB file and return binary string */
  "return XLSX.write(wb, {type: 'binary', bookType: 'xlsb'});",
"}" ].join("");

page.open('https://sheetjs.com/demos/table', function() {
  /* Load the browser script from the UNPKG CDN */
  page.includeJs("https://unpkg.com/xlsx/dist/xlsx.full.min.js", function() {
    /* The code will return an XLSB file encoded as binary string */
    var bin = page.evaluateJavaScript(code);

    var workbook = XLSX.read(bin, {type: "binary"});
    /* DO SOMETHING WITH workbook HERE */

    phantom.exit();
  });
});
```

</details>

<details>
  <summary><b>NodeJS HTML Tables without a browser</b> (click to show)</summary>

NodeJS does not include a DOM implementation and Puppeteer requires a hefty
Chromium build.  [`jsdom`](https://npm.im/jsdom) is a lightweight alternative:

```js
const XLSX = require("xlsx");
const { readFileSync } = require("fs");
const { JSDOM } = require("jsdom");

/* obtain HTML string.  This example reads from test.html */
const html_str = fs.readFileSync("test.html", "utf8");
/* get first TABLE element */
const doc = new JSDOM(html_str).window.document.querySelector("table");
/* generate workbook */
const workbook = XLSX.utils.table_to_book(doc);
```

</details>

