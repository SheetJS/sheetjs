### Generating JSON and JS Data

JSON and JS data tend to represent single worksheets. The utility functions in
this section work with single worksheets.

The ["Common Spreadsheet Format"](#common-spreadsheet-format) section describes
the object structure in more detail.  `workbook.SheetNames` is an ordered list
of the worksheet names.  `workbook.Sheets` is an object whose keys are sheet
names and whose values are worksheet objects.

The "first worksheet" is stored at `workbook.Sheets[workbook.SheetNames[0]]`.

**API**

_Create an array of JS objects from a worksheet_

```js
var jsa = XLSX.utils.sheet_to_json(worksheet, opts);
```

_Create an array of arrays of JS values from a worksheet_

```js
var aoa = XLSX.utils.sheet_to_json(worksheet, {...opts, header: 1});
```

The `sheet_to_json` utility function walks a workbook in row-major order,
generating an array of objects.  The second `opts` argument controls a number of
export decisions including the type of values (JS values or formatted text). The
["JSON"](#json) section describes the argument in more detail.

By default, `sheet_to_json` scans the first row and uses the values as headers.
With the `header: 1` option, the function exports an array of arrays of values.

**Examples**

[`x-spreadsheet`](https://github.com/myliang/x-spreadsheet) is an interactive
data grid for previewing and modifying structured data in the web browser.  The
[`xspreadsheet` demo](/demos/xspreadsheet) includes a sample script with the
`stox` function for converting from a workbook to x-spreadsheet data object.
<https://oss.sheetjs.com/sheetjs/x-spreadsheet> is a live demo.

<details>
  <summary><b>Previewing data in a React data grid</b> (click to show)</summary>

[`react-data-grid`](https://npm.im/react-data-grid) is a data grid tailored for
react.  It expects two properties: `rows` of data objects and `columns` which
describe the columns.  For the purposes of massaging the data to fit the react
data grid API it is easiest to start from an array of arrays.

This demo starts by fetching a remote file and using `XLSX.read` to extract:

```js
import { useEffect, useState } from "react";
import DataGrid from "react-data-grid";
import { read, utils } from "xlsx";

const url = "https://oss.sheetjs.com/test_files/RkNumber.xls";

export default function App() {
  const [columns, setColumns] = useState([]);
  const [rows, setRows] = useState([]);
  useEffect(() => {(async () => {
    const wb = read(await (await fetch(url)).arrayBuffer(), { WTF: 1 });

    /* use sheet_to_json with header: 1 to generate an array of arrays */
    const data = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });

    /* see react-data-grid docs to understand the shape of the expected data */
    setColumns(data[0].map((r) => ({ key: r, name: r })));
    setRows(data.slice(1).map((r) => r.reduce((acc, x, i) => {
      acc[data[0][i]] = x;
      return acc;
    }, {})));
  })(); });

  return <DataGrid columns={columns} rows={rows} />;
}
```

</details>

<details>
  <summary><b>Previewing data in a VueJS data grid</b> (click to show)</summary>

[`vue3-table-lite`](https://github.com/linmasahiro/vue3-table-lite) is a simple
VueJS 3 data table.  It is featured [in the VueJS demo](/demos/vue/modify/).

</details>

<details>
  <summary><b>Populating a database (SQL or no-SQL)</b> (click to show)</summary>

The [`database` demo](/demos/database/) includes examples of working with
databases and query results.

</details>

<details>
  <summary><b>Numerical Computations with TensorFlow.js</b> (click to show)</summary>

[`@tensorflow/tfjs`](@tensorflow/tfjs) and other libraries expect data in simple
arrays, well-suited for worksheets where each column is a data vector.  That is
the transpose of how most people use spreadsheets, where each row is a vector.

A single `Array#map` can pull individual named rows from `sheet_to_json` export:

```js
const XLSX = require("xlsx");
const tf = require('@tensorflow/tfjs');

const key = "age"; // this is the field we want to pull
const ages = XLSX.utils.sheet_to_json(worksheet).map(r => r[key]);
const tf_data = tf.tensor1d(ages);
```

All fields can be processed at once using a transpose of the 2D tensor generated
with the `sheet_to_json` export with `header: 1`. The first row, if it contains
header labels, should be removed with a slice:

```js
const XLSX = require("xlsx");
const tf = require('@tensorflow/tfjs');

/* array of arrays of the data starting on the second row */
const aoa = XLSX.utils.sheet_to_json(worksheet, {header: 1}).slice(1);
/* dataset in the "correct orientation" */
const tf_dataset = tf.tensor2d(aoa).transpose();
/* pull out each dataset with a slice */
const tf_field0 = tf_dataset.slice([0,0], [1,tensor.shape[1]]).flatten();
const tf_field1 = tf_dataset.slice([1,0], [1,tensor.shape[1]]).flatten();
```

The [`array` demo](demos/array/) shows a complete example.

</details>


### Generating HTML Tables

**API**

_Generate HTML Table from Worksheet_

```js
var html = XLSX.utils.sheet_to_html(worksheet);
```

The `sheet_to_html` utility function generates HTML code based on the worksheet
data.  Each cell in the worksheet is mapped to a `<TD>` element.  Merged cells
in the worksheet are serialized by setting `colspan` and `rowspan` attributes.

**Examples**

The `sheet_to_html` utility function generates HTML code that can be added to
any DOM element by setting the `innerHTML`:

```js
var container = document.getElementById("tavolo");
container.innerHTML = XLSX.utils.sheet_to_html(worksheet);
```

Combining with `fetch`, constructing a site from a workbook is straightforward:

<details>
  <summary><b>Vanilla JS + HTML fetch workbook and generate table previews</b> (click to show)</summary>

```html
<body>
  <style>TABLE { border-collapse: collapse; } TD { border: 1px solid; }</style>
  <div id="tavolo"></div>
  <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
  <script type="text/javascript">
(async() => {
  /* fetch and parse workbook -- see the fetch example for details */
  const workbook = XLSX.read(await (await fetch("sheetjs.xlsx")).arrayBuffer());

  let output = [];
  /* loop through the worksheet names in order */
  workbook.SheetNames.forEach(name => {

    /* generate HTML from the corresponding worksheets */
    const worksheet = workbook.Sheets[name];
    const html = XLSX.utils.sheet_to_html(worksheet);

    /* add a header with the title name followed by the table */
    output.push(`<H3>${name}</H3>${html}`);
  });
  /* write to the DOM at the end */
  tavolo.innerHTML = output.join("\n");
})();
  </script>
</body>
```

</details>

<details>
  <summary><b>React fetch workbook and generate HTML table previews</b> (click to show)</summary>

It is generally recommended to use a React-friendly workflow, but it is possible
to generate HTML and use it in React with `dangerouslySetInnerHTML`:

```jsx
function Tabeller(props) {
  /* the workbook object is the state */
  const [workbook, setWorkbook] = React.useState(XLSX.utils.book_new());

  /* fetch and update the workbook with an effect */
  React.useEffect(() => { (async() => {
    /* fetch and parse workbook -- see the fetch example for details */
    const wb = XLSX.read(await (await fetch("sheetjs.xlsx")).arrayBuffer());
    setWorkbook(wb);
  })(); });

  return workbook.SheetNames.map(name => (<>
    <h3>name</h3>
    <div dangerouslySetInnerHTML={{
      /* this __html mantra is needed to set the inner HTML */
      __html: XLSX.utils.sheet_to_html(workbook.Sheets[name])
    }} />
  </>));
}
```

The [`react` demo](demos/react) includes more React examples.

</details>

<details>
  <summary><b>VueJS fetch workbook and generate HTML table previews</b> (click to show)</summary>

It is generally recommended to use a VueJS-friendly workflow, but it is possible
to generate HTML and use it in VueJS with the `v-html` directive:

```jsx
import { read, utils } from 'xlsx';
import { reactive } from 'vue';

const S5SComponent = {
  mounted() { (async() => {
    /* fetch and parse workbook -- see the fetch example for details */
    const workbook = read(await (await fetch("sheetjs.xlsx")).arrayBuffer());
    /* loop through the worksheet names in order */
    workbook.SheetNames.forEach(name => {
      /* generate HTML from the corresponding worksheets */
      const html = utils.sheet_to_html(workbook.Sheets[name]);
      /* add to state */
      this.wb.wb.push({ name, html });
    });
  })(); },
  /* this state mantra is required for array updates to work */
  setup() { return { wb: reactive({ wb: [] }) }; },
  template: `
  <div v-for="ws in wb.wb" :key="ws.name">
    <h3>{{ ws.name }}</h3>
    <div v-html="ws.html"></div>
  </div>`
};
```

The [`vuejs` demo](demos/vue) includes more React examples.

</details>

### Generating Single-Worksheet Snapshots

The `sheet_to_*` functions accept a worksheet object.

**API**

_Generate a CSV from a single worksheet_

```js
var csv = XLSX.utils.sheet_to_csv(worksheet, opts);
```

This snapshot is designed to replicate the "CSV UTF8 (`.csv`)" output type.
["Delimiter-Separated Output"](#delimiter-separated-output) describes the
function and the optional `opts` argument in more detail.

_Generate "Text" from a single worksheet_

```js
var txt = XLSX.utils.sheet_to_txt(worksheet, opts);
```

This snapshot is designed to replicate the "UTF16 Text (`.txt`)" output type.
["Delimiter-Separated Output"](#delimiter-separated-output) describes the
function and the optional `opts` argument in more detail.

_Generate a list of formulae from a single worksheet_

```js
var fmla = XLSX.utils.sheet_to_formulae(worksheet);
```

This snapshot generates an array of entries representing the embedded formulae.
Array formulae are rendered in the form `range=formula` while plain cells are
rendered in the form `cell=formula or value`.  String literals are prefixed with
an apostrophe `'`, consistent with Excel's formula bar display.

["Formulae Output"](#formulae-output) describes the function in more detail.

