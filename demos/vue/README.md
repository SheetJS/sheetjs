# VueJS 2

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

The library can also be imported directly from single-file components with:

```js
import XLSX from 'xlsx';
```

This demo directly generates HTML using `sheet_to_html` and adds an element to
a pre-generated template.  It also has a button for exporting as XLSX.

Other scripts in this demo show:
- server-rendered VueJS component (with `nuxt.js`)
- `weex` deployment for iOS

## Internal State

The plain JS demo embeds state in the DOM.  Other demos use proper state.

The simplest state representation is an array of arrays.  To avoid having the
table component depend on the library, the column labels are precomputed.  The
state in this demo is shaped like the following object:

```js
{
  cols: [{ name: "A", key: 0 }, { name: "B", key: 1 }, { name: "C", key: 2 }],
  data: [
    [ "id",    "name", "value" ],
    [    1, "sheetjs",    7262 ],
    [    2, "js-xlsx",    6969 ]
  ]
}
```

`sheet_to_json` and `aoa_to_sheet` utility functions can convert between arrays
of arrays and worksheets:

```js
/* convert from workbook to array of arrays */
var first_worksheet = workbook.Sheets[workbook.SheetNames[0]];
var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});

/* convert from array of arrays to workbook */
var worksheet = XLSX.utils.aoa_to_sheet(data);
var new_workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(new_workbook, worksheet, "SheetJS");
```

The column objects can be generated with the `encode_col` utility function:

```js
function make_cols(refstr/*:string*/) {
  var o = [];
  var range = XLSX.utils.decode_range(refstr);
  for(var i = 0; i <= range.e.c; ++i) {
    o.push({name: XLSX.utils.encode_col(i), key:i});
  }
  return o;
}
```

## WeeX

<img src="screen.png" width="400px"/>

Reproducing the full project is a little bit tricky.  The included `weex.sh`
script performs the necessary installation steps.

WeeX is a framework for building real mobile apps, akin to React Native.  The
ecosystem is not quite as mature as React Native, missing basic features like
document access.  As a result, this demo uses the `stream.fetch` API to upload
Base64-encoded documents to <https://hastebin.com> and download a precomputed
[Base64-encoded workbook](http://sheetjs.com/sheetjs.xlsx.b64).

Using NodeJS it is straightforward to convert to/from Base64:

```js
/* convert sheetjs.xlsx -> sheetjs.xlsx.b64 */
var buf = fs.readFileSync("sheetjs.xlsx");
fs.writeFileSync("sheetjs.xlsx.b64", buf.toString("base64"));

/* convert sheetjs.xls.b64 -> sheetjs.xls */
var str = fs.readFileSync("sheetjs.xls.b64").toString();
fs.writeFileSync("sheetjs.xls", new Buffer(str, "base64"));
```

## Other Demos

#### Server-Rendered VueJS Components with Nuxt.js

Due to webpack configuration issues on client/server bundles, the library should
be explicitly included in the layout HTML (as script tag) and in the component:

```js
const _XLSX = require('xlsx');
const X = typeof XLSX !== 'undefined' ? XLSX : _XLSX;
/* use the variable X rather than XLSX in the component */
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
