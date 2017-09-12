# VueJS 2

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags e.g.

```html
<script src="xlsx.full.min.js"></script>
```

Strictly speaking, there should be no need for a Vue.JS demo!  You can proceed
as you would with any other browser-friendly library.

This demo directly generates HTML using `sheet_to_html` and adds an element to
a pregenerated template.  It also has a button for exporting as XLSX.

Other scripts in this demo show:
- server-rendered VueJS component (with `nuxt.js`)
- `weex` deployment for iOS

## Single File Components

For Single File Components, a simple `import XLSX from 'xlsx'` should suffice.
The webpack demo includes a sample `webpack.config.js`.

## WeeX

WeeX is a framework for building real mobile apps, akin to React Native.  The
ecosystem is not quite as mature as React Native, missing basic features like
document access.  As a result, this demo uses the `stream.fetch` API to upload
Base64-encoded documents to <https://hastebin.com> and download a precomputed
[Base64-encoded workbook](http://sheetjs.com/sheetjs.xlsx.b64).

Using NodeJS it is straightforward to convert to/from base64:

```js
/* convert sheetjs.xlsx -> sheetjs.xlsx.b64 */
var buf = fs.readFileSync("sheetjs.xlsx");
fs.writeFileSync("sheetjs.xlsx.b64", buf.toString("base64"));

/* convert sheetjs.xls.b64 -> sheetjs.xls */
var str = fs.readFileSync("sheetjs.xls.b64").toString();
fs.writeFileSync("sheetjs.xls", new Buffer(str, "base64"));
```

## Nuxt and State

The `nuxt.js` demo uses the same state approach as the React next.js demo:

```js
{
  cols: [
    { name: "A", key: 0 },
    { name: "B", key: 1 },
    { name: "C", key: 2 },
  ],
  data: [
    [ "id",    "name", "value" ],
    [    1, "sheetjs",    7262 ]
    [    2, "js-xlsx",    6969 ]
  ]
}
```

Due to webpack configuration issues on client/server bundles, the library should
be explicitly included in the layout HTML (as script tag) and in the component:

```js
const _XLSX = require('xlsx');
const X = typeof XLSX !== 'undefined' ? XLSX : _XLSX;
/* use the variable X rather than XLSX in the component */
```
