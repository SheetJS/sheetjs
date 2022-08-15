# VueJS

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

The library can also be imported directly from single-file components with:

```js
// full import
import * as XLSX from 'xlsx';

// named imports
import { read, utils, writeFileXLSX } from 'xlsx';
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

## Mobile Apps

[The new demo](https://docs.sheetjs.com/docs/demos/mobile#quasar) uses the
Quasar Framework in a VueJS + Vite project to generate a native iOS app.

## Nuxt Content

[The new demo](https://docs.sheetjs.com/docs/demos/content#nuxtjs) includes a
complete example starting from `create-nuxt-app`.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
