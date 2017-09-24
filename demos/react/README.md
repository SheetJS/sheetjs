# React

The `xlsx.core.min.js` and `xlsx.full.min.js` scripts are designed to be dropped
into web pages with script tags:

```html
<script src="xlsx.full.min.js"></script>
```

The library can also be imported directly from JSX code with:

```js
import * as XLSX from 'xlsx';
```

This demo shows a simple JSX component transpiled in the browser using the babel
standalone library.  Since there is no standard React table model, this demo
settles on the array of arrays approach.

Other scripts in this demo show:
- server-rendered React component (with `next.js`)
- `preact` using the react compatibility library
- `react-native` deployment for iOS and android

## Internal State

The simplest state representation is an array of arrays.  To avoid having the
table component depend on the library, the column labels are precomputed.  The
state in this demo is shaped like the following object:

```js
{
  cols: [{ name: "A", key: 0 }, { name: "B", key: 1 }, { name: "C", key: 2 }],
  data: [
    [ "id",    "name", "value" ],
    [    1, "sheetjs",    7262 ]
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

## React Native

<img src="screen.png" width="400px"/>

Reproducing the full project is straightforward:

```bash
# see native.sh
react-native init SheetJS
cd SheetJS
npm i -S xlsx react react-native react-native-table-component react-native-fs
cp ../react-native.js index.ios.js
cp ../react-native.js index.android.js
react-native link
```

`react-native-table-component` draws the data table.  `react-native-fs` reads
and write files on devices.  The app will prompt before reading and after
writing data.  The printed location will be:

- android: path in the device filesystem
- iOS simulator: local path to file
- iOS device: a path accessible from iTunes App Documents view

`react-native-fs` supports `"ascii"` encoding for `readFile` and `writeFile`.
In practice, that encoding uses binary strings compatible with `"binary"` type:

```js
import { writeFile, readFile } from 'react-native-fs';

/* read a workbook */
readFile(file, 'ascii').then((res) => {
  const workbook = XLSX.read(res, {type:'binary'});
  /* DO SOMETHING WITH workbook HERE */
});

/* write a workbook */
const wbout = XLSX.write(wb, {type:'binary', bookType:"xlsx"});
writeFile(file, wbout, 'ascii').then((r)=>{/* :) */}).catch((e)=>{/* :( */});
```

## Other Demos

#### Preact

`preact-compat` is an easy-to-use compatibility layer that provides equivalents
for `React` and `ReactDOM`.  The `preact` demo uses the same JSX component code!
[The docs](https://npm.im/preact-compat#use-without-webpackbrowserify) explain
how to convert the in-browser React demo to Preact.

#### Server-Rendered React Components with Next.js

The demo uses the same component code as the in-browser version, but the build
step adds a small header that imports the library.  The import is not needed in
deployments that use script tags to include the library.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
