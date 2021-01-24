# Adobe ExtendScript

ExtendScript adds some features to a limited form of ECMAScript version 3.  With
the included shim, the library can run within Photoshop and other Adobe apps!

The main file is `test.jsx`.  Target-specific files prepend target directives.
Copy the `test.jsx` file as well as the `xlsx.extendscript.js` library script 
to wherever you want the scripts to reside.


## ExtendScript Quirks

There are numerous quirks in ExtendScript code parsing, especially related to
Boolean and bit operations.  Most JS tooling will generate code that is not
compatible with ExtendScript.  It is highly recommended to `#include` the `dist`
file directly and avoid trying to minify or pack as part of a larger project.


## File I/O

Using the `"binary"` encoding, file operations will work with binary strings
that play nice with the `"binary"` type of this library.  The `readFile` and
`writeFile` library functions wrap the File logic:

```js
/* Read file from disk */
var workbook = XLSX.readFile(filename);

/* Write file to disk */
XLSX.writeFile(workbook, filename);
```

The `readFile` and `writeFile` functions use `"binary"` encoding under the hood:


```js
/* Read file from disk without using readFile */
var infile = File(filename);
infile.open("r");
infile.encoding = "binary";
var data = infile.read();
var workbook = XLSX.read(data, {type:"binary"});
infile.close();

/* Write file to disk without using writeFile */
var outFile = File(filename);
outFile.open("w");
outFile.encoding = "binary";
outFile.write(workbook);
outFile.close();
```


## Demo

The demo shows:

- loading the library in ExtendScript using `#include`:

```js
#include "xlsx.extendscript.js"
```

- opening a file with `XLSX.readFile`:

```js
var workbook = XLSX.readFile("sheetjs.xlsx");
```

- converting a worksheet to an array of arrays:

```js
var first_sheet_name = workbook.SheetNames[0];
var first_worksheet = workbook.Sheets[first_sheet_name];
var data = XLSX.utils.sheet_to_json(first_worksheet, {header:1});

alert(data);
```

- writing a new workbook file:

```js
XLSX.writeFile(workbook, "sheetjs.slk");
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
