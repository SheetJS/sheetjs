# Adobe ExtendScript

ExtendScript adds some features to a limited form of ECMAScript version 3.  With
the included shim, the library can run within Photoshop and other Adobe apps!

The main file is `test.jsx`.  Target-specific files prepend target directives.
Copy the `test.jsx` file as well as the `shim.js` and `xlsx.core.min.js` files
to wherever you want the scripts to reside.

The demo shows opening a file and converting to an array of arrays:

```js
/* include library */
#include "shim.js"
#include "xlsx.core.min.js"

/* get data as binary string */
var filename = "sheetjs.xlsx";
var base = new File($.fileName);
var infile = File(base.path + "/" + filename);
infile.open("r");
infile.encoding = "binary";
var data = infile.read();

/* parse data */
var workbook = XLSX.read(data, {type:"binary"});

/* DO SOMETHING WITH workbook HERE */
```

NOTE: [We forked the minifier](https://www.npmjs.com/package/@sheetjs/uglify-js)
and included a patch for ExtendScript's switch statement semicolon issue.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
