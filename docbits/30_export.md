## Packaging and Releasing Data

### Writing Workbooks

**API**

_Generate spreadsheet bytes (file) from data_

```js
var data = XLSX.write(workbook, opts);
```

The `write` method attempts to package data from the workbook into a file in
memory.  By default, XLSX files are generated, but that can be controlled with
the `bookType` property of the `opts` argument.  Based on the `type` option,
the data can be stored as a "binary string", JS string, `Uint8Array` or Buffer.

The second `opts` argument is required.  ["Writing Options"](#writing-options)
covers the supported properties and behaviors.

_Generate and attempt to save file_

```js
XLSX.writeFile(workbook, filename, opts);
```

The `writeFile` method packages the data and attempts to save the new file.  The
export file format is determined by the extension of `filename` (`SheetJS.xlsx`
signals XLSX export, `SheetJS.xlsb` signals XLSB export, etc).

The `writeFile` method uses platform-specific APIs to initiate the file save. In
NodeJS, `fs.readFileSync` can create a file.  In the web browser, a download is
attempted using the HTML5 `download` attribute, with fallbacks for IE.

_Generate and attempt to save an XLSX file_

```js
XLSX.writeFileXLSX(workbook, filename, opts);
```

The `writeFile` method embeds a number of different export functions.  This is
great for developer experience but not amenable to tree shaking using the
current developer tools.  When only XLSX exports are needed, this method avoids
referencing the other export functions.

The second `opts` argument is optional.  ["Writing Options"](#writing-options)
covers the supported properties and behaviors.

**Examples**

<details>
  <summary><b>Local file in a NodeJS server</b> (click to show)</summary>

`writeFile` uses `fs.writeFileSync` in server environments:

```js
var XLSX = require("xlsx");

/* output format determined by filename */
XLSX.writeFile(workbook, "out.xlsb");
```

For Node ESM, the `writeFile` helper is not enabled. Instead, `fs.writeFileSync`
should be used to write the file data to a `Buffer` for use with `XLSX.write`:

```js
import { writeFileSync } from "fs";
import { write } from "xlsx/xlsx.mjs";

const buf = write(workbook, {type: "buffer", bookType: "xlsb"});
/* buf is a Buffer */
const workbook = writeFileSync("out.xlsb", buf);
```

</details>

<details>
  <summary><b>Local file in a Deno application</b> (click to show)</summary>

`writeFile` uses `Deno.writeFileSync` under the hood:

```js
// @deno-types="https://deno.land/x/sheetjs/types/index.d.ts"
import * as XLSX from 'https://deno.land/x/sheetjs/xlsx.mjs'

XLSX.writeFile(workbook, "test.xlsx");
```

Applications writing files must be invoked with the `--allow-write` flag.  The
[`deno` demo](demos/deno/) has more examples

</details>

<details>
  <summary><b>Local file in a PhotoShop or InDesign plugin</b> (click to show)</summary>

`writeFile` wraps the `File` logic in Photoshop and other ExtendScript targets.
The specified path should be an absolute path:

```js
#include "xlsx.extendscript.js"

/* output format determined by filename */
XLSX.writeFile(workbook, "out.xlsx");
/* at this point, out.xlsx is a file that you can distribute */
```

The [`extendscript` demo](demos/extendscript/) includes a more complex example.

</details>

<details>
  <summary><b>Download a file in the browser to the user machine</b> (click to show)</summary>

`XLSX.writeFile` wraps a few techniques for triggering a file save:

- `URL` browser API creates an object URL for the file, which the library uses
  by creating a link and forcing a click. It is supported in modern browsers.
- `msSaveBlob` is an IE10+ API for triggering a file save.
- `IE_FileSave` uses VBScript and ActiveX to write a file in IE6+ for Windows
  XP and Windows 7.  The shim must be included in the containing HTML page.

There is no standard way to determine if the actual file has been downloaded.

```js
/* output format determined by filename */
XLSX.writeFile(workbook, "out.xlsb");
/* at this point, out.xlsb will have been downloaded */
```

</details>

<details>
  <summary><b>Download a file in legacy browsers</b> (click to show)</summary>

`XLSX.writeFile` techniques work for most modern browsers as well as older IE.
For much older browsers, there are workarounds implemented by wrapper libraries.

[`FileSaver.js`](https://github.com/eligrey/FileSaver.js/) implements `saveAs`.
Note: `XLSX.writeFile` will automatically call `saveAs` if available.

```js
/* bookType can be any supported output type */
var wopts = { bookType:"xlsx", bookSST:false, type:"array" };

var wbout = XLSX.write(workbook,wopts);

/* the saveAs call downloads a file on the local machine */
saveAs(new Blob([wbout],{type:"application/octet-stream"}), "test.xlsx");
```

[`Downloadify`](https://github.com/dcneiner/downloadify) uses a Flash SWF button
to generate local files, suitable for environments where ActiveX is unavailable:

```js
Downloadify.create(id,{
  /* other options are required! read the downloadify docs for more info */
  filename: "test.xlsx",
  data: function() { return XLSX.write(wb, {bookType:"xlsx", type:"base64"}); },
  append: false,
  dataType: "base64"
});
```

The [`oldie` demo](demos/oldie/) shows an IE-compatible fallback scenario.

</details>

<details>
  <summary><b>Browser upload file (ajax)</b> (click to show)</summary>

A complete example using XHR is [included in the XHR demo](demos/xhr/), along
with examples for fetch and wrapper libraries.  This example assumes the server
can handle Base64-encoded files (see the demo for a basic nodejs server):

```js
/* in this example, send a base64 string to the server */
var wopts = { bookType:"xlsx", bookSST:false, type:"base64" };

var wbout = XLSX.write(workbook,wopts);

var req = new XMLHttpRequest();
req.open("POST", "/upload", true);
var formdata = new FormData();
formdata.append("file", "test.xlsx"); // <-- server expects `file` to hold name
formdata.append("data", wbout); // <-- `data` holds the base64-encoded data
req.send(formdata);
```

</details>

<details>
  <summary><b>PhantomJS (Headless Webkit) File Generation</b> (click to show)</summary>

The [`headless` demo](demos/headless/) includes a complete demo to convert HTML
files to XLSB workbooks using [PhantomJS](https://phantomjs.org/). PhantomJS
`fs.write` supports writing files from the main process but has a different
interface from the NodeJS `fs` module:

```js
var XLSX = require('xlsx');
var fs = require('fs');

/* generate a binary string */
var bin = XLSX.write(workbook, { type:"binary", bookType: "xlsx" });
/* write to file */
fs.write("test.xlsx", bin, "wb");
```

Note: The section ["Processing HTML Tables"](#processing-html-tables) shows how
to generate a workbook from HTML tables in a page in "Headless WebKit".

</details>



The [included demos](demos/) cover mobile apps and other special deployments.

### Writing Examples

- <http://sheetjs.com/demos/table.html> exporting an HTML table
- <http://sheetjs.com/demos/writexlsx.html> generates a simple file

### Streaming Write

The streaming write functions are available in the `XLSX.stream` object.  They
take the same arguments as the normal write functions but return a NodeJS
Readable Stream.

- `XLSX.stream.to_csv` is the streaming version of `XLSX.utils.sheet_to_csv`.
- `XLSX.stream.to_html` is the streaming version of `XLSX.utils.sheet_to_html`.
- `XLSX.stream.to_json` is the streaming version of `XLSX.utils.sheet_to_json`.

<details>
  <summary><b>nodejs convert to CSV and write file</b> (click to show)</summary>

```js
var output_file_name = "out.csv";
var stream = XLSX.stream.to_csv(worksheet);
stream.pipe(fs.createWriteStream(output_file_name));
```

</details>

<details>
  <summary><b>nodejs write JSON stream to screen</b> (click to show)</summary>

```js
/* to_json returns an object-mode stream */
var stream = XLSX.stream.to_json(worksheet, {raw:true});

/* the following stream converts JS objects to text via JSON.stringify */
var conv = new Transform({writableObjectMode:true});
conv._transform = function(obj, e, cb){ cb(null, JSON.stringify(obj) + "\n"); };

stream.pipe(conv); conv.pipe(process.stdout);
```

</details>

<details>
  <summary><b>Exporting NUMBERS files</b> (click to show)</summary>

The NUMBERS writer requires a fairly large base.  The supplementary `xlsx.zahl`
scripts provide support.  `xlsx.zahl.js` is designed for standalone and NodeJS
use, while `xlsx.zahl.mjs` is suitable for ESM.

_Browser_

```html
<meta charset="utf8">
<script src="xlsx.full.min.js"></script>
<script src="xlsx.zahl.js"></script>
<script>
var wb = XLSX.utils.book_new(); var ws = XLSX.utils.aoa_to_sheet([
  ["SheetJS", "<3","விரிதாள்"],
  [72,,"Arbeitsblätter"],
  [,62,"数据"],
  [true,false,],
]); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
XLSX.writeFile(wb, "textport.numbers", {numbers: XLSX_ZAHL, compression: true});
</script>
```

_Node_

```js
var XLSX = require("./xlsx.flow");
var XLSX_ZAHL = require("./dist/xlsx.zahl");
var wb = XLSX.utils.book_new(); var ws = XLSX.utils.aoa_to_sheet([
  ["SheetJS", "<3","விரிதாள்"],
  [72,,"Arbeitsblätter"],
  [,62,"数据"],
  [true,false,],
]); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
XLSX.writeFile(wb, "textport.numbers", {numbers: XLSX_ZAHL, compression: true});
```

_Deno_

```ts
import * as XLSX from './xlsx.mjs';
import XLSX_ZAHL from './dist/xlsx.zahl.mjs';

var wb = XLSX.utils.book_new(); var ws = XLSX.utils.aoa_to_sheet([
  ["SheetJS", "<3","விரிதாள்"],
  [72,,"Arbeitsblätter"],
  [,62,"数据"],
  [true,false,],
]); XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
XLSX.writeFile(wb, "textports.numbers", {numbers: XLSX_ZAHL, compression: true});
```

</details>

<https://github.com/sheetjs/sheetaki> pipes write streams to nodejs response.

