## Acquiring and Extracting Data

### Parsing Workbooks

**API**

_Extract data from spreadsheet bytes_

```js
var workbook = XLSX.read(data, opts);
```

The `read` method can extract data from spreadsheet bytes stored in a JS string,
"binary string", NodeJS buffer or typed array (`Uint8Array` or `ArrayBuffer`).


_Read spreadsheet bytes from a local file and extract data_

```js
var workbook = XLSX.readFile(filename, opts);
```

The `readFile` method attempts to read a spreadsheet file at the supplied path.
Browsers generally do not allow reading files in this way (it is deemed a
security risk), and attempts to read files in this way will throw an error.

The second `opts` argument is optional. ["Parsing Options"](#parsing-options)
covers the supported properties and behaviors.

**Examples**

Here are a few common scenarios (click on each subtitle to see the code):

<details>
  <summary><b>Local file in a NodeJS server</b> (click to show)</summary>

`readFile` uses `fs.readFileSync` under the hood:

```js
var XLSX = require("xlsx");

var workbook = XLSX.readFile("test.xlsx");
```

For Node ESM, the `readFile` helper is not enabled. Instead, `fs.readFileSync`
should be used to read the file data as a `Buffer` for use with `XLSX.read`:

```js
import { readFileSync } from "fs";
import { read } from "xlsx/xlsx.mjs";

const buf = readFileSync("test.xlsx");
/* buf is a Buffer */
const workbook = read(buf);
```

</details>

<details>
  <summary><b>Local file in a Deno application</b> (click to show)</summary>

`readFile` uses `Deno.readFileSync` under the hood:

```js
// @deno-types="https://deno.land/x/sheetjs/types/index.d.ts"
import * as XLSX from 'https://deno.land/x/sheetjs/xlsx.mjs'

const workbook = XLSX.readFile("test.xlsx");
```

Applications reading files must be invoked with the `--allow-read` flag.  The
[`deno` demo](demos/deno/) has more examples

</details>

<details>
  <summary><b>User-submitted file in a web page ("Drag-and-Drop")</b> (click to show)</summary>

For modern websites targeting Chrome 76+, `File#arrayBuffer` is recommended:

```js
// XLSX is a global from the standalone script

async function handleDropAsync(e) {
  e.stopPropagation(); e.preventDefault();
  const f = e.dataTransfer.files[0];
  /* f is a File */
  const data = await f.arrayBuffer();
  /* data is an ArrayBuffer */
  const workbook = XLSX.read(data);

  /* DO SOMETHING WITH workbook HERE */
}
drop_dom_element.addEventListener("drop", handleDropAsync, false);
```

For maximal compatibility, the `FileReader` API should be used:

```js
function handleDrop(e) {
  e.stopPropagation(); e.preventDefault();
  var f = e.dataTransfer.files[0];
  /* f is a File */
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    /* reader.readAsArrayBuffer(file) -> data will be an ArrayBuffer */
    var workbook = XLSX.read(data);

    /* DO SOMETHING WITH workbook HERE */
  };
  reader.readAsArrayBuffer(f);
}
drop_dom_element.addEventListener("drop", handleDrop, false);
```

<https://oss.sheetjs.com/sheetjs/> demonstrates the FileReader technique.

</details>

<details>
  <summary><b>User-submitted file with an HTML INPUT element</b> (click to show)</summary>

Starting with an HTML INPUT element with `type="file"`:

```html
<input type="file" id="input_dom_element">
```

For modern websites targeting Chrome 76+, `Blob#arrayBuffer` is recommended:

```js
// XLSX is a global from the standalone script

async function handleFileAsync(e) {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  /* data is an ArrayBuffer */
  const workbook = XLSX.read(data);

  /* DO SOMETHING WITH workbook HERE */
}
input_dom_element.addEventListener("change", handleFileAsync, false);
```

For broader support (including IE10+), the `FileReader` approach is recommended:

```js
function handleFile(e) {
  var file = e.target.files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    /* reader.readAsArrayBuffer(file) -> data will be an ArrayBuffer */
    var workbook = XLSX.read(e.target.result);

    /* DO SOMETHING WITH workbook HERE */
  };
  reader.readAsArrayBuffer(file);
}
input_dom_element.addEventListener("change", handleFile, false);
```

The [`oldie` demo](demos/oldie/) shows an IE-compatible fallback scenario.

</details>

<details>
  <summary><b>Fetching a file in the web browser ("Ajax")</b> (click to show)</summary>

For modern websites targeting Chrome 42+, `fetch` is recommended:

```js
// XLSX is a global from the standalone script

(async() => {
  const url = "http://oss.sheetjs.com/test_files/formula_stress_test.xlsx";
  const data = await (await fetch(url)).arrayBuffer();
  /* data is an ArrayBuffer */
  const workbook = XLSX.read(data);

  /* DO SOMETHING WITH workbook HERE */
})();
```

For broader support, the `XMLHttpRequest` approach is recommended:

```js
var url = "http://oss.sheetjs.com/test_files/formula_stress_test.xlsx";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
  var workbook = XLSX.read(req.response);

  /* DO SOMETHING WITH workbook HERE */
};

req.send();
```

The [`xhr` demo](demos/xhr/) includes a longer discussion and more examples.

<http://oss.sheetjs.com/sheetjs/ajax.html> shows fallback approaches for IE6+.

</details>

<details>
  <summary><b>Local file in a PhotoShop or InDesign plugin</b> (click to show)</summary>

`readFile` wraps the `File` logic in Photoshop and other ExtendScript targets.
The specified path should be an absolute path:

```js
#include "xlsx.extendscript.js"

/* Read test.xlsx from the Documents folder */
var workbook = XLSX.readFile(Folder.myDocuments + "/test.xlsx");
```

The [`extendscript` demo](demos/extendscript/) includes a more complex example.

</details>

<details>
  <summary><b>Local file in an Electron app</b> (click to show)</summary>

`readFile` can be used in the renderer process:

```js
/* From the renderer process */
var XLSX = require("xlsx");

var workbook = XLSX.readFile(path);
```

Electron APIs have changed over time.  The [`electron` demo](demos/electron/)
shows a complete example and details the required version-specific settings.

</details>

<details>
  <summary><b>Local file in a mobile app with React Native</b> (click to show)</summary>

The [`react` demo](demos/react) includes a sample React Native app.

Since React Native does not provide a way to read files from the filesystem, a
third-party library must be used.  The following libraries have been tested:

- [`react-native-file-access`](https://npm.im/react-native-file-access)

The `base64` encoding returns strings compatible with the `base64` type:

```js
import XLSX from "xlsx";
import { FileSystem } from "react-native-file-access";

const b64 = await FileSystem.readFile(path, "base64");
/* b64 is a base64 string */
const workbook = XLSX.read(b64, {type: "base64"});
```

- [`react-native-fs`](https://npm.im/react-native-fs)

The `ascii` encoding returns binary strings compatible with the `binary` type:

```js
import XLSX from "xlsx";
import { readFile } from "react-native-fs";

const bstr = await readFile(path, "ascii");
/* bstr is a binary string */
const workbook = XLSX.read(bstr, {type: "binary"});
```

</details>

<details>
  <summary><b>NodeJS Server File Uploads</b> (click to show)</summary>

`read` can accept a NodeJS buffer.  `readFile` can read files generated by a
HTTP POST request body parser like [`formidable`](https://npm.im/formidable):

```js
const XLSX = require("xlsx");
const http = require("http");
const formidable = require("formidable");

const server = http.createServer((req, res) => {
  const form = new formidable.IncomingForm();
  form.parse(req, (err, fields, files) => {
    /* grab the first file */
    const f = Object.entries(files)[0][1];
    const path = f.filepath;
    const workbook = XLSX.readFile(path);

    /* DO SOMETHING WITH workbook HERE */
  });
}).listen(process.env.PORT || 7262);
```

The [`server` demo](demos/server) has more advanced examples.

</details>

<details>
  <summary><b>Download files in a NodeJS process</b> (click to show)</summary>

Node 17.5 and 18.0 have native support for fetch:

```js
const XLSX = require("xlsx");

const data = await (await fetch(url)).arrayBuffer();
/* data is an ArrayBuffer */
const workbook = XLSX.read(data);
```

For broader compatibility, third-party modules are recommended.

[`request`](https://npm.im/request) requires a `null` encoding to yield Buffers:

```js
var XLSX = require("xlsx");
var request = require("request");

request({url: url, encoding: null}, function(err, resp, body) {
  var workbook = XLSX.read(body);

  /* DO SOMETHING WITH workbook HERE */
});
```

[`axios`](https://npm.im/axios) works the same way in browser and in NodeJS:

```js
const XLSX = require("xlsx");
const axios = require("axios");

(async() => {
  const res = await axios.get(url, {responseType: "arraybuffer"});
  /* res.data is a Buffer */
  const workbook = XLSX.read(res.data);

  /* DO SOMETHING WITH workbook HERE */
})();
```

</details>

<details>
  <summary><b>Download files in an Electron app</b> (click to show)</summary>

The `net` module in the main process can make HTTP/HTTPS requests to external
resources.  Responses should be manually concatenated using `Buffer.concat`:

```js
const XLSX = require("xlsx");
const { net } = require("electron");

const req = net.request(url);
req.on("response", (res) => {
  const bufs = []; // this array will collect all of the buffers
  res.on("data", (chunk) => { bufs.push(chunk); });
  res.on("end", () => {
    const workbook = XLSX.read(Buffer.concat(bufs));

    /* DO SOMETHING WITH workbook HERE */
  });
});
req.end();
```

</details>

<details>
  <summary><b>Readable Streams in NodeJS</b> (click to show)</summary>

When dealing with Readable Streams, the easiest approach is to buffer the stream
and process the whole thing at the end:

```js
var fs = require("fs");
var XLSX = require("xlsx");

function process_RS(stream, cb) {
  var buffers = [];
  stream.on("data", function(data) { buffers.push(data); });
  stream.on("end", function() {
    var buffer = Buffer.concat(buffers);
    var workbook = XLSX.read(buffer, {type:"buffer"});

    /* DO SOMETHING WITH workbook IN THE CALLBACK */
    cb(workbook);
  });
}
```

</details>

<details>
  <summary><b>ReadableStream in the browser</b> (click to show)</summary>

When dealing with `ReadableStream`, the easiest approach is to buffer the stream
and process the whole thing at the end:

```js
// XLSX is a global from the standalone script

async function process_RS(stream) {
  /* collect data */
  const buffers = [];
  const reader = stream.getReader();
  for(;;) {
    const res = await reader.read();
    if(res.value) buffers.push(res.value);
    if(res.done) break;
  }

  /* concat */
  const out = new Uint8Array(buffers.reduce((acc, v) => acc + v.length, 0));

  let off = 0;
  for(const u8 of arr) {
    out.set(u8, off);
    off += u8.length;
  }

  return out;
}

const data = await process_RS(stream);
/* data is Uint8Array */
const workbook = XLSX.read(data);
```

</details>

More detailed examples are covered in the [included demos](demos/)

