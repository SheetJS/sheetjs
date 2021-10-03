## Parsing Workbooks

For parsing, the first step is to read the file.  This involves acquiring the
data and feeding it into the library.  Here are a few common scenarios:

<details>
  <summary><b>nodejs read a file</b> (click to show)</summary>

`readFile` is only available in server environments. Browsers have no API for
reading arbitrary files given a path, so another strategy must be used.

```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

</details>

<details>
  <summary><b>Photoshop ExtendScript read a file</b> (click to show)</summary>

`readFile` wraps the `File` logic in Photoshop and other ExtendScript targets.
The specified path should be an absolute path:

```js
#include "xlsx.extendscript.js"
/* Read test.xlsx from the Documents folder */
var workbook = XLSX.readFile(Folder.myDocuments + '/' + 'test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

The [`extendscript` demo](demos/extendscript/) includes a more complex example.

</details>

<details>
  <summary><b>Browser read TABLE element from page</b> (click to show)</summary>

The `table_to_book` and `table_to_sheet` utility functions take a DOM TABLE
element and iterate through the child nodes.

```js
var workbook = XLSX.utils.table_to_book(document.getElementById('tableau'));
/* DO SOMETHING WITH workbook HERE */
```

Multiple tables on a web page can be converted to individual worksheets:

```js
/* create new workbook */
var workbook = XLSX.utils.book_new();

/* convert table 'table1' to worksheet named "Sheet1" */
var ws1 = XLSX.utils.table_to_sheet(document.getElementById('table1'));
XLSX.utils.book_append_sheet(workbook, ws1, "Sheet1");

/* convert table 'table2' to worksheet named "Sheet2" */
var ws2 = XLSX.utils.table_to_sheet(document.getElementById('table2'));
XLSX.utils.book_append_sheet(workbook, ws2, "Sheet2");

/* workbook now has 2 worksheets */
```

Alternatively, the HTML code can be extracted and parsed:

```js
var htmlstr = document.getElementById('tableau').outerHTML;
var workbook = XLSX.read(htmlstr, {type:'string'});
```

</details>

<details>
  <summary><b>Browser download file (ajax)</b> (click to show)</summary>

Note: for a more complete example that works in older browsers, check the demo
at <http://oss.sheetjs.com/sheetjs/ajax.html>.  The [`xhr` demo](demos/xhr/)
includes more examples with `XMLHttpRequest` and `fetch`.

```js
var url = "http://oss.sheetjs.com/test_files/formula_stress_test.xlsx";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
  var workbook = XLSX.read(req.response);

  /* DO SOMETHING WITH workbook HERE */
}

req.send();
```

</details>

<details>
  <summary><b>Browser drag-and-drop</b> (click to show)</summary>

For modern browsers, `Blob#arrayBuffer` can read data from files:

```js
async function handleDropAsync(e) {
  e.stopPropagation(); e.preventDefault();
  const f = evt.dataTransfer.files[0];
  const data = await f.arrayBuffer();
  const workbook = XLSX.read(data);

  /* DO SOMETHING WITH workbook HERE */
}
drop_dom_element.addEventListener('drop', handleDropAsync, false);
```

For maximal compatibility, the `FileReader` API should be used:

```js
function handleDrop(e) {
  e.stopPropagation(); e.preventDefault();
  var f = e.dataTransfer.files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var workbook = XLSX.read(e.target.result);

    /* DO SOMETHING WITH workbook HERE */
  };
  reader.readAsArrayBuffer(f);
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

</details>

<details>
  <summary><b>Browser file upload form element</b> (click to show)</summary>

Data from file input elements can be processed using the same APIs as in the
drag-and-drop example.

Using `Blob#arrayBuffer`:

```js
async function handleFileAsync(e) {
  const file = evt.target.files[0];
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data);

  /* DO SOMETHING WITH workbook HERE */
}
input_dom_element.addEventListener('change', handleFileAsync, false);
```

Using `FileReader`:

```js
function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var workbook = XLSX.read(e.target.result);

    /* DO SOMETHING WITH workbook HERE */
  };
  reader.readAsArrayBuffer(f);
}
input_dom_element.addEventListener('change', handleFile, false);
```

The [`oldie` demo](demos/oldie/) shows an IE-compatible fallback scenario.

</details>

More specialized cases, including mobile app file processing, are covered in the
[included demos](demos/)

### Parsing Examples

- <https://oss.sheetjs.com/sheetjs/> HTML5 File API / Base64 Text / Web Workers

Note that older versions of IE do not support HTML5 File API, so the Base64 mode
is used for testing.

<details>
  <summary><b>Get Base64 encoding on OSX / Windows</b> (click to show)</summary>

On OSX you can get the Base64 encoding with:

```bash
$ <target_file base64 | pbcopy
```

On Windows XP and up you can get the Base64 encoding using `certutil`:

```cmd
> certutil -encode target_file target_file.b64
```

(note: You have to open the file and remove the header and footer lines)

</details>

- <http://oss.sheetjs.com/sheetjs/ajax.html> XMLHttpRequest

