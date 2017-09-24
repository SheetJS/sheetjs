## Parsing Workbooks

For parsing, the first step is to read the file.  This involves acquiring the
data and feeding it into the library.  Here are a few common scenarios:

<details>
  <summary><b>nodejs read a file</b> (click to show)</summary>

```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
var workbook = XLSX.readFile('test.xlsx');
/* DO SOMETHING WITH workbook HERE */
```

</details>

<details>
  <summary><b>Browser read TABLE element from page</b> (click to show)</summary>

```js
var worksheet = XLSX.utils.table_to_book(document.getElementById('tableau'));
/* DO SOMETHING WITH workbook HERE */
```

</details>

<details>
  <summary><b>Browser download file (ajax)</b> (click to show)</summary>

Note: for a more complete example that works in older browsers, check the demo
at <http://oss.sheetjs.com/js-xlsx/ajax.html>).  The <demos/xhr/> directory also
includes more examples with `XMLHttpRequest` and `fetch`.

```js
var url = "http://oss.sheetjs.com/test_files/formula_stress_test.xlsx";

/* set up async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
  var data = new Uint8Array(req.response);
  var workbook = XLSX.read(data, {type:"array"});

  /* DO SOMETHING WITH workbook HERE */
}

req.send();
```

</details>

<details>
  <summary><b>Browser drag-and-drop</b> (click to show)</summary>

Drag-and-drop uses the HTML5 `FileReader` API, loading the data with
`readAsBinaryString` or `readAsArrayBuffer`.  Since not all browsers support the
full `FileReader` API, dynamic feature tests are highly recommended.

```js
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleDrop(e) {
  e.stopPropagation(); e.preventDefault();
  var files = e.dataTransfer.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
drop_dom_element.addEventListener('drop', handleDrop, false);
```

</details>

<details>
  <summary><b>Browser file upload form element</b> (click to show)</summary>

Data from file input elements can be processed using the same `FileReader` API
as in the drag-and-drop example:

```js
var rABS = true; // true: readAsBinaryString ; false: readAsArrayBuffer
function handleFile(e) {
  var files = e.target.files, f = files[0];
  var reader = new FileReader();
  reader.onload = function(e) {
    var data = e.target.result;
    if(!rABS) data = new Uint8Array(data);
    var workbook = XLSX.read(data, {type: rABS ? 'binary' : 'array'});

    /* DO SOMETHING WITH workbook HERE */
  };
  if(rABS) reader.readAsBinaryString(f); else reader.readAsArrayBuffer(f);
}
input_dom_element.addEventListener('change', handleFile, false);
```

</details>


### Parsing Examples

- <http://oss.sheetjs.com/js-xlsx/> HTML5 File API / Base64 Text / Web Workers

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

- <http://oss.sheetjs.com/js-xlsx/ajax.html> XMLHttpRequest

