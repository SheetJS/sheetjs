## Writing Workbooks

For writing, the first step is to generate output data.  The helper functions
`write` and `writeFile` will produce the data in various formats suitable for
dissemination.  The second step is to actual share the data with the end point.
Assuming `workbook` is a workbook object:

<details>
  <summary><b>nodejs write a file</b> (click to show)</summary>

`writeFile` is only available in server environments. Browsers have no API for
writing arbitrary files given a path, so another strategy must be used.

```js
if(typeof require !== 'undefined') XLSX = require('xlsx');
/* output format determined by filename */
XLSX.writeFile(workbook, 'out.xlsb');
/* at this point, out.xlsb is a file that you can distribute */
```

</details>

<details>
  <summary><b>Browser add to web page</b> (click to show)</summary>

The `sheet_to_html` utility function generates HTML code that can be added to
any DOM element.

```js
var worksheet = workbook.Sheets[workbook.SheetNames[0]];
var container = document.getElementById('tableau');
container.innerHTML = XLSX.utils.sheet_to_html(worksheet);
```


</details>

<details>
  <summary><b>Browser save file</b> (click to show)</summary>

Note: browser generates binary blob and forces a "download" to client.  This
example uses [FileSaver](https://github.com/eligrey/FileSaver.js/):

```js
/* bookType can be any supported output type */
var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };

var wbout = XLSX.write(workbook,wopts);

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

/* the saveAs call downloads a file on the local machine */
saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "test.xlsx");
```

</details>

<details>
  <summary><b>Browser upload to server</b> (click to show)</summary>

A complete example using XHR is [included in the XHR demo](demos/xhr/), along
with examples for fetch and wrapper libraries.  This example assumes the server
can handle Base64-encoded files (see the demo for a basic nodejs server):

```js
/* in this example, send a base64 string to the server */
var wopts = { bookType:'xlsx', bookSST:false, type:'base64' };

var wbout = XLSX.write(workbook,wopts);

var req = new XMLHttpRequest();
req.open("POST", "/upload", true);
var formdata = new FormData();
formdata.append('file', 'test.xlsx'); // <-- server expects `file` to hold name
formdata.append('data', wbout); // <-- `data` holds the base64-encoded data
req.send(formdata);
```

</details>

### Writing Examples

- <http://sheetjs.com/demos/table.html> exporting an HTML table
- <http://sheetjs.com/demos/writexlsx.html> generates a simple file

