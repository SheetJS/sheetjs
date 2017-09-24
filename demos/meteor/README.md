# Meteor

This library is universal: outside of environment-specific features (parsing DOM
tables in the browser, streaming write in nodejs), the core is ES3/ES5 and can
be used in any reasonably compliant JS implementation.  It should play nice with
meteor out of the box.

Using the npm module, the library can be imported from client or server side:

```js
import XLSX from 'xlsx'
```

All of the functions and utilities are available in both realms. Since the core
data representations are simple JS objects, the workbook object can be passed on
the wire, enabling hybrid workflows where the server processes data and client
finishes the work.


## This demonstration

Note: the obvious extremes of pure-client code and pure-server code are covered
in other demos.

### Reading Data

The parse demo:
- accepts files from the client side
- sends binary string to server
- processes data on server side
- sends workbook object to client
- renders HTML and adds to a DOM element

The logic from within the `FileReader` is split as follows:

```js
// CLIENT SIDE
const bstr = e.target.result;
// SERVER SIDE
const wb = XLSX.read(bstr, { type: 'binary' });
// CLIENT SIDE
const ws = wb.Sheets[wb.SheetNames[0]];
const html = XLSX.utils.sheet_to_html(ws, { editable: true });
document.getElementById('out').innerHTML = html;
```

### Writing Data

The write demo:
- grabs HTML from the client side
- sends HTML string to server
- processes data on server side
- sends workbook object to client
- generates file on client side and triggers a download

The logic from within the `click` event is split as follows:

```js
// CLIENT SIDE
const html = document.getElementById('out').innerHTML;
// SERVER SIDE
const wb = XLSX.read(html, { type: 'binary' });
// CLIENT SIDE
const o = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
saveAs(new Blob([s2ab(o)], {type:'application/octet-stream'}), 'sheetjs.xlsx');
```

This demo uses the FileSaver library for writing files, installed through the
[`pfafman:filesaver` wrapper](https://atmospherejs.com/pfafman/filesaver).


## Setup

This tree does not include the `.meteor` structure.  Rebuild the project with:

```bash
meteor create .
npm install babel-runtime meteor-node-stubs xlsx
meteor add pfafman:filesaver
meteor
```


## Environment-Specific Features

File-related operations like `XLSX.readFile` and `XLSX.writeFile` will not be
available in client-side code. If you need to read a local file from the client,
use a file input or drag-and-drop.

Browser-specific operations like `XLSX.utils.table_to_book` are limited to
client side code. You should never have to read from DOM elements on the server
side, but you can use a third-party virtual DOM to provide the required API.

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
