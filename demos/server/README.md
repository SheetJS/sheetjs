# NodeJS Server Deployments

This library is 100% pure JS.  This is great for compatibility but tends to lock
up long-running processes.  In the web browser, Web Workers are used to offload
work from the main browser thread.  In NodeJS, there are other strategies.  This
demo shows a few different strategies applied to different server frameworks.

NOTE: these examples merely demonstrate the core concepts and do not include
appropriate error checking or other production-level features.


### Express Setup

The following commands are required in order to test the [Express](https://github.com/expressjs/express) demo:

```bash
npm install express printj xlsx express-formidable
node express.js
```

### Koa Setup

The following commands are required in order to test the [Koa](https://github.com/koajs/koa) demo:

```bash
npm install koa printj formidable xlsx
node koa.js
``` 

### Hapi Setup

**Note: Hapi demo as written only works with Hapi version 16 and below.**

The following commands are required in order to test the [Hapi](https://github.com/hapijs/hapi) demo:

```bash
npm install hapi@16.x printj tiny-worker xlsx
node hapi.js
``` 



### Node Buffer

The `read` and `write` functions can handle `Buffer` data with `type:"buffer"`.
For example, the `request` library returns data in a buffer:

```js
var XLSX = require('xlsx'), request = require('request');
request(url, {encoding: null}, function(err, res, data) {
	if(err || res.statusCode !== 200) return;

	/* data is a node Buffer that can be passed to XLSX.read */
	var workbook = XLSX.read(data, {type:'buffer'});

	/* DO SOMETHING WITH workbook HERE */
});
```

The `readFile` / `writeFile` functions wrap `fs.{read,write}FileSync`:

```js
/* equivalent to `var wb = XLSX.readFile("sheetjs.xlsx");` */
var buf = fs.readFileSync("sheetjs.xlsx");
var wb = XLSX.read(buf, {type:'buffer'});
```

### Responding to Form Uploads

Using `formidable`, files uploaded to forms are stored to temporary files that
can be read with `readFile`:

```js
/* within the server callback function(request, response) { */
var form = new formidable.IncomingForm();
form.parse(req, function(err, fields, files) {
	var f = files[Object.keys(files)[0]];
	var workbook = XLSX.readFile(f.path);
	/* DO SOMETHING WITH workbook HERE */
});
```

The `node.js` demo shows a plain HTTP server that accepts file uploads and
converts data to requested output format.

### Example servers

Each example server is expected to hold an array-of-arrays in memory.  They are
expected to handle:

- `POST /         ` accepts an encoded `file` and updates the internal storage
- `GET  /?t=<type>` returns the internal storage in the specified type
- `POST /?f=<name>` reads the local file and updates the internal storage
- `GET  /?f=<name>` writes the file to the specified name

Testing with cURL is straightforward:

```bash
# upload sheetjs.csv and update data
curl -X POST -F "data=@sheetjs.csv" http://localhost:7262/
# download data in SYLK format
curl -X GET http://localhost:7262/?t=slk
# read sheetjs.csv from the server directory
curl -X POST http://localhost:7262/?f=sheetjs.csv
# write sheetjs.xlsb in the XLSB format
curl -X GET http://localhost:7262/?f=sheetjs.xlsb
```


## Main-process logic with express

The most straightforward approach is to handle the data directly in HTTP event
handlers.  The `buffer` type for `XLSX.read` and `XLSX.write` work with `http`
module and with express directly.  The following snippet generates a workbook
based on an array of arrays and sends it to the client:

```js
function send_aoa_to_client(req, res, data, bookType) {
	/* generate workbook */
	var ws = XLSX.utils.aoa_to_sheet(data);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");

	/* generate buffer */
	var buf = XLSX.write(wb, {type:'buffer', bookType:bookType || "xlsx"});

	/* send to client */
	res.status(200).send(buf);
}
```


## fork with koa

`child_process.fork` provides a light-weight and customizable way to offload
work from the main server process.  This demo passes commands to a custom child
process and the child passes back buffers of data.

The main server script is `koa.js` and the worker script is `koasub.js`.  State
is maintained in the worker script.



## tiny-worker with hapi

`tiny-worker` provides a Web Worker-like interface.  Binary strings and simple
objects are readily passed across the Worker divide.

The main server script is `hapi.js` and the worker script is `worker.js`.  State
is maintained in the server script.

Note: due to an issue with hapi payload parsing, the route `POST /file` is used
to handle the case of reading from file, so the cURL test is:

```bash
# upload sheetjs.csv and update data
curl -X POST -F "data=@sheetjs.csv" http://localhost:7262/
# download data in SYLK format
curl -X GET http://localhost:7262/?t=slk
# read sheetjs.csv from the server directory
curl -X POST http://localhost:7262/file?f=sheetjs.csv
# write sheetjs.xlsb in the XLSB format
curl -X GET http://localhost:7262/?f=sheetjs.xlsb
```

[![Analytics](https://ga-beacon.appspot.com/UA-36810333-1/SheetJS/js-xlsx?pixel)](https://github.com/SheetJS/js-xlsx)
