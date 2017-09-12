# XMLHttpRequest and fetch

`XMLHttpRequest` and `fetch` browser APIs enable binary data transfer between
web browser clients and web servers.  Since this library works in web browsers,
server conversion work can be offloaded to the client!  This demo shows a few
common scenarios involving browser APIs and popular wrapper libraries.

## Sample Server

The `server.js` nodejs server serves static files on `GET` request.  On a `POST`
request to `/upload`, the server processes the body and looks for the `file` and
`data` fields.  It will write the Base64-decoded data from `data` to the file
name specified in `file`.

To start the demo, run `npm start` and navigate to <http://localhost:7262/>

## XMLHttpRequest

For downloading data, the `arraybuffer` response type generates an `ArrayBuffer`
that can be viewed as an `Uint8Array` and fed to `XLSX.read` using `array` type:

```js
/* set up an async GET request */
var req = new XMLHttpRequest();
req.open("GET", url, true);
req.responseType = "arraybuffer";

req.onload = function(e) {
	/* parse the data when it is received */
	var data = new Uint8Array(oReq.response);
	var workbook = XLSX.read(data, {type:"array"});
	/* DO SOMETHING WITH workbook HERE */
};
req.send();
```

For uploading data, this demo populates a `FormData` object with string data
generated with the `base64` output type:

```js
/* generate XLSX as base64 string */
var b64 = XLSX.write(workbook, {bookType:'xlsx', type:'base64'});

/* build FormData with the generated file */
var fd = new FormData();
fd.append('data', b64);

/* send data */
var req = new XMLHttpRequest();
req.open("POST", "/upload", true);
req.send(formdata);
```

axios and superagent patterns are similar to the XMLHttpRequest pattern but
involve much less boilerplate:

```js
/* set up an async GET request with axios */
axios(url, {responseType:'arraybuffer'}).then(function(res) {
	/* parse the data when it is received */
	var data = new Uint8Array(res.data);
	var workbook = XLSX.read(data, {type:"array"});

	/* DO SOMETHING WITH workbook HERE */
});

/* set up an async GET request with superagent */
superagent.get(url).responseType('arraybuffer').end(function(err, res) {
	/* parse the data when it is received */
	var data = new Uint8Array(res.body);
	var workbook = XLSX.read(data, {type:"array"});

	/* DO SOMETHING WITH workbook HERE */
});

```

## fetch

For downloading data, `response.blob()` resolves to a `Blob` object that can be
converted to `ArrayBuffer` using a `FileReader`:

```js
fetch(url).then(function(res) {
	/* get the data as a Blob */
	if(!res.ok) throw new Error("fetch failed");
	return res.blob();
}).then(function(blob) {
	/* configure a FileReader to process the blob */
	var reader = new FileReader();
	reader.addEventListener("loadend", function() {
		/* parse the data when it is received */
		var data = new Uint8Array(this.result);
		var workbook = XLSX.read(data, {type:"array"});

		/* DO SOMETHING WITH workbook HERE */
	});
	reader.readAsArrayBuffer(blob);
});
```

