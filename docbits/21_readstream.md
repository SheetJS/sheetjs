### Streaming Read

<details>
	<summary><b>Why is there no Streaming Read API?</b> (click to show)</summary>

The most common and interesting formats (XLS, XLSX/M, XLSB, ODS) are ultimately
ZIP or CFB containers of files.  Neither format puts the directory structure at
the beginning of the file: ZIP files place the Central Directory records at the
end of the logical file, while CFB files can place the FAT structure anywhere in
the file! As a result, to properly handle these formats, a streaming function
would have to buffer the entire file before commencing.  That belies the
expectations of streaming, so we do not provide any streaming read API.

</details>

When dealing with Readable Streams, the easiest approach is to buffer the stream
and process the whole thing at the end.  This can be done with a temporary file
or by explicitly concatenating the stream:

<details>
	<summary><b>Explicitly concatenating streams</b> (click to show)</summary>

```js
var fs = require('fs');
var XLSX = require('xlsx');
function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/{
	var buffers = [];
	stream.on('data', function(data) { buffers.push(data); });
	stream.on('end', function() {
		var buffer = Buffer.concat(buffers);
		var workbook = XLSX.read(buffer, {type:"buffer"});

		/* DO SOMETHING WITH workbook IN THE CALLBACK */
		cb(workbook);
	});
}
```

More robust solutions are available using modules like `concat-stream`.

</details>

<details>
	<summary><b>Writing to filesystem first</b> (click to show)</summary>

This example uses [`tempfile`](https://npm.im/tempfile) for filenames:

```js
var fs = require('fs'), tempfile = require('tempfile');
var XLSX = require('xlsx');
function process_RS(stream/*:ReadStream*/, cb/*:(wb:Workbook)=>void*/)/*:void*/{
	var fname = tempfile('.sheetjs');
	console.log(fname);
	var ostream = fs.createWriteStream(fname);
	stream.pipe(ostream);
	ostream.on('finish', function() {
		var workbook = XLSX.readFile(fname);
		fs.unlinkSync(fname);

		/* DO SOMETHING WITH workbook IN THE CALLBACK */
		cb(workbook);
	});
}
```

</details>

