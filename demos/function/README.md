# "Serverless" Functions

Because the library is pure JS, the hard work of reading and writing files can
be performed in the client browser or on the server side.  On the server side,
the mechanical process is essentially independent from the data parsing or
generation.  As a result, it is sometimes sensible to organize applications so
that the "last mile" conversion between JSON data and spreadsheet files is
independent from the main application.

The most obvious architecture would split off the JSON data conversion as a
separate microservice or application.  Since it is only needed when an import or
export is requested, and since the process itself is relatively independent from
the rest of a typical service, a "Serverless" architecture makes a great fit.
Since the "function" is separate from the rest of the application, it is easy to
integrate into a platform built in Java or Go or Python or another language!

This demo discusses general architectures and provides examples for popular
commercial systems and self-hosted alternatives.  The examples are merely
intended to demonstrate very basic functionality.


## Simple Strategies

#### Data Normalization

Most programming languages and platforms can process CSV or JSON but can't use
XLS or XLSX or XLSB directly.  Form data from an HTTP POST request can be parsed
and contained files can be converted to CSV or JSON.  The `XLSX.stream.to_csv`
utility can stream rows to a standard HTTP response.  `XLSX.utils.sheet_to_json`
can generate an array of objects that can be fed to another service.

At the simplest level, a file on the filesystem can be converted using the bin
script that ships with the `npm` module:

```bash
$ xlsx /path/to/uploads/file > /tmp/new_csv_file
```

From a utility script, workbooks can be converted in two lines:

```js
var workbook = XLSX.readFile("path/to/file.xlsb");
XLSX.writeFile(workbook, "output/path/file.csv");
```

The `mcstream.js` demo uses the `microcule` framework to show a simple body
converter.  It accepts raw data from a POST connection, parses as a workbook,
and streams back the first worksheet as CSV:

<details>
	<summary><b>Code Sketch</b> (click to show)</summary>

```js
const XLSX = require('xlsx');

module.exports = (hook) => {
	/* process_RS from the main README under "Streaming Read" section */
	process_RS(hook.req, (wb) => {
		hook.res.writeHead(200, { 'Content-Type': 'text/csv' });
		/* get first worksheet */
		const ws = wb.Sheets[wb.SheetNames[0]];
		/* generate CSV stream and pipe to response */
		const stream = XLSX.stream.to_csv(ws);
		stream.pipe(hook.res);
	});
};
```

</details>


#### Report Generation

For an existing platform that already generates JSON or CSV or HTML output, it
is very easy to embellish output in an Excel-friendly XLSX file.  The
`XLSX.utils.sheet_add_json` and `XLSX.utils.sheet_add_aoa` functions can add
data rows to an existing worksheet:

```js
var ws = XLSX.utils.aoa_to_sheet([
	["Company Report"],
	[],
	["Item", "Cost"]
]);
XLSX.utils.sheet_add_json(ws, [
	{ item: "Coffee", cost: 5 },
	{ item: "Cake", cost: 20 }
], { skipHeader: true, origin: -1, header: ["item", "cost"] });
```


## Deployment Targets

The library is supported in Node versions starting from `0.8` as well as a
myriad of ES3 and ES5 compatible JS engines.  All major services use Node
versions beyond major release 4, so there should be no problem directly using
the library in those environments.

Note that most cloud providers proactively convert form data to UTF8 strings.
This is especially problematic when dealing with XLSX and XLSB files, as they
naturally contain codes that are not valid UTF8 characters.  As a result, these
demos specifically handle Base64-encoded files only.  To test on the command
line, use the `base64` tool to encode data before piping to `curl`:

```
base64 test.xlsb | curl -F "data=@-;filename=test.xlsb" http://localhost/
```

#### AWS Lambda

Through the AWS Gateway API, Lambda functions can be triggered on HTTP requests.
The `LambdaProxy` example reads files from form data and converts to CSV.

When deploying on AWS, be sure to `npm install` locally and include the modules
in the ZIP file.

#### Azure Functions

Azure supports many types of triggers.  The `AzureHTTPTrigger` shows an example
HTTP trigger that converts the submitted file to CSV.

When deploying on Azure, be sure to install the module from the remote console,
as described in the "Azure Functions JavaScript developer guide".
