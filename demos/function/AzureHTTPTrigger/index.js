/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
// base64 sheetjs.xlsb | curl -F "data=@-;filename=test.xlsb" http://localhost:7262/api/AzureHTTPTrigger

const XLSX = require('xlsx');
const formidable = require('formidable');
const Readable = require('stream').Readable;
var fs = require('fs');

/* formidable expects the request object to be a stream */
const streamify = (req) => {
	if(typeof req.on !== 'undefined') return req;
	const s = new Readable();
	s._read = ()=>{};
	s.push(new Buffer(req.body));
	s.push(null);
	Object.assign(s, req);
	return s;
};

module.exports = (context, req) => {
	const form = new formidable.IncomingForm();
	form.parse(streamify(req), (err, fields, files) => {
		/* grab the first file */
		var f = Object.values(files)[0];
		if(!f) {
			context.res = { status: 400, body: "Must submit a file for processing!" };
		} else {
			/* since the file is Base64-encoded, read the file and parse as "base64" */
			const b64 = fs.readFileSync(f.path).toString();
			const wb = XLSX.read(b64, {type:"base64"});

			/* convert to specified output type -- default CSV */
			const ext = (fields.bookType || "csv").toLowerCase();
			const out = XLSX.write(wb, {type:"string", bookType:ext});

			context.res = {
				status: 200,
				headers: { "Content-Disposition": `attachment; filename="download.${ext}";` },
				body: out
			};
		}
		context.done();
	});
};
