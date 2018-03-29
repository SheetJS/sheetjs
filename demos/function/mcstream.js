/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* eslint-env node */
// cat file.xlsx | curl --header 'content-type: application/octet-stream' --data-binary @- "http://localhost:3000/"
const XLSX = require('xlsx');

const process_RS = (stream, cb) => {
  var buffers = [];
  stream.on('data', function(data) { buffers.push(data); });
  stream.on('end', function() {
    var buffer = Buffer.concat(buffers);
    var workbook = XLSX.read(buffer, {type:"buffer"});
    cb(workbook);
  });
};

module.exports = (hook) => {
	process_RS(hook.req, (wb) => {
		hook.res.writeHead(200, { 'Content-Type': 'text/csv' });
		const stream = XLSX.stream.to_csv(wb.Sheets[wb.SheetNames[0]]);
		stream.pipe(hook.res);
	});
};
