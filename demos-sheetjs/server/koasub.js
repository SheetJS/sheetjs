/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
const XLSX = require('xlsx');
let data = "a,b,c\n1,2,3".split("\n").map(x => x.split(","));
process.on('message', ([m, data] = _) => {
	switch(m) {
		case 'load data': load_data(data); break;
		case 'get data': get_data(data); break;
		case 'get file': get_file(data); break;
	}
});

function load_data(file) {
	var wb = XLSX.readFile(file);
	/* generate array of arrays */
	data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
	console.log(data);
	process.send("done");
}

/* helper to generate the workbook object */
function make_book() {
	var ws = XLSX.utils.aoa_to_sheet(data);
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
	return wb;
}

function get_data(type) {
	var wb = make_book();
	/* send buffer back */
	process.send(XLSX.write(wb, {type:'buffer', bookType:type}));
}

function get_file(file) {
	var wb = make_book();
	/* write using XLSX.writeFile */
	XLSX.writeFile(wb, file);
	process.send("wrote to " + file + "\n");
}
