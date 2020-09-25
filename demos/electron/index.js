/* xlsx.js (C) 2013-present SheetJS -- https://sheetjs.com */
const XLSX = require('xlsx');
const electron = require('electron').remote;

const EXTENSIONS = "xls|xlsx|xlsm|xlsb|xml|csv|txt|dif|sylk|slk|prn|ods|fods|htm|html".split("|");

const processWb = function(wb) {
	const HTMLOUT = document.getElementById('htmlout');
	const XPORT = document.getElementById('exportBtn');
	XPORT.disabled = false;
	HTMLOUT.innerHTML = "";
	wb.SheetNames.forEach(function(sheetName) {
		const htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:true});
		HTMLOUT.innerHTML += htmlstr;
	});
};

const readFile = function(files) {
	const f = files[0];
	const reader = new FileReader();
	reader.onload = function(e) {
		let data = e.target.result;
		data = new Uint8Array(data);
		processWb(XLSX.read(data, {type: 'array'}));
	};
	reader.readAsArrayBuffer(f);
};

const handleReadBtn = async function() {
	const o = await electron.dialog.showOpenDialog({
		title: 'Select a file',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}],
		properties: ['openFile']
	});
	if(o.filePaths.length > 0) processWb(XLSX.readFile(o.filePaths[0]));
};

const exportXlsx = async function() {
	const HTMLOUT = document.getElementById('htmlout');
	const wb = XLSX.utils.table_to_book(HTMLOUT);
	const o = await electron.dialog.showSaveDialog({
		title: 'Save file as',
		filters: [{
			name: "Spreadsheets",
			extensions: EXTENSIONS
		}]
	});
	console.log(o.filePath);
	XLSX.writeFile(wb, o.filePath);
	electron.dialog.showMessageBox({ message: "Exported data to " + o.filePath, buttons: ["OK"] });
};

// add event listeners
const readBtn = document.getElementById('readBtn');
const readIn = document.getElementById('readIn');
const exportBtn = document.getElementById('exportBtn');
const drop = document.getElementById('drop');

readBtn.addEventListener('click', handleReadBtn, false);
readIn.addEventListener('change', (e) => { readFile(e.target.files); }, false);
exportBtn.addEventListener('click', exportXlsx, false);
drop.addEventListener('dragenter', (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('dragover', (e) => {
	e.stopPropagation();
	e.preventDefault();
	e.dataTransfer.dropEffect = 'copy';
}, false);
drop.addEventListener('drop', (e) => {
	e.stopPropagation();
	e.preventDefault();
	readFile(e.dataTransfer.files);
}, false);
