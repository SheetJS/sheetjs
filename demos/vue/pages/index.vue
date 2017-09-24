<!-- xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com -->
<template>
<div @drop="_drop" @dragenter="_suppress" @dragover="_suppress">
	<div class="row"><div class="col-xs-12">
		<form class="form-inline">
			<div class="form-group">
				<label for="file">Spreadsheet</label>
				<input type="file" class="form-control" id="file" :accept="SheetJSFT" @change="_change" />
			</div>
		</form>
	</div></div>
	<div class="row"><div class="col-xs-12">
		<button :disabled="data.length ? false : true" class="btn btn-success" @click="_export">Export</button>
	</div></div>
	<div class="row"><div class="col-xs-12">
		<div class="table-responsive">
			<table class="table table-striped">
				<thead><tr>
					<th v-for="c in cols" :key="c.key">{{c.name}}</th>
				</tr></thead>
				<tbody>
					<tr v-for="(r, key) in data" :key="key">
						<td v-for="c in cols" :key="c.key"> {{ r[c.key] }}</td>
					</tr>
				</tbody>
			</table>
		</div>
	</div></div>
</div>
</template>

<script>
const _XLSX = require('xlsx');
const X = typeof XLSX !== 'undefined' ? XLSX : _XLSX;
const make_cols = refstr => Array(X.utils.decode_range(refstr).e.c + 1).fill(0).map((x,i) => ({name:X.utils.encode_col(i), key:i}));

/* see Browser download file example in docs */
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

const _SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function(x) { return "." + x; }).join(",");
export default {
	data() {
		return {
			data: ["SheetJS".split(""), "1234567".split("")],
			cols: [
				{name:"A", key:0},
				{name:"B", key:1},
				{name:"C", key:2},
				{name:"D", key:3},
				{name:"E", key:4},
				{name:"F", key:5},
				{name:"G", key:6},
			],
			SheetJSFT: _SheetJSFT
	}; },
	methods: {
		_suppress(evt) { evt.stopPropagation(); evt.preventDefault(); },
		_drop(evt) {
			evt.stopPropagation(); evt.preventDefault();
			const files = evt.dataTransfer.files;
			if(files && files[0]) this._file(files[0]);
		},
		_change(evt) {
			const files = evt.target.files;
			if(files && files[0]) this._file(files[0]);
		},
		_export(evt) {
			/* convert state to workbook */
			const ws = X.utils.aoa_to_sheet(this.data);
			const wb = X.utils.book_new();
			X.utils.book_append_sheet(wb, ws, "SheetJS");
			/* generate X file */
			const wbout = X.write(wb, {type:"binary", bookType:"xlsx"});
			/* send to client */
			saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "sheetjs.xlsx");
		},
		_file(file) {
			/* Boilerplate to set up FileReader */
			const reader = new FileReader();
			reader.onload = (e) => {
				/* Parse data */
				const bstr = e.target.result;
				const wb = X.read(bstr, {type:'binary'});
				/* Get first worksheet */
				const wsname = wb.SheetNames[0];
				const ws = wb.Sheets[wsname];
				/* Convert array of arrays */
				const data = X.utils.sheet_to_json(ws, {header:1});
				/* Update state */
				this.data = data;
				this.cols = make_cols(ws['!ref']);
			};
			reader.readAsBinaryString(file);
		}
	}
};
</script>
