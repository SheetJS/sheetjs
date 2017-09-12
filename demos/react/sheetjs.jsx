/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
const SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(function(x) { return "." + x; }).join(",");

/*
  Simple HTML5 file drag-and-drop wrapper
  usage: <DragDropFile handleFile={handleFile}>...</DragDropFile>
    handleFile(file:File):void;
*/
class DragDropFile extends React.Component {
	constructor(props) {
		super(props);
		this.onDrop = this.onDrop.bind(this);
	};
	suppress(evt) { evt.stopPropagation(); evt.preventDefault(); };
	onDrop(evt) { evt.stopPropagation(); evt.preventDefault();
		const files = evt.dataTransfer.files;
		if(files && files[0]) this.props.handleFile(files[0]);
	};
	render() { return (
<div onDrop={this.onDrop} onDragEnter={this.suppress} onDragOver={this.suppress}>
	{this.props.children}
</div>
	); };
};

/*
  Simple HTML5 file input wrapper
  usage: <DataInput handleFile={callback} />
    handleFile(file:File):void;
*/
class DataInput extends React.Component {
	constructor(props) {
		super(props);
		this.handleChange = this.handleChange.bind(this);
	};
	handleChange(e) {
		const files = e.target.files;
		if(files && files[0]) this.props.handleFile(files[0]);
	};
	render() { return (
<form className="form-inline">
	<div className="form-group">
		<label htmlFor="file">Spreadsheet</label>
		<input type="file" className="form-control" id="file" accept={SheetJSFT} onChange={this.handleChange} />
	</div>
</form>
	); };
}

/* generate an array of column objects */
const make_cols = refstr => Array(XLSX.utils.decode_range(refstr).e.c + 1).fill(0).map((x,i) => ({name:XLSX.utils.encode_col(i), key:i}));

/*
  Simple HTML Table
  usage: <OutTable data={data} cols={cols} />
    data:Array<Array<any> >;
    cols:Array<{name:string, key:number|string}>;
*/
class OutTable extends React.Component {
	constructor(props) { super(props); };
	render() { return (
<div className="table-responsive">
	<table className="table table-striped">
		<thead>
			<tr>{this.props.cols.map((c) => <th>{c.name}</th>)}</tr>
		</thead>
		<tbody>
			{this.props.data.map(r => <tr>
				{this.props.cols.map(c => <td key={c.key}>{ r[c.key] }</td>)}
			</tr>)}
		</tbody>
	</table>
</div>
	); };
};

/* see Browser download file example in docs */
function s2ab(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf;
}

class SheetJSApp extends React.Component {
	constructor(props) {
		super(props);
		this.state = {
			data: [], /* Array of Arrays e.g. [["a","b"],[1,2]] */
			cols: []  /* Array of column objects e.g. { name: "C", key: 2 } */
		};
		this.handleFile = this.handleFile.bind(this);
		this.exportFile = this.exportFile.bind(this);
	};
	handleFile(file/*:File*/) {
		/* Boilerplate to set up FileReader */
		const reader = new FileReader();
		reader.onload = (e) => {
			/* Parse data */
			const bstr = e.target.result;
			const wb = XLSX.read(bstr, {type:'binary'});
			/* Get first worksheet */
			const wsname = wb.SheetNames[0];
			const ws = wb.Sheets[wsname];
			/* Convert array of arrays */
			const data = XLSX.utils.sheet_to_json(ws, {header:1});
			/* Update state */
			this.setState({ data: data, cols: make_cols(ws['!ref']) });
		};
		reader.readAsBinaryString(file);
	};
	exportFile() {
		/* convert state to workbook */
		const ws = XLSX.utils.aoa_to_sheet(this.state.data);
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
		/* generate XLSX file */
		const wbout = XLSX.write(wb, {type:"binary", bookType:"xlsx"});
		/* send to client */
		saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "sheetjs.xlsx");
	};
	render() { return (
<DragDropFile handleFile={this.handleFile}>
	<div className="row"><div className="col-xs-12">
		<DataInput handleFile={this.handleFile} />
	</div></div>
	<div className="row"><div className="col-xs-12">
		<button disabled={!this.state.data.length} className="btn btn-success" onClick={this.exportFile}>Export</button>
	</div></div>
	<div className="row"><div className="col-xs-12">
		<OutTable data={this.state.data} cols={this.state.cols} />
	</div></div>
</DragDropFile>
); };
};

if(typeof module !== 'undefined') module.exports = SheetJSApp
