/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* Notes:
   - usage: `ReactDOM.render( <SheetJSApp />, document.getElementById('app') );`
   - xlsx.full.min.js is loaded in the head of the HTML page
   - this script should be referenced with type="text/babel"
   - babel.js in-browser transpiler should be loaded before this script
*/
function SheetJSApp() {
	const [data, setData] = React.useState([]);
	const [cols, setCols] = React.useState([]);

	const handleFile = (file) => {
		const reader = new FileReader();
		const rABS = !!reader.readAsBinaryString;
		reader.onload = (e) => {
			/* Parse data */
			const bstr = e.target.result;
			const wb = XLSX.read(bstr, {type:rABS ? 'binary' : 'array'});
			/* Get first worksheet */
			const wsname = wb.SheetNames[0];
			const ws = wb.Sheets[wsname];
			/* Convert array of arrays */
			const data = XLSX.utils.sheet_to_json(ws, {header:1});
			/* Update state */
			setData(data);
			setCols(make_cols(ws['!ref']))
		};
		if(rABS) reader.readAsBinaryString(file); else reader.readAsArrayBuffer(file);
	}

	const exportFile = () => {
		/* convert state to workbook */
		const ws = XLSX.utils.aoa_to_sheet(data);
		const wb = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
		/* generate XLSX file and send to client */
		XLSX.writeFile(wb, "sheetjs.xlsx")
	};

	return (
	<DragDropFile handleFile={handleFile}>
		<div className="row"><div className="col-xs-12">
			<DataInput handleFile={handleFile} />
		</div></div>
		<div className="row"><div className="col-xs-12">
			<button disabled={!data.length} className="btn btn-success" onClick={exportFile}>Export</button>
		</div></div>
		<div className="row"><div className="col-xs-12">
			<OutTable data={data} cols={cols} />
		</div></div>
	</DragDropFile>
	);
}

if(typeof module !== 'undefined') module.exports = SheetJSApp

/* -------------------------------------------------------------------------- */

/*
  Simple HTML5 file drag-and-drop wrapper
  usage: <DragDropFile handleFile={handleFile}>...</DragDropFile>
    handleFile(file:File):void;
*/

function DragDropFile({ handleFile, children }) {
	const suppress = (e) => { e.stopPropagation(); e.preventDefault(); };
	const handleDrop = (e) => { e.stopPropagation(); e.preventDefault();
		const files = e.dataTransfer.files;
		if(files && files[0]) handleFile(files[0]);
	};

	return (
		<div 
			onDrop={handleDrop} 
			onDragEnter={suppress} 
			onDragOver={suppress}
		>
		{children}
		</div>
	);
}

/*
  Simple HTML5 file input wrapper
  usage: <DataInput handleFile={callback} />
    handleFile(file:File):void;
*/

function DataInput({ handleFile }) {
	const handleChange = (e) => {
		const files = e.target.files;
		if(files && files[0]) handleFile(files[0]);
	};

	return (
		<form className="form-inline">
			<div className="form-group">
				<label htmlFor="file">Drag or choose a spreadsheet file</label>
				<br />
				<input 
					type="file" 
					className="form-control" 
					id="file" 
					accept={SheetJSFT} 
					onChange={handleChange} 
				/>
			</div>
		</form>
	)
}

/*
  Simple HTML Table
  usage: <OutTable data={data} cols={cols} />
    data:Array<Array<any> >;
    cols:Array<{name:string, key:number|string}>;
*/
function OutTable({ data, cols }) {
	return (
		<div className="table-responsive">
			<table className="table table-striped">
				<thead>
					<tr>{cols.map((c) => <th key={c.key}>{c.name}</th>)}</tr>
				</thead>
				<tbody>
					{data.map((r,i) => <tr key={i}>
						{cols.map(c => <td key={c.key}>{ r[c.key] }</td>)}
					</tr>)}
				</tbody>
			</table>
		</div>
	);
}

/* list of supported file types */
const SheetJSFT = [
	"xlsx", "xlsb", "xlsm", "xls", "xml", "csv", "txt", "ods", "fods", "uos", "sylk", "dif", "dbf", "prn", "qpw", "123", "wb*", "wq*", "html", "htm"
].map(x => `.${x}`).join(",");

/* generate an array of column objects */
const make_cols = refstr => {
	let o = [], C = XLSX.utils.decode_range(refstr).e.c + 1;
	for(var i = 0; i < C; ++i) o[i] = {name:XLSX.utils.encode_col(i), key:i}
	return o;
};
