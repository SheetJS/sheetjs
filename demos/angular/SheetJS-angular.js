/* xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
function SheetJSExportService(uiGridExporterService) {

	function exportSheetJS(gridApi, wopts) {
		var columns = uiGridExporterService.getColumnHeaders(gridApi.grid, 'all');
		var data = uiGridExporterService.getData(gridApi.grid, 'all', 'all');

		var fileName = gridApi.grid.options.filename || 'SheetJS';
		fileName += wopts.bookType ? "." + wopts.bookType : '.xlsx';

		var sheetName = gridApi.grid.options.sheetname || 'Sheet1';

		var wb = XLSX.utils.book_new(), ws = uigrid_to_sheet(data, columns);
		XLSX.utils.book_append_sheet(wb, ws, sheetName);
		var wbout = XLSX.write(wb, wopts);
		saveAs(new Blob([s2ab(wbout)], { type: 'application/octet-stream' }), fileName);
	}

	var service = {};
	service.exportXLSB = function exportXLSB(gridApi) { return exportSheetJS(gridApi, { bookType: 'xlsb', bookSST: true, type: 'binary' }); };
	service.exportXLSX = function exportXLSX(gridApi) { return exportSheetJS(gridApi, { bookType: 'xlsx', bookSST: true, type: 'binary' }); }

	return service;

	/* utilities */
	function uigrid_to_sheet(data, columns) {
		var o = [], oo = [], i = 0, j = 0;

		/* column headers */
		for(j = 0; j < columns.length; ++j) oo.push(get_value(columns[j]));
		o.push(oo);

		/* table data */
		for(i = 0; i < data.length; ++i) {
			oo = [];
			for(j = 0; j < data[i].length; ++j) oo.push(get_value(data[i][j]));
			o.push(oo);
		}
		/* aoa_to_sheet converts an array of arrays into a worksheet object */
		return XLSX.utils.aoa_to_sheet(o);
	}

	function get_value(col) {
		if(!col) return col;
		if(col.value) return col.value;
		if(col.displayName) return col.displayName;
		if(col.name) return col.name;
		return null;
	}

	function s2ab(s) {
		if(typeof ArrayBuffer !== 'undefined') {
			var buf = new ArrayBuffer(s.length);
			var view = new Uint8Array(buf);
			for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
			return buf;
		} else {
			var buf = new Array(s.length);
			for (var i=0; i!=s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
			return buf;
		}
	}
}

var SheetJSImportDirective = function() {
	return {
		scope: { opts: '=' },
		link: function ($scope, $elm, $attrs) {
			$elm.on('change', function (changeEvent) {
				var reader = new FileReader();

				reader.onload = function (e) {
					/* read workbook */
					var bstr = e.target.result;
					var wb = XLSX.read(bstr, {type:'binary'});

					/* grab first sheet */
					var wsname = wb.SheetNames[0];
					var ws = wb.Sheets[wsname];

					/* grab first row and generate column headers */
					var aoa = XLSX.utils.sheet_to_json(ws, {header:1, raw:false});
					var cols = [];
					for(var i = 0; i < aoa[0].length; ++i) cols[i] = { field: aoa[0][i] };

					/* generate rest of the data */
					var data = [];
					for(var r = 1; r < aoa.length; ++r) {
						data[r-1] = {};
						for(i = 0; i < aoa[r].length; ++i) {
							if(aoa[r][i] == null) continue;
							data[r-1][aoa[0][i]] = aoa[r][i]
						}
					}

					/* update scope */
					$scope.$apply(function() {
						$scope.opts.columnDefs = cols;
						$scope.opts.data = data;
					});
				};

				reader.readAsBinaryString(changeEvent.target.files[0]);
			});
		}
	};
};
