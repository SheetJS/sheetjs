/* actual implementation in utils, wrappers are for read/write */
function write_csv_str(wb/*:Workbook*/, o/*:WriteOpts*/) {
	var idx = 0;
	for(var i=0;i<wb.SheetNames.length;++i) if(wb.SheetNames[i] == o.sheet) idx=i;
	if(idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) throw new Error("Sheet not found: " + o.sheet);
	return sheet_to_csv(wb.Sheets[wb.SheetNames[idx]], o);
}
