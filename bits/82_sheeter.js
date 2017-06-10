/* actual implementation elsewhere, wrappers are for read/write */
function write_obj_str(factory/*:WriteObjStrFactory*/) {
	return function write_str(wb/*:Workbook*/, o/*:WriteOpts*/)/*:string*/ {
		var idx = 0;
		for(var i=0;i<wb.SheetNames.length;++i) if(wb.SheetNames[i] == o.sheet) idx=i;
		if(idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) throw new Error("Sheet not found: " + o.sheet);
		return factory.from_sheet(wb.Sheets[wb.SheetNames[idx]], o, wb);
	};
}

var write_htm_str = write_obj_str(HTML_);
var write_csv_str = write_obj_str({from_sheet:sheet_to_csv});
var write_slk_str = write_obj_str(SYLK);
var write_dif_str = write_obj_str(DIF);
var write_prn_str = write_obj_str(PRN);
var write_txt_str = write_obj_str({from_sheet:sheet_to_txt});
