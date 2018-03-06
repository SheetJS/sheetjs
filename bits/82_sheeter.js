function write_sheet_index(wb/*:Workbook*/, sheet/*:?string*/)/*:number*/ {
	if(!sheet) return 0;
	var idx = wb.SheetNames.indexOf(sheet);
	if(idx == -1) throw new Error("Sheet not found: " + sheet);
	return idx;
}

function write_obj_str(factory/*:WriteObjStrFactory*/) {
	return function write_str(wb/*:Workbook*/, o/*:WriteOpts*/)/*:string*/ {
		var idx = write_sheet_index(wb, o.sheet);
		return factory.from_sheet(wb.Sheets[wb.SheetNames[idx]], o, wb);
	};
}

var write_htm_str = write_obj_str(HTML_);
var write_csv_str = write_obj_str({from_sheet:sheet_to_csv});
var write_slk_str = write_obj_str(SYLK);
var write_dif_str = write_obj_str(DIF);
var write_prn_str = write_obj_str(PRN);
var write_rtf_str = write_obj_str(RTF);
var write_txt_str = write_obj_str({from_sheet:sheet_to_txt});
var write_dbf_buf = write_obj_str(DBF);
var write_eth_str = write_obj_str(ETH);

