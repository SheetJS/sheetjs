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
var write_slk_str = write_obj_str(typeof SYLK !== "undefined" ? SYLK : {});
var write_dif_str = write_obj_str(typeof DIF !== "undefined" ? DIF : {});
var write_prn_str = write_obj_str(typeof PRN !== "undefined" ? PRN : {});
var write_rtf_str = write_obj_str(typeof RTF !== "undefined" ? RTF : {});
var write_txt_str = write_obj_str({from_sheet:sheet_to_txt});
var write_dbf_buf = write_obj_str(typeof DBF !== "undefined" ? DBF : {});
var write_eth_str = write_obj_str(typeof ETH !== "undefined" ? ETH : {});

