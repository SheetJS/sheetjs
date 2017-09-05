/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
package com.sheetjs;

import org.mozilla.javascript.Function;
import org.mozilla.javascript.NativeObject;

public class SheetJSSheet {
	public NativeObject ws;
	public SheetJSFile wb;
	public SheetJSSheet(SheetJSFile wb, int idx) throws ObjectNotFoundException {
		this.wb = wb;
		this.ws = (NativeObject)JSHelper.get_object("Sheets." + wb.get_sheet_names()[idx],wb.wb);
	}
	public String get_range() throws ObjectNotFoundException {
		return JSHelper.get_object("!ref",this.ws).toString();
	}
	public String get_string_value(String address) throws ObjectNotFoundException {
		return JSHelper.get_object(address + ".v",this.ws).toString();
	}

	public String get_csv() throws ObjectNotFoundException {
		Function csvify = (Function)JSHelper.get_object("XLSX.utils.sheet_to_csv",this.wb.sheetjs.scope);
		Object csvArgs[] = {this.ws};
		Object csv = csvify.call(this.wb.sheetjs.cx, this.wb.sheetjs.scope, this.wb.sheetjs.scope, csvArgs);
		return csv.toString();
	}
}

