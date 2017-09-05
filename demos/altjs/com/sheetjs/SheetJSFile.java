/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
package com.sheetjs;

import org.mozilla.javascript.NativeObject;
import org.mozilla.javascript.Function;

public class SheetJSFile {
	public NativeObject wb;
	public SheetJS sheetjs;
	public SheetJSFile() {}
	public SheetJSFile(NativeObject wb, SheetJS sheetjs) { this.wb = wb; this.sheetjs = sheetjs; }
	public String[] get_sheet_names() {
		try {
			return JSHelper.get_string_array("SheetNames", this.wb);
		} catch(ObjectNotFoundException e) {
			return null;
		}
	}
	public SheetJSSheet get_sheet(int idx) throws ObjectNotFoundException {
		return new SheetJSSheet(this, idx);
	}
}

