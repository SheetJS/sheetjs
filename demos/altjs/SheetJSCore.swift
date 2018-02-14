/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import JavaScriptCore;

enum SJSError: Error {
	case badJSContext;
	case badJSWorkbook;
	case badJSWorksheet;
};

class SJSWorksheet {
	var context: JSContext!;
	var wb: JSValue!; var ws: JSValue!;
	var idx: Int32;

	func toCSV() throws -> String {
		let XLSX: JSValue! = context.objectForKeyedSubscript("XLSX");
		let utils: JSValue! = XLSX.objectForKeyedSubscript("utils");
		let sheet_to_csv: JSValue! = utils.objectForKeyedSubscript("sheet_to_csv");
		return sheet_to_csv.call(withArguments: [ws]).toString();
	}

	init(ctx: JSContext, workbook: JSValue, worksheet: JSValue, idx: Int32) throws {
		self.context = ctx; self.wb = workbook; self.ws = worksheet; self.idx = idx;
	}
}

class SJSWorkbook {
	var context: JSContext!;
	var wb: JSValue!; var SheetNames: JSValue!; var Sheets: JSValue!;

	func getSheetAtIndex(idx: Int32) throws -> SJSWorksheet {
		let SheetName: String = SheetNames.atIndex(Int(idx)).toString();
		let ws: JSValue! = Sheets.objectForKeyedSubscript(SheetName);
		return try SJSWorksheet(ctx: context, workbook: wb, worksheet: ws, idx: idx);
	}

	func writeBStr(bookType: String = "xlsx") throws -> String {
		let XLSX: JSValue! = context.objectForKeyedSubscript("XLSX");
		context.evaluateScript(String(format: "var writeopts = {type:'binary', bookType:'%@'}", bookType));
		let writeopts: JSValue! = context.objectForKeyedSubscript("writeopts");
		let writefunc: JSValue! = XLSX.objectForKeyedSubscript("write");
		return writefunc.call(withArguments: [wb, writeopts]).toString();
	}

	init(ctx: JSContext, wb: JSValue) throws {
		self.context = ctx;
		self.wb = wb;
		self.SheetNames = wb.objectForKeyedSubscript("SheetNames");
		self.Sheets = wb.objectForKeyedSubscript("Sheets");
	}
}

class SheetJSCore {
	var context: JSContext!;
	var XLSX: JSValue!;

	func init_context() throws -> JSContext {
		var context: JSContext!
		do {
			context = JSContext();
			context.exceptionHandler = { ctx, X in if let e = X { print(e.toString()); }; };
			context.evaluateScript("var global = (function(){ return this; }).call(null);");
			context.evaluateScript("if(typeof wbs == 'undefined') wbs = [];");
			let src = try String(contentsOfFile: "xlsx.full.min.js");
			context.evaluateScript(src);
			if context != nil { return context!; }
		} catch { print(error.localizedDescription); }
		throw SJSError.badJSContext;
	}

	func version() throws -> String {
		if let version = XLSX.objectForKeyedSubscript("version") { return version.toString(); }
		throw SJSError.badJSContext;
	}

	func readFile(file: String) throws -> SJSWorkbook {
		let data: String! = try String(contentsOfFile: file, encoding: String.Encoding.isoLatin1);
		return try readBStr(data: data);
	}

	func readBStr(data: String) throws -> SJSWorkbook {
		context.setObject(data, forKeyedSubscript: "payload" as (NSCopying & NSObjectProtocol)!);
		context.evaluateScript("var wb = XLSX.read(payload, {type:'binary'});");
		let wb: JSValue! = context.objectForKeyedSubscript("wb");
		if wb == nil { throw SJSError.badJSWorkbook; }
		return try SJSWorkbook(ctx: context, wb: wb);
	}

	init() throws {
		do {
			self.context = try init_context();
			self.XLSX = self.context.objectForKeyedSubscript("XLSX");
			if self.XLSX == nil { throw SJSError.badJSContext; }
		} catch { print(error.localizedDescription); }
	}
}
