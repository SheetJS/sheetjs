#!/usr/bin/env xcrun swift
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
import JavaScriptCore;

class SheetJS {
	var context: JSContext!;
	var XLSX: JSValue!;

	enum SJSError: Error {
		case badJSContext;
	};

	func init_context() throws -> JSContext {
		var context: JSContext!
		do {
			context = JSContext();
			context.exceptionHandler = { ctx, X in if let e = X { print(e.toString()); }; }
			var src = "var global = (function(){ return this; }).call(null);";
			context.evaluateScript(src);
			src = try String(contentsOfFile: "xlsx.swift.js");
			context.evaluateScript(src);
			if context != nil { return context!; }
		} catch { print(error.localizedDescription); }
		throw SheetJS.SJSError.badJSContext;
	}

	func version() throws -> String {
		if let version = XLSX.objectForKeyedSubscript("version") { return version.toString(); }
		throw SheetJS.SJSError.badJSContext;
	}

	func readFileToCSV(file: String) throws -> String {
		let data:String! = try String(contentsOfFile: file, encoding:String.Encoding.isoLatin1);
		self.context.setObject(data, forKeyedSubscript:"payload" as (NSCopying & NSObjectProtocol)!);

		let src = [
			"var wb = XLSX.read(payload, {type:'binary'});",
			"var ws = wb.Sheets[wb.SheetNames[0]];",
			"var result = XLSX.utils.sheet_to_csv(ws);"
		].joined(separator: "\n");
		self.context.evaluateScript(src);

		return context.objectForKeyedSubscript("result").toString();
	}

	init() throws {
		do {
			self.context = try init_context();
			self.XLSX = context.objectForKeyedSubscript("XLSX");
			if self.XLSX == nil {
				throw SheetJS.SJSError.badJSContext;
			}
		} catch { print(error.localizedDescription); }
	}
}

let sheetjs = try SheetJS();
try print(sheetjs.version());
try print(sheetjs.readFileToCSV(file:"sheetjs.xlsx"));
try print(sheetjs.readFileToCSV(file:"sheetjs.xlsb"));
try print(sheetjs.readFileToCSV(file:"sheetjs.xls"));
try print(sheetjs.readFileToCSV(file:"sheetjs.xml.xls"));
