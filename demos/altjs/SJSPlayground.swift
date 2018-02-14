/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
/* This only works in a playground, see SheetJSCore.swift for standalone use */
import JavaScriptCore;
import PlaygroundSupport;

/* build path variable for the library */
let shared_dir = PlaygroundSupport.playgroundSharedDataDirectory;
let lib_path = shared_dir.appendingPathComponent("xlsx.full.min.js");

/* prepare JS context */
var context: JSContext! = JSContext();
var src = "var global = (function(){ return this; }).call(null);";
context.evaluateScript(src);

/* load library */
var lib = try? String(contentsOf: lib_path);
context.evaluateScript(lib);
let XLSX: JSValue! = context.objectForKeyedSubscript("XLSX");

/* to verify the library was loaded, get the version string */
let XLSXversion: JSValue! = XLSX.objectForKeyedSubscript("version")
var version  = XLSXversion.toString();

/* parse sheetjs.xls */
let file_path = shared_dir.appendingPathComponent("sheetjs.xls");
let data: String! = try String(contentsOf: file_path, encoding: String.Encoding.isoLatin1);
context.setObject(data, forKeyedSubscript: "payload" as (NSCopying & NSObjectProtocol)!);
src = "var wb = XLSX.read(payload, {type:'binary'});";
context.evaluateScript(src);

/* write to sheetjsw.xlsx  */
let out_path = shared_dir.appendingPathComponent("sheetjsw.xlsx");
src = "var out = XLSX.write(wb, {type:'binary', bookType:'xlsx'})";
context.evaluateScript(src);
let outvalue: JSValue! = context.objectForKeyedSubscript("out");
var out: String! = outvalue.toString();
try? out.write(to: out_path, atomically: false, encoding: String.Encoding.isoLatin1);
