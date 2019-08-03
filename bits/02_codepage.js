var current_codepage = 1200, current_ansi = 1252;
/*:: declare var cptable:any; */
/*global cptable:true, window */
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') {
		if(typeof global !== 'undefined') global.cptable = require('./dist/cpexcel.js');
		else if(typeof window !== 'undefined') window.cptable = require('./dist/cpexcel.js');
	}
}

var VALID_ANSI = [ 874, 932, 936, 949, 950 ];
for(var i = 0; i <= 8; ++i) VALID_ANSI.push(1250 + i);
/* ECMA-376 Part I 18.4.1 charset to codepage mapping */
var CS2CP = ({
	/*::[*/0/*::]*/:    1252, /* ANSI */
	/*::[*/1/*::]*/:   65001, /* DEFAULT */
	/*::[*/2/*::]*/:   65001, /* SYMBOL */
	/*::[*/77/*::]*/:  10000, /* MAC */
	/*::[*/128/*::]*/:   932, /* SHIFTJIS */
	/*::[*/129/*::]*/:   949, /* HANGUL */
	/*::[*/130/*::]*/:  1361, /* JOHAB */
	/*::[*/134/*::]*/:   936, /* GB2312 */
	/*::[*/136/*::]*/:   950, /* CHINESEBIG5 */
	/*::[*/161/*::]*/:  1253, /* GREEK */
	/*::[*/162/*::]*/:  1254, /* TURKISH */
	/*::[*/163/*::]*/:  1258, /* VIETNAMESE */
	/*::[*/177/*::]*/:  1255, /* HEBREW */
	/*::[*/178/*::]*/:  1256, /* ARABIC */
	/*::[*/186/*::]*/:  1257, /* BALTIC */
	/*::[*/204/*::]*/:  1251, /* RUSSIAN */
	/*::[*/222/*::]*/:   874, /* THAI */
	/*::[*/238/*::]*/:  1250, /* EASTEUROPE */
	/*::[*/255/*::]*/:  1252, /* OEM */
	/*::[*/69/*::]*/:   6969  /* MISC */
}/*:any*/);

var set_ansi = function(cp/*:number*/) { if(VALID_ANSI.indexOf(cp) == -1) return; current_ansi = CS2CP[0] = cp; };
function reset_ansi() { set_ansi(1252); }

var set_cp = function(cp/*:number*/) { current_codepage = cp; set_ansi(cp); };
function reset_cp() { set_cp(1200); reset_ansi(); }

function char_codes(data/*:string*/)/*:Array<number>*/ { var o/*:Array<number>*/ = []; for(var i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i); return o; }

function utf16leread(data/*:string*/)/*:string*/ {
	var o/*:Array<string>*/ = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i) + (data.charCodeAt(2*i+1)<<8));
	return o.join("");
}
function utf16beread(data/*:string*/)/*:string*/ {
	var o/*:Array<string>*/ = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i+1) + (data.charCodeAt(2*i)<<8));
	return o.join("");
}

var debom = function(data/*:string*/)/*:string*/ {
	var c1 = data.charCodeAt(0), c2 = data.charCodeAt(1);
	if(c1 == 0xFF && c2 == 0xFE) return utf16leread(data.slice(2));
	if(c1 == 0xFE && c2 == 0xFF) return utf16beread(data.slice(2));
	if(c1 == 0xFEFF) return data.slice(1);
	return data;
};

var _getchar = function _gc1(x/*:number*/)/*:string*/ { return String.fromCharCode(x); };
var _getansi = function _ga1(x/*:number*/)/*:string*/ { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') {
	set_cp = function(cp/*:number*/) { current_codepage = cp; set_ansi(cp); };
	debom = function(data/*:string*/) {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.slice(2))); }
		return data;
	};
	_getchar = function _gc2(x/*:number*/)/*:string*/ {
		if(current_codepage === 1200) return String.fromCharCode(x);
		return cptable.utils.decode(current_codepage, [x&255,x>>8])[0];
	};
	_getansi = function _ga2(x/*:number*/)/*:string*/ {
		return cptable.utils.decode(current_ansi, [x])[0];
	};
}
