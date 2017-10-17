var current_codepage = 1200;
/*:: declare var cptable:any; */
/*global cptable:true */
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') global.cptable = require('./dist/cpexcel.js');
}
function reset_cp() { set_cp(1200); }
var set_cp = function(cp) { current_codepage = cp; };

function char_codes(data) { var o = []; for(var i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i); return o; }

function utf16leread(data/*:string*/)/*:string*/ {
	var o = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i) + (data.charCodeAt(2*i+1)<<8));
	return o.join("");
}
function utf16beread(data/*:string*/)/*:string*/ {
	var o = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i+1) + (data.charCodeAt(2*i)<<8));
	return o.join("");
}

var debom = function(data/*:string*/)/*:string*/ {
	var c1 = data.charCodeAt(0), c2 = data.charCodeAt(1);
	if(c1 == 0xFF && c2 == 0xFE) return utf16leread(data.substr(2));
	if(c1 == 0xFE && c2 == 0xFF) return utf16beread(data.substr(2));
	if(c1 == 0xFEFF) return data.substr(1);
	return data;
};

var _getchar = function _gc1(x/*:number*/)/*:string*/ { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') {
	set_cp = function(cp/*:number*/) { current_codepage = cp; };
	debom = function(data/*:string*/) {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.substr(2))); }
		return data;
	};
	_getchar = function _gc2(x/*:number*/)/*:string*/ {
		if(current_codepage === 1200) return String.fromCharCode(x);
		return cptable.utils.decode(current_codepage, [x&255,x>>8])[0];
	};
}
