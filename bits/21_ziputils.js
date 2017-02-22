function getdatastr(data)/*:?string*/ {
	if(!data) return null;
	if(data.data) return debom_xml(data.data);
	if(data.asNodeBuffer && has_buf) return debom_xml(data.asNodeBuffer().toString('binary'));
	if(data.asBinary) return debom_xml(data.asBinary());
	if(data._data && data._data.getContent) return debom_xml(cc2str(Array.prototype.slice.call(data._data.getContent(),0)));
	return null;
}

function getdatabin(data) {
	if(!data) return null;
	if(data.data) return char_codes(data.data);
	if(data.asNodeBuffer && has_buf) return data.asNodeBuffer();
	if(data._data && data._data.getContent) {
		var o = data._data.getContent();
		if(typeof o == "string") return str2cc(o);
		return Array.prototype.slice.call(o);
	}
	return null;
}

function getdata(data) { return (data && data.name.slice(-4) === ".bin") ? getdatabin(data) : getdatastr(data); }

function safegetzipfile(zip, file/*:string*/) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	return null;
}

function getzipfile(zip, file/*:string*/) {
	var o = safegetzipfile(zip, file);
	if(o == null) throw new Error("Cannot find file " + file + " in zip");
	return o;
}

function getzipdata(zip, file/*:string*/, safe/*:?boolean*/) {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

function getzipstr(zip, file/*:string*/, safe/*:?boolean*/)/*:?string*/ {
	if(!safe) return getdatastr(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipstr(zip, file); } catch(e) { return null; }
}

var _fs, jszip;
/*:: declare var JSZip:any; */
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = require('./js'+'zip');
		_fs = require('f'+'s');
	}
}
