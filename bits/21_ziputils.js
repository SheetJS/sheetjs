function getdata(data) {
	if(!data) return null;
	if(data.name.substr(-4) === ".bin") {
		if(data.data) return char_codes(data.data);
		if(data.asNodeBuffer && typeof Buffer !== 'undefined') return data.asNodeBuffer();
		if(data._data && data._data.getContent) return Array.prototype.slice.call(data._data.getContent());
	} else {
		if(data.data) return data.name.substr(-4) !== ".bin" ? debom_xml(data.data) : char_codes(data.data);
		if(data.asNodeBuffer && typeof Buffer !== 'undefined') return debom_xml(data.asNodeBuffer().toString('binary'));
		if(data.asBinary) return debom_xml(data.asBinary());
		if(data._data && data._data.getContent) return debom_xml(cc2str(Array.prototype.slice.call(data._data.getContent(),0)));
	}
	return null;
}

function getzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	throw new Error("Cannot find file " + file + " in zip");
}

function getzipdata(zip, file, safe) {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

var _fs, jszip;
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof Buffer !== 'undefined' && typeof jszip === 'undefined') jszip = require('jszip');
		if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
		_fs = require('fs');
	}
}
