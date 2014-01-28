function getdata(data) {
	if(!data) return null;
	if(data.data) return data.data;
	if(data._data && data._data.getContent) {
		/* TODO: something far more intelligent */
		if(data.name.substr(-4) === ".bin") return Array.prototype.slice.call(data._data.getContent());
		return Array.prototype.slice.call(data._data.getContent(),0).map(function(x) { return String.fromCharCode(x); }).join("");
	}
	return null;
}

function getzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	throw new Error("Cannot find file " + file + " in zip");
}

var _fs, jszip;
if(typeof JSZip !== 'undefined') jszip = JSZip;
if (typeof exports !== 'undefined') {
	if (typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = require('./jszip').JSZip;
		_fs = require('fs');
	}
}
