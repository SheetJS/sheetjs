function getdatastr(data)/*:?string*/ {
	if(!data) return null;
	if(data.data) return debom(data.data);
	if(data.asNodeBuffer && has_buf) return debom(data.asNodeBuffer().toString('binary'));
	if(data.asBinary) return debom(data.asBinary());
	if(data._data && data._data.getContent) return debom(cc2str(Array.prototype.slice.call(data._data.getContent(),0)));
	if(data.content && data.type) return debom(cc2str(data.content));
	return null;
}

function getdatabin(data) {
	if(!data) return null;
	if(data.data) return char_codes(data.data);
	if(data.asNodeBuffer && has_buf) return data.asNodeBuffer();
	if(data._data && data._data.getContent) {
		var o = data._data.getContent();
		if(typeof o == "string") return char_codes(o);
		return Array.prototype.slice.call(o);
	}
	if(data.content && data.type) return data.content;
	return null;
}

function getdata(data) { return (data && data.name.slice(-4) === ".bin") ? getdatabin(data) : getdatastr(data); }

/* Part 2 Section 10.1.2 "Mapping Content Types" Names are case-insensitive */
/* OASIS does not comment on filename case sensitivity */
function safegetzipfile(zip, file/*:string*/) {
	var k = zip.FullPaths || keys(zip.files);
	var f = file.toLowerCase(), g = f.replace(/\//g,'\\');
	for(var i=0; i<k.length; ++i) {
		var n = k[i].toLowerCase();
		if(f == n || g == n) return zip.files[k[i]];
	}
	return null;
}

function getzipfile(zip, file/*:string*/) {
	var o = safegetzipfile(zip, file);
	if(o == null) throw new Error("Cannot find file " + file + " in zip");
	return o;
}

function getzipdata(zip, file/*:string*/, safe/*:?boolean*/)/*:any*/ {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

function getzipstr(zip, file/*:string*/, safe/*:?boolean*/)/*:?string*/ {
	if(!safe) return getdatastr(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipstr(zip, file); } catch(e) { return null; }
}

function zipentries(zip) {
	var k = zip.FullPaths || keys(zip.files), o = [];
	for(var i = 0; i < k.length; ++i) if(k[i].slice(-1) != '/') o.push(k[i]);
	return o.sort();
}

function zip_add_file(zip, path, content) {
	if(zip.FullPaths) CFB.utils.cfb_add(zip, path, content);
	else zip.file(path, content);
}

var jszip;
/*:: declare var JSZipSync:any; */
/*global JSZipSync:true */
if(typeof JSZipSync !== 'undefined') jszip = JSZipSync;
if(typeof exports !== 'undefined') {
	if(typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = require('./jszip.js');
	}
}

function zip_new() {
	if(!jszip) return CFB.utils.cfb_new();
	return new jszip();
}

function zip_read(d, o) {
	var zip;
	if(jszip) switch(o.type) {
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": case "array": zip = new jszip(d, { base64:false }); break;
		case "buffer": zip = new jszip(d); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	else switch(o.type) {
		case "base64": zip = CFB.read(d, { type: "base64" }); break;
		case "binary": zip = CFB.read(d, { type: "binary" }); break;
		case "buffer": case "array": zip = CFB.read(d, { type: "buffer" }); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	return zip;
}

function resolve_path(path/*:string*/, base/*:string*/)/*:string*/ {
	if(path.charAt(0) == "/") return path.slice(1);
	var result = base.split('/');
	if(base.slice(-1) != "/") result.pop(); // folder path
	var target = path.split('/');
	while (target.length !== 0) {
		var step = target.shift();
		if (step === '..') result.pop();
		else if (step !== '.') result.push(step);
	}
	return result.join('/');
}
