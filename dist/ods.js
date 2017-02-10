/* ods.js (C) 2014-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint -W041 */
var ODS = {};
(function make_ods(ODS) {
/* Open Document Format for Office Applications (OpenDocument) Version 1.2 */
var get_utils = function() {
	if(typeof XLSX !== 'undefined') return XLSX.utils;
	if(typeof module !== "undefined" && typeof require !== 'undefined') try {
		return require('../' + 'xlsx').utils;
	} catch(e) {
		try { return require('./' + 'xlsx').utils; }
		catch(ee) { return require('xl' + 'sx').utils; }
	}
	throw new Error("Cannot find XLSX utils");
};
var has_buf = (typeof Buffer !== 'undefined');

function cc2str(arr) {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function dup(o) {
	if(typeof JSON != 'undefined') return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || !o) return o;
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}
function getdata(data) {
	if(!data) return null;
	if(data.data) return data.data;
	if(data.asNodeBuffer && has_buf) return data.asNodeBuffer().toString('binary');
	if(data.asBinary) return data.asBinary();
	if(data._data && data._data.getContent) return cc2str(Array.prototype.slice.call(data._data.getContent(),0));
	return null;
}

function safegetzipfile(zip, file) {
	var f = file; if(zip.files[f]) return zip.files[f];
	f = file.toLowerCase(); if(zip.files[f]) return zip.files[f];
	f = f.replace(/\//g,'\\'); if(zip.files[f]) return zip.files[f];
	return null;
}

function getzipfile(zip, file) {
	var o = safegetzipfile(zip, file);
	if(o == null) throw new Error("Cannot find file " + file + " in zip");
	return o;
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
		if(typeof jszip === 'undefined') jszip = require('./js'+'zip');
		_fs = require('f'+'s');
	}
}
var attregexg=/\b[\w:-]+=["'][^"]*['"]/g;
var tagregex=/<[^>]*>/g;
var nsregex=/<\w*:/, nsregex2 = /<(\/?)\w+:/;
function parsexmltag(tag, skip_root) {
	var z = [];
	var eq = 0, c = 0;
	for(; eq !== tag.length; ++eq) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
	if(!skip_root) z[0] = tag.substr(0, eq);
	if(eq === tag.length) return z;
	var m = tag.match(attregexg), j=0, v="", i=0, q="", cc="";
	if(m) for(i = 0; i != m.length; ++i) {
		cc = m[i];
		for(c=0; c != cc.length; ++c) if(cc.charCodeAt(c) === 61) break;
		q = cc.substr(0,c); v = cc.substring(c+2, cc.length-1);
		for(j=0;j!=q.length;++j) if(q.charCodeAt(j) === 58) break;
		if(j===q.length) z[q] = v;
		else z[(j===5 && q.substr(0,5)==="xmlns"?"xmlns":"")+q.substr(j+1)] = v;
	}
	return z;
}
function strip_ns(x) { return x.replace(nsregex2, "<$1"); }

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = {
	'"': '&quot;',
	"'": '&apos;',
	'>': '&gt;',
	'<': '&lt;',
	'&': '&amp;'
};
var rencstr = "&<>'\"".split("");

// TODO: CP remap (need to read file version to determine OS)
var encregex = /&[a-z]*;/g, coderegex = /_x([\da-fA-F]+)_/g;
function unescapexml(text){
	var s = text + '';
	return s.replace(encregex, function($$) { return encodings[$$]; }).replace(coderegex,function(m,c) {return String.fromCharCode(parseInt(c,16));});
}
var decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
function escapexml(text){
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(charegex,function(s) { return "_x" + ("000"+s.charCodeAt(0).toString(16)).substr(-4) + "_";});
}

function parsexmlbool(value) {
	switch(value) {
		case '1': case 'true': case 'TRUE': return true;
		/* case '0': case 'false': case 'FALSE':*/
		default: return false;
	}
}

function datenum(v) {
	var epoch = Date.parse(v);
	return (epoch + 2209161600000) / (24 * 60 * 60 * 1000);
}

/* ISO 8601 Duration */
function parse_isodur(s) {
	var sec = 0, mt = 0, time = false;
	var m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
	if(!m) throw new Error("|" + s + "| is not an ISO8601 Duration");
	for(var i = 1; i != m.length; ++i) {
		if(!m[i]) continue;
		mt = 1;
		if(i > 3) time = true;
		switch(m[i].substr(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].substr(m[i].length-1));
			case 'D': mt *= 24;
				/* falls through */
			case 'H': mt *= 60;
				/* falls through */
			case 'M':
				if(!time) throw new Error("Unsupported ISO Duration Field: M");
				else mt *= 60;
				/* falls through */
			case 'S': break;
		}
		sec += mt * parseInt(m[i], 10);
	}
	return sec;
}

var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
/* copied from js-xls (C) SheetJS Apache2 license */
function xlml_normalize(d) {
	if(has_buf && Buffer.isBuffer(d)) return d.toString('utf8');
	if(typeof d === 'string') return d;
	throw "badf";
}

var xlmlregex = /<(\/?)([a-z0-9]*:|)([\w-]+)[^>]*>/mg;
/* Part 3 Section 4 Manifest File */
var CT_ODS = "application/vnd.oasis.opendocument.spreadsheet";
function parse_manifest(d, opts) {
	var str = xlml_normalize(d);
	var Rn;
	var FEtag;
	while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {
		case 'manifest': break; // 4.2 <manifest:manifest>
		case 'file-entry': // 4.3 <manifest:file-entry>
			FEtag = parsexmltag(Rn[0], false);
			if(FEtag.path == '/' && FEtag.type !== CT_ODS) throw new Error("This OpenDocument is not a spreadsheet");
			break;
		case 'encryption-data': // 4.4 <manifest:encryption-data>
		case 'algorithm': // 4.5 <manifest:algorithm>
		case 'start-key-generation': // 4.6 <manifest:start-key-generation>
		case 'key-derivation': // 4.7 <manifest:key-derivation>
			throw new Error("Unsupported ODS Encryption");
		default: if(opts && opts.WTF) throw Rn;
	}
}

function write_manifest(manifest, opts) {
	var o = [XML_HEADER];
	o.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
	o.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
	for(var i = 0; i < manifest.length; ++i) o.push('  <manifest:file-entry manifest:full-path="' + manifest[i][0] + '" manifest:media-type="' + manifest[i][1] + '"/>\n');
	o.push('</manifest:manifest>');
	return o.join("");
}
/* Part 3 Section 6 Metadata Manifest File */
function write_rdf_type(file, res, tag) {
	return [
		'  <rdf:Description rdf:about="' + file + '">\n',
		'    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/' + (tag || "odf") + '#' + res + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf_has(base, file) {
	return [
		'  <rdf:Description rdf:about="' + base + '">\n',
		'    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="' + file + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf(rdf, opts) {
	var o = [XML_HEADER];
	o.push('<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n');
	for(var i = 0; i != rdf.length; ++i) {
		o.push(write_rdf_type(rdf[i][0], rdf[i][1]));
		o.push(write_rdf_has("",rdf[i][0]));
	}
	o.push(write_rdf_type("","Document", "pkg"));
	o.push('</rdf:RDF>');
	return o.join("");
}
var parse_text_p = function(text, tag) {
	return unescapexml(utf8read(text.replace(/<text:s\/>/g," ").replace(/<[^>]*>/g,"")));
};

var utf8read = function utf8reada(orig) {
	var out = "", i = 0, c = 0, d = 0, e = 0, f = 0, w = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) { out += String.fromCharCode(c); continue; }
		d = orig.charCodeAt(i++);
		if (c>191 && c<224) { out += String.fromCharCode(((c & 31) << 6) | (d & 63)); continue; }
		e = orig.charCodeAt(i++);
		if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
		f = orig.charCodeAt(i++);
		w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63))-65536;
		out += String.fromCharCode(0xD800 + ((w>>>10)&1023));
		out += String.fromCharCode(0xDC00 + (w&1023));
	}
	return out;
};
var parse_content_xml = (function() {

	var number_formats = {
		/* ods name: [short ssf fmt, long ssf fmt] */
		day: ["d", "dd"],
		month: ["m", "mm"],
		year: ["y", "yy"],
		hours: ["h", "hh"],
		minutes: ["m", "mm"],
		seconds: ["s", "ss"],
		"am-pm": ["A/P", "AM/PM"],
		"day-of-week": ["ddd", "dddd"]
	};

	return function pcx(d, opts) {
		var str = xlml_normalize(d);
		var state = [], tmp;
		var tag;
		var NFtag = {name:""}, NF = "", pidx = 0;
		var sheetag;
		var Sheets = {}, SheetNames = [], ws = {};
		var Rn, q;
		var ctag = {value:""};
		var textp = "", textpidx = 0, textptag;
		var R = -1, C = -1, range = {s: {r:1000000,c:10000000}, e: {r:0, c:0}};
		var number_format_map = {};
		var merges = [], mrange = {}, mR = 0, mC = 0;
		var rept = 1;
		xlmlregex.lastIndex = 0;
		while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {

			case 'table': // 9.1.2 <table:table>
				if(Rn[1]==='/') {
					if(range.e.c >= range.s.c && range.e.r >= range.s.r) ws['!ref'] = get_utils().encode_range(range);
					if(merges.length) ws['!merges'] = merges;
					sheetag.name = utf8read(sheetag.name);
					SheetNames.push(sheetag.name);
					Sheets[sheetag.name] = ws;
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					sheetag = parsexmltag(Rn[0], false);
					R = C = -1;
					range.s.r = range.s.c = 10000000; range.e.r = range.e.c = 0;
					ws = {}; merges = [];
				}
				break;

			case 'table-row': // 9.1.3 <table:table-row>
				if(Rn[1] === '/') break;
				++R; C = -1; break;
			case 'covered-table-cell': // 9.1.5 table:covered-table-cell
				++C; break; /* stub */
			case 'table-cell':
				if(Rn[0].charAt(Rn[0].length-2) === '/') {
					ctag = parsexmltag(Rn[0], false);
					if(ctag['number-columns-repeated']) C+= parseInt(ctag['number-columns-repeated'], 10);
					else ++C;
				}
				else if(Rn[1]!=='/') {
					++C;
					rept = 1;
					if(C > range.e.c) range.e.c = C;
					if(R > range.e.r) range.e.r = R;
					if(C < range.s.c) range.s.c = C;
					if(R < range.s.r) range.s.r = R;
					ctag = parsexmltag(Rn[0], false);
					q = {t:ctag['value-type'], v:null};
					if(ctag['number-columns-spanned'] || ctag['number-rows-spanned']) {
						mR = parseInt(ctag['number-rows-spanned'],10) || 0;
						mC = parseInt(ctag['number-columns-spanned'],10) || 0;
						mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
						merges.push(mrange);
					}

					/* 19.675.2 table:number-columns-repeated */
					if(ctag['number-columns-repeated']) rept = parseInt(ctag['number-columns-repeated'], 10);

					/* 19.385 office:value-type */
					switch(q.t) {
						case 'boolean': q.t = 'b'; q.v = parsexmlbool(ctag['boolean-value']); break;
						case 'float': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'percentage': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'currency': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'date': q.t = 'n'; q.v = datenum(ctag['date-value']); q.z = 'm/d/yy'; break;
						case 'time': q.t = 'n'; q.v = parse_isodur(ctag['time-value'])/86400; break;
						default:
							if(q.t === 'string' || !q.t) {
								q.t = 's';
								if(ctag['string-value'] != null) textp = ctag['string-value'];
							} else throw new Error('Unsupported value type ' + q.t);
					}
				} else {
					if(q.t === 's') q.v = textp || '';
					if(textp) q.w = textp;
					if(!(opts.sheetRows && opts.sheetRows < R)) {
						ws[get_utils().encode_cell({r:R,c:C})] = q;
						while(--rept > 0) ws[get_utils().encode_cell({r:R,c:++C})] = dup(q);
						if(range.e.c <= C) range.e.c = C;
					}
					q = {};
					textp = "";
				}
				break; // 9.1.4 <table:table-cell>

			/* pure state */
			case 'document-content': // 3.1.3.2 <office:document-content>
			case 'spreadsheet': // 3.7 <office:spreadsheet>
			case 'scripts': // 3.12 <office:scripts>
			case 'font-face-decls': // 3.14 <office:font-face-decls>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
				break;

			/* ignore state */
			case 'shapes': // 9.2.8 <table:shapes>
			case 'frame': // 10.4.2 <draw:frame>
			case 'text-box': // 10.4.3 <draw:text-box>
			case 'image': // 10.4.4 <draw:image>
			case 'data-pilot-tables': // 9.6.2 <table:data-pilot-tables>
			case 'list-style': // 16.30 <text:list-style>
			case 'form': // 13.13 <form:form>
			case 'dde-links': // 9.8 <table:dde-links>
			case 'annotation': // 14.1 <office:annotation>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], false]);
				break;

			case 'number-style': // 16.27.2 <number:number-style>
			case 'percentage-style': // 16.27.9 <number:percentage-style>
			case 'date-style': // 16.27.10 <number:date-style>
			case 'time-style': // 16.27.18 <number:time-style>
				if(Rn[1]==='/'){
					number_format_map[NFtag.name] = NF;
					if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					NF = "";
					NFtag = parsexmltag(Rn[0], false);
					state.push([Rn[3], true]);
				} break;

			case 'script': break; // 3.13 <office:script>
			case 'automatic-styles': break; // 3.15.3 <office:automatic-styles>

			case 'style': break; // 16.2 <style:style>
			case 'map': break; // 16.3 <style:map>
			case 'font-face': break; // 16.21 <style:font-face>

			case 'paragraph-properties': break; // 17.6 <style:paragraph-properties>
			case 'table-properties': break; // 17.15 <style:table-properties>
			case 'table-column-properties': break; // 17.16 <style:table-column-properties>
			case 'table-row-properties': break; // 17.17 <style:table-row-properties>
			case 'table-cell-properties': break; // 17.18 <style:table-cell-properties>

			case 'number': // 16.27.3 <number:number>
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'fraction': break; // TODO 16.27.6 <number:fraction>

			case 'day': // 16.27.11 <number:day>
			case 'month': // 16.27.12 <number:month>
			case 'year': // 16.27.13 <number:year>
			case 'era': // 16.27.14 <number:era>
			case 'day-of-week': // 16.27.15 <number:day-of-week>
			case 'week-of-year': // 16.27.16 <number:week-of-year>
			case 'quarter': // 16.27.17 <number:quarter>
			case 'hours': // 16.27.19 <number:hours>
			case 'minutes': // 16.27.20 <number:minutes>
			case 'seconds': // 16.27.21 <number:seconds>
			case 'am-pm': // 16.27.22 <number:am-pm>
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'boolean-style': break; // 16.27.23 <number:boolean-style>
			case 'boolean': break; // 16.27.24 <number:boolean>
			case 'text-style': break; // 16.27.25 <number:text-style>
			case 'text': // 16.27.26 <number:text>
				if(Rn[0].substr(-2) === "/>") break;
				else if(Rn[1]==="/") switch(state[state.length-1][0]) {
					case 'number-style':
					case 'date-style':
					case 'time-style':
						NF += str.slice(pidx, Rn.index);
						break;
				}
				else pidx = Rn.index + Rn[0].length;
				break;
			case 'text-content': break; // 16.27.27 <number:text-content>
			case 'text-properties': break; // 16.27.27 <style:text-properties>

			case 'body': break; // 3.3 16.9.6 19.726.3

			case 'forms': break; // 12.25.2 13.2
			case 'table-column': break; // 9.1.6 <table:table-column>

			case 'null-date': break; // 9.4.2 <table:null-date> TODO: date1904

			case 'graphic-properties': break; // 17.21 <style:graphic-properties>
			case 'calculation-settings': break; // 9.4.1 <table:calculation-settings>
			case 'named-expressions': break; // 9.4.11 <table:named-expressions>
			case 'named-range': break; // 9.4.12 <table:named-range>
			case 'named-expression': break; // 9.4.13 <table:named-expression>
			case 'sort': break; // 9.4.19 <table:sort>
			case 'sort-by': break; // 9.4.20 <table:sort-by>
			case 'sort-groups': break; // 9.4.22 <table:sort-groups>

			case 'span': break; // <text:span>
			case 'line-break': break; // 6.1.5 <text:line-break>
			case 'p':
				if(Rn[1]==='/') textp = parse_text_p(str.slice(textpidx,Rn.index), textptag);
				else { textptag = parsexmltag(Rn[0], false); textpidx = Rn.index + Rn[0].length; }
				break; // <text:p>
			case 's': break; // <text:s>
			case 'date': break; // <*:date>

			case 'object': break; // 10.4.6.2 <draw:object>
			case 'title': break; // <*:title>
			case 'desc': break; // <*:desc>

			case 'table-source': break; // 9.2.6

			case 'iteration': break; // 9.4.3 <table:iteration>
			case 'content-validations': break; // 9.4.4 <table:
			case 'content-validation': break; // 9.4.5 <table:
			case 'error-message': break; // 9.4.7 <table:
			case 'database-ranges': break; // 9.4.14 <table:database-ranges>
			case 'database-range': break; // 9.4.15 <table:database-range>
			case 'filter': break; // 9.5.2 <table:filter>
			case 'filter-and': break; // 9.5.3 <table:filter-and>
			case 'filter-or': break; // 9.5.4 <table:filter-or>
			case 'filter-condition': break; // 9.5.5 <table:filter-condition>

			case 'list-level-style-bullet': break; // 16.31 <text:
			case 'list-level-style-number': break; // 16.32 <text:
			case 'list-level-properties': break; // 17.19 <style:

			/* 7.3 Document Fields */
			case 'sender-firstname': // 7.3.6.2
			case 'sender-lastname': // 7.3.6.3
			case 'sender-initials': // 7.3.6.4
			case 'sender-title': // 7.3.6.5
			case 'sender-position': // 7.3.6.6
			case 'sender-email': // 7.3.6.7
			case 'sender-phone-private': // 7.3.6.8
			case 'sender-fax': // 7.3.6.9
			case 'sender-company': // 7.3.6.10
			case 'sender-phone-work': // 7.3.6.11
			case 'sender-street': // 7.3.6.12
			case 'sender-city': // 7.3.6.13
			case 'sender-postal-code': // 7.3.6.14
			case 'sender-country': // 7.3.6.15
			case 'sender-state-or-province': // 7.3.6.16
			case 'author-name': // 7.3.7.1
			case 'author-initials': // 7.3.7.2
			case 'chapter': // 7.3.8
			case 'file-name': // 7.3.9
			case 'template-name': // 7.3.9
			case 'sheet-name': // 7.3.9
				break;

			/* 9.6 Data Pilot Tables <table: */
			case 'data-pilot-table': // 9.6.3
			case 'source-cell-range': // 9.6.5
			case 'source-service': // 9.6.6
			case 'data-pilot-field': // 9.6.7
			case 'data-pilot-level': // 9.6.8
			case 'data-pilot-subtotals': // 9.6.9
			case 'data-pilot-subtotal': // 9.6.10
			case 'data-pilot-members': // 9.6.11
			case 'data-pilot-member': // 9.6.12
			case 'data-pilot-display-info': // 9.6.13
			case 'data-pilot-sort-info': // 9.6.14
			case 'data-pilot-layout-info': // 9.6.15
			case 'data-pilot-field-reference': // 9.6.16
			case 'data-pilot-groups': // 9.6.17
			case 'data-pilot-group': // 9.6.18
			case 'data-pilot-group-member': // 9.6.19
				break;

			/* 10.3 Drawing Shapes */
			case 'rect': // 10.3.2
				break;

			/* 14.6 DDE Connections */
			case 'dde-connection-decls': // 14.6.2 <text:
			case 'dde-connection-decl': // 14.6.3 <text:
			case 'dde-link': // 14.6.4 <table:
			case 'dde-source': // 14.6.5 <office:
				break;

			case 'properties': break; // 13.7 <form:properties>
			case 'property': break; // 13.8 <form:property>

			case 'a': break; // 6.1.8 hyperlink

			/* non-standard */
			case 'table-protection': break;
			case 'data-pilot-grand-total': break; // <table:
			default:
				if(Rn[2] === 'dc:') break; // TODO: properties
				if(Rn[2] === 'draw:') break; // TODO: drawing
				if(Rn[2] === 'calcext:') break; // ignore undocumented extensions
				if(opts.WTF) throw Rn;
		}
		var out = {
			Sheets: Sheets,
			SheetNames: SheetNames
		};
		return out;
	};
})();
var write_content_xml = (function() {
	var null_cell_xml = '          <table:table-cell />\n';
	var write_ws = function(ws, wb, i, opts) {
		/* Section 9 Tables */
		var o = [];
		o.push('      <table:table table:name="' + escapexml(wb.SheetNames[i]) + '">\n');
		var R=0,C=0, range = get_utils().decode_range(ws['!ref']);
		for(R = 0; R < range.s.r; ++R) o.push('        <table:table-row></table:table-row>\n');
		for(; R <= range.e.r; ++R) {
			o.push('        <table:table-row>\n');
			for(C=0; C < range.s.c; ++C) o.push(null_cell_xml);
			for(; C <= range.e.c; ++C) {
				var ref = get_utils().encode_cell({r:R, c:C}), cell = ws[ref];
				if(cell) switch(cell.t) {
					case 'b': o.push('          <table:table-cell office:value-type="boolean" office:boolean-value="' + (cell.v ? 'true' : 'false') + '"><text:p>' + (cell.v ? 'TRUE' : 'FALSE') + '</text:p></table:table-cell>\n'); break;
					case 'n': o.push('          <table:table-cell office:value-type="float" office:value="' + cell.v + '"><text:p>' + (cell.w||cell.v) + '</text:p></table:table-cell>\n'); break;
					case 's': case 'str': o.push('          <table:table-cell office:value-type="string"><text:p>' + escapexml(cell.v) + '</text:p></table:table-cell>\n'); break;
					//case 'd': // TODO
					//case 'e':
					default: o.push(null_cell_xml);
				} else o.push(null_cell_xml);
			}
			o.push('        </table:table-row>\n');
		}
		o.push('      </table:table>\n');
		return o.join("");
	};

	return function wcx(wb, opts) {
		var o = [XML_HEADER];
		/* 3.1.3.2 */
		o.push('<office:document-content office:version="1.2" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">\n'); // TODO
		o.push('  <office:body>\n');
		o.push('    <office:spreadsheet>\n');
		for(var i = 0; i != wb.SheetNames.length; ++i) o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts));
		o.push('    </office:spreadsheet>\n');
		o.push('  </office:body>\n');
		o.push('</office:document-content>');
		return o.join("");
	};
})();
/* Part 3: Packages */
function parse_ods(zip, opts) {
	opts = opts || ({});
	var manifest = parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts);
	return parse_content_xml(getzipdata(zip, 'content.xml'), opts);
}
function write_ods(wb, opts) {
var zip = new jszip();
	var f = "";

	var manifest = [];
	var rdf = [];

	/* 3:3.3 and 2:2.2.4 */
	f = "mimetype";
	zip.file(f, "application/vnd.oasis.opendocument.spreadsheet");

	/* Part 2 Section 2.2 Documents */
	f = "content.xml";
	zip.file(f, write_content_xml(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "ContentFile"]);

	/* Part 3 Section 6 Metadata Manifest File */
	f = "manifest.rdf";
	zip.file(f, write_rdf(rdf, opts));
	manifest.push([f, "application/rdf+xml"]);

	/* Part 3 Section 4 Manifest File */
	f = "META-INF/manifest.xml";
	zip.file(f, write_manifest(manifest, opts));

	return zip;
}
ODS.parse_ods = parse_ods;
ODS.write_ods = write_ods;
})(typeof exports !== 'undefined' ? exports : ODS);
