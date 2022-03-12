/* Open Document Format for Office Applications (OpenDocument) Version 1.2 */
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

function write_manifest(manifest/*:Array<Array<string> >*/)/*:string*/ {
	var o = [XML_HEADER];
	o.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
	o.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
	for(var i = 0; i < manifest.length; ++i) o.push('  <manifest:file-entry manifest:full-path="' + manifest[i][0] + '" manifest:media-type="' + manifest[i][1] + '"/>\n');
	o.push('</manifest:manifest>');
	return o.join("");
}

/* Part 3 Section 6 Metadata Manifest File */
function write_rdf_type(file/*:string*/, res/*:string*/, tag/*:?string*/) {
	return [
		'  <rdf:Description rdf:about="' + file + '">\n',
		'    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/' + (tag || "odf") + '#' + res + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf_has(base/*:string*/, file/*:string*/) {
	return [
		'  <rdf:Description rdf:about="' + base + '">\n',
		'    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="' + file + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf(rdf) {
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
/* TODO: pull properties */
function write_meta_ods(/*:: wb: Workbook, opts: any*/)/*:string*/ {
	return '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2"><office:meta><meta:generator>Sheet' + 'JS ' + XLSX.version + '</meta:generator></office:meta></office:document-meta>';
}

