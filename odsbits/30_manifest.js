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

function write_manifest(manifest/*:Array<Array<string> >*/, opts)/*:string*/ {
	var o = [XML_HEADER];
	o.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
	o.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
	for(var i = 0; i < manifest.length; ++i) o.push('  <manifest:file-entry manifest:full-path="' + manifest[i][0] + '" manifest:media-type="' + manifest[i][1] + '"/>\n');
	o.push('</manifest:manifest>');
	return o.join("");
}
