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
