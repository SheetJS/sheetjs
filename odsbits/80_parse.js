/* Part 3: Packages */
var parse_ods = function(zip, opts) {
	//var manifest = parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'));
	return parse_content_xml(getzipdata(zip, 'content.xml'), opts);
};
