/* Part 3: Packages */
function parse_ods(zip/*:ZIPFile*/, opts/*:?ParseOpts*/) {
	opts = opts || ({}/*:any*/);
	var manifest = parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts);
	return parse_content_xml(getzipdata(zip, 'content.xml'), opts);
}
