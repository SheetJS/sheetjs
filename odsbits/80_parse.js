/* Part 3: Packages */
function parse_ods(zip/*:ZIPFile*/, opts/*:?ParseOpts*/) {
	opts = opts || ({}/*:any*/);
	var ods = !!safegetzipfile(zip, 'objectdata');
	if(ods) var manifest = parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts);
	var content = getzipdata(zip, 'content.xml');
	return parse_content_xml(ods ? content : utf8read(content), opts);
}

function parse_fods(data/*:string*/, opts/*:?ParseOpts*/) {
	return parse_content_xml(data, opts);
}
