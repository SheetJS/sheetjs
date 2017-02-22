function write_ods(wb/*:any*/, opts/*:any*/) {
	if(opts.bookType == "fods") return write_content_xml(wb, opts);

	/*:: if(!jszip) throw new Error("JSZip is not available"); */
	var zip = new jszip();
	var f = "";

	var manifest/*:Array<Array<string> >*/ = [];
	var rdf = [];

	/* 3:3.3 and 2:2.2.4 */
	f = "mimetype";
	zip.file(f, "application/vnd.oasis.opendocument.spreadsheet");

	/* Part 1 Section 2.2 Documents */
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
