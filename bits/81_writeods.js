/* OpenDocument */
var write_styles_ods/*:{(wb:any, opts:any):string}*/ = (function() {
	var payload = '<office:document-styles ' + wxt_helper({
		'xmlns:office':   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		'xmlns:table':    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		'xmlns:style':    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		'xmlns:text':     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		'xmlns:draw':     "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
		'xmlns:fo':       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		'xmlns:xlink':    "http://www.w3.org/1999/xlink",
		'xmlns:dc':       "http://purl.org/dc/elements/1.1/",
		'xmlns:number':   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		'xmlns:svg':      "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		'xmlns:of':       "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		'office:version': "1.2"
	}) + '></office:document-styles>';
	return function wso(/*::wb, opts*/) {
		return XML_HEADER + payload;
	};
})();
var write_content_ods/*:{(wb:any, opts:any):string}*/ = (function() {
	/* 6.1.2 White Space Characters */
	var write_text_p = function(text/*:string*/)/*:string*/ {
		return escapexml(text)
			.replace(/  +/g, function($$){return '<text:s text:c="'+$$.length+'"/>';})
			.replace(/\t/g, "<text:tab/>")
			.replace(/\n/g, "<text:line-break/>")
			.replace(/^ /, "<text:s/>").replace(/ $/, "<text:s/>");
	};

	var null_cell_xml = '          <table:table-cell />\n';
	var covered_cell_xml = '          <table:covered-table-cell/>\n';
	var write_ws = function(ws, wb/*:Workbook*/, i/*:number*//*::, opts*/)/*:string*/ {
		/* Section 9 Tables */
		var o/*:Array<string>*/ = [];
		o.push('      <table:table table:name="' + escapexml(wb.SheetNames[i]) + '">\n');
		var R=0,C=0, range = decode_range(ws['!ref']);
		var marr/*:Array<Range>*/ = ws['!merges'] || [], mi = 0;
		var dense = Array.isArray(ws);
		for(R = 0; R < range.s.r; ++R) o.push('        <table:table-row></table:table-row>\n');
		for(; R <= range.e.r; ++R) {
			o.push('        <table:table-row>\n');
			for(C=0; C < range.s.c; ++C) o.push(null_cell_xml);
			for(; C <= range.e.c; ++C) {
				var skip = false, ct = {}, textp = "";
				for(mi = 0; mi != marr.length; ++mi) {
					if(marr[mi].s.c > C) continue;
					if(marr[mi].s.r > R) continue;
					if(marr[mi].e.c < C) continue;
					if(marr[mi].e.r < R) continue;
					if(marr[mi].s.c != C || marr[mi].s.r != R) skip = true;
					ct['table:number-columns-spanned'] = (marr[mi].e.c - marr[mi].s.c + 1);
					ct['table:number-rows-spanned'] =    (marr[mi].e.r - marr[mi].s.r + 1);
					break;
				}
				if(skip) { o.push(covered_cell_xml); continue; }
				var ref = encode_cell({r:R, c:C}), cell = dense ? (ws[R]||[])[C]: ws[ref];
				if(cell && cell.f) {
					ct['table:formula'] = escapexml(csf_to_ods_formula(cell.f));
					if(cell.F) {
						if(cell.F.slice(0, ref.length) == ref) {
							var _Fref = decode_range(cell.F);
							ct['table:number-matrix-columns-spanned'] = (_Fref.e.c - _Fref.s.c + 1);
							ct['table:number-matrix-rows-spanned'] =    (_Fref.e.r - _Fref.s.r + 1);
						}
					}
				}
				if(!cell) { o.push(null_cell_xml); continue; }
				switch(cell.t) {
					case 'b':
						textp = (cell.v ? 'TRUE' : 'FALSE');
						ct['office:value-type'] = "boolean";
						ct['office:boolean-value'] = (cell.v ? 'true' : 'false');
						break;
					case 'n':
						textp = (cell.w||String(cell.v||0));
						ct['office:value-type'] = "float";
						ct['office:value'] = (cell.v||0);
						break;
					case 's': case 'str':
						textp = cell.v;
						ct['office:value-type'] = "string";
						break;
					case 'd':
						textp = (cell.w||(parseDate(cell.v).toISOString()));
						ct['office:value-type'] = "date";
						ct['office:date-value'] = (parseDate(cell.v).toISOString());
						ct['table:style-name'] = "ce1";
						break;
					//case 'e':
					default: o.push(null_cell_xml); continue;
				}
				var text_p = write_text_p(textp);
				if(cell.l && cell.l.Target) {
					var _tgt = cell.l.Target; _tgt = _tgt.charAt(0) == "#" ? "#" + csf_to_ods_3D(_tgt.slice(1)) : _tgt;
					text_p = writextag('text:a', text_p, {'xlink:href': _tgt});
				}
				o.push('          ' + writextag('table:table-cell', writextag('text:p', text_p, {}), ct) + '\n');
			}
			o.push('        </table:table-row>\n');
		}
		o.push('      </table:table>\n');
		return o.join("");
	};

	var write_automatic_styles_ods = function(o/*:Array<string>*/) {
		o.push(' <office:automatic-styles>\n');
		o.push('  <number:date-style style:name="N37" number:automatic-order="true">\n');
		o.push('   <number:month number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:day number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:year/>\n');
		o.push('  </number:date-style>\n');

		o.push('  <style:style style:name="ta1" style:family="table">\n'); // style:master-page-name="mp1">\n');
		o.push('   <style:table-properties table:display="true" style:writing-mode="lr-tb"/>\n');
		o.push('  </style:style>\n');

		o.push('  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>\n');
		o.push(' </office:automatic-styles>\n');
	};

	return function wcx(wb, opts) {
		var o = [XML_HEADER];
		/* 3.1.3.2 */
		var attr = wxt_helper({
			'xmlns:office':       "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
			'xmlns:table':        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
			'xmlns:style':        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
			'xmlns:text':         "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
			'xmlns:draw':         "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
			'xmlns:fo':           "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
			'xmlns:xlink':        "http://www.w3.org/1999/xlink",
			'xmlns:dc':           "http://purl.org/dc/elements/1.1/",
			'xmlns:meta':         "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
			'xmlns:number':       "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
			'xmlns:presentation': "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
			'xmlns:svg':          "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
			'xmlns:chart':        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
			'xmlns:dr3d':         "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
			'xmlns:math':         "http://www.w3.org/1998/Math/MathML",
			'xmlns:form':         "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
			'xmlns:script':       "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
			'xmlns:ooo':          "http://openoffice.org/2004/office",
			'xmlns:ooow':         "http://openoffice.org/2004/writer",
			'xmlns:oooc':         "http://openoffice.org/2004/calc",
			'xmlns:dom':          "http://www.w3.org/2001/xml-events",
			'xmlns:xforms':       "http://www.w3.org/2002/xforms",
			'xmlns:xsd':          "http://www.w3.org/2001/XMLSchema",
			'xmlns:xsi':          "http://www.w3.org/2001/XMLSchema-instance",
			'xmlns:sheet':        "urn:oasis:names:tc:opendocument:sh33tjs:1.0",
			'xmlns:rpt':          "http://openoffice.org/2005/report",
			'xmlns:of':           "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
			'xmlns:xhtml':        "http://www.w3.org/1999/xhtml",
			'xmlns:grddl':        "http://www.w3.org/2003/g/data-view#",
			'xmlns:tableooo':     "http://openoffice.org/2009/table",
			'xmlns:drawooo':      "http://openoffice.org/2010/draw",
			'xmlns:calcext':      "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
			'xmlns:loext':        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
			'xmlns:field':        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
			'xmlns:formx':        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
			'xmlns:css3t':        "http://www.w3.org/TR/css3-text/",
			'office:version':     "1.2"
		});

		var fods = wxt_helper({
			'xmlns:config':    "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
			'office:mimetype': "application/vnd.oasis.opendocument.spreadsheet"
		});

		if(opts.bookType == "fods") o.push('<office:document' + attr + fods + '>\n');
		else o.push('<office:document-content' + attr  + '>\n');
		write_automatic_styles_ods(o);
		o.push('  <office:body>\n');
		o.push('    <office:spreadsheet>\n');
		for(var i = 0; i != wb.SheetNames.length; ++i) o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts));
		o.push('    </office:spreadsheet>\n');
		o.push('  </office:body>\n');
		if(opts.bookType == "fods") o.push('</office:document>');
		else o.push('</office:document-content>');
		return o.join("");
	};
})();

function write_ods(wb/*:any*/, opts/*:any*/) {
	if(opts.bookType == "fods") return write_content_ods(wb, opts);

	/*:: if(!jszip) throw new Error("JSZip is not available"); */
	var zip = zip_new();
	var f = "";

	var manifest/*:Array<Array<string> >*/ = [];
	var rdf/*:Array<[string, string]>*/ = [];

	/* Part 3 Section 3.3 MIME Media Type */
	f = "mimetype";
	zip_add_file(zip, f, "application/vnd.oasis.opendocument.spreadsheet");

	/* Part 1 Section 2.2 Documents */
	f = "content.xml";
	zip_add_file(zip, f, write_content_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "ContentFile"]);

	/* TODO: these are hard-coded styles to satiate excel */
	f = "styles.xml";
	zip_add_file(zip, f, write_styles_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "StylesFile"]);

	/* TODO: this is hard-coded to satiate excel */
	f = "meta.xml";
	zip_add_file(zip, f, write_meta_ods(/*::wb, opts*/));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "MetadataFile"]);

	/* Part 3 Section 6 Metadata Manifest File */
	f = "manifest.rdf";
	zip_add_file(zip, f, write_rdf(rdf/*, opts*/));
	manifest.push([f, "application/rdf+xml"]);

	/* Part 3 Section 4 Manifest File */
	f = "META-INF/manifest.xml";
	zip_add_file(zip, f, write_manifest(manifest/*, opts*/));

	return zip;
}

