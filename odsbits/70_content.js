var write_content_xml/*:{(wb:any, opts:any):string}*/ = (function() {
	var null_cell_xml = '          <table:table-cell />\n';
	var write_ws = function(ws, wb, i/*:number*/, opts)/*:string*/ {
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
		if(opts.bookType == "fods") o.push('<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">');
		else o.push('<office:document-content office:version="1.2" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">\n'); // TODO
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
