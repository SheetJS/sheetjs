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
