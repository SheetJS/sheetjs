/* OpenDocument */
var write_styles_ods/*:{(wb:any, opts:any):string}*/ = /* @__PURE__ */(function() {
	var master_styles = [
		'<office:master-styles>',
			'<style:master-page style:name="mp1" style:page-layout-name="mp1">',
				'<style:header/>',
				'<style:header-left style:display="false"/>',
				'<style:footer/>',
				'<style:footer-left style:display="false"/>',
			'</style:master-page>',
		'</office:master-styles>'
	].join("");

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
	}) + '>' + master_styles + '</office:document-styles>';

	return function wso(/*::wb, opts*/) {
		return XML_HEADER + payload;
	};
})();

// TODO: find out if anyone actually read the spec.  LO has some wild errors
function write_number_format_ods(nf/*:string*/, nfidx/*:string*/)/*:string*/ {
	var type = "number", payload = "", nopts = { "style:name": nfidx }, c = "", i = 0;
	nf = nf.replace(/"[$]"/g, "$");
	/* TODO: replace with an actual parser based on a real grammar */
	j: {
		// TODO: support style maps
		if(nf.indexOf(";") > -1) {
			console.error("Unsupported ODS Style Map exported.  Using first branch of " + nf);
			nf = nf.slice(0, nf.indexOf(";"));
		}

		if(nf == "@") { type = "text"; payload = "<number:text-content/>"; break j; }

		/* currency flag */
		if(nf.indexOf(/\$/) > -1) { type = "currency"; }

		/* opening string literal */
		if(nf[i] == '"') {
			c = "";
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			if(nf[i+1] == "*") {
				i++;
				payload += '<number:fill-character>' + escapexml(c.replace(/""/g, '"')) + '</number:fill-character>';
			} else {
				payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
			}
			nf = nf.slice(i+1); i = 0;
		}

		/* fractions */
		var t = nf.match(/# (\?+)\/(\?+)/);
		if(t) { payload += writextag("number:fraction", null, {"number:min-integer-digits":0, "number:min-numerator-digits": t[1].length, "number:max-denominator-value": Math.max(+(t[1].replace(/./g, "9")), +(t[2].replace(/./g, "9"))) }); break j; }
		if((t=nf.match(/# (\?+)\/(\d+)/))) { payload += writextag("number:fraction", null, {"number:min-integer-digits":0, "number:min-numerator-digits": t[1].length, "number:denominator-value": +t[2]}); break j; }

		/* percentages */
		if((t=nf.match(/(\d+)(|\.\d+)%/))) { type = "percentage"; payload += writextag("number:number", null, {"number:decimal-places": t[2] && t.length - 1 || 0, "number:min-decimal-places": t[2] && t.length - 1 || 0, "number:min-integer-digits": t[1].length }) + "<number:text>%</number:text>"; break j; }

		/* datetime */
		var has_time = false;
		if(["y","m","d"].indexOf(nf[0]) > -1) {
			type = "date";
			k: for(; i < nf.length; ++i) switch((c = nf[i].toLowerCase())) {
				case "h": case "s": has_time = true; --i; break k;
				case "m":
					l: for(var h = i+1; h < nf.length; ++h) switch(nf[h]) {
						case "y": case "d": break l;
						case "h": case "s": has_time = true; --i; break k;
					}
					/* falls through */
				case "y": case "d":
					while((nf[++i]||"").toLowerCase() == c[0]) c += c[0]; --i;
					switch(c) {
						case "y": case "yy": payload += "<number:year/>"; break;
						case "yyy": case "yyyy": payload += '<number:year number:style="long"/>'; break;
						case "mmmmm": console.error("ODS has no equivalent of format |mmmmm|");
							/* falls through */
						case "m": case "mm": case "mmm": case "mmmm":
							payload += '<number:month number:style="' + (c.length % 2 ? "short" : "long") + '" number:textual="' + (c.length >= 3 ? "true" : "false") + '"/>';
							break;
						case "d": case "dd": payload += '<number:day number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "ddd": case "dddd": payload += '<number:day-of-week number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
					}
					break;
				case '"':
					while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
					payload += '<number:text>' + escapexml(c.slice(1).replace(/""/g, '"')) + '</number:text>';
					break;
				case '/': payload += '<number:text>' + escapexml(c) + '</number:text>'; break;
				default: console.error("unrecognized character " + c + " in ODF format " + nf);
			}
			if(!has_time) break j;
			nf = nf.slice(i+1); i = 0;
		}
		if(nf.match(/^\[?[hms]/)) {
			if(type == "number") type = "time";
			if(nf.match(/\[/)) {
				nf = nf.replace(/[\[\]]/g, "");
				nopts['number:truncate-on-overflow'] = "false";
			}
			for(; i < nf.length; ++i) switch((c = nf[i].toLowerCase())) {
				case "h": case "m": case "s":
					while((nf[++i]||"").toLowerCase() == c[0]) c += c[0]; --i;
					switch(c) {
						case "h": case "hh": payload += '<number:hours number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "m": case "mm": payload += '<number:minutes number:style="' + (c.length % 2 ? "short" : "long") + '"/>'; break;
						case "s": case "ss":
							if(nf[i+1] == ".") do { c += nf[i+1]; ++i; } while(nf[i+1] == "0");
							payload += '<number:seconds number:style="' + (c.match("ss") ? "long" : "short") + '"' + (c.match(/\./) ? ' number:decimal-places="' + (c.match(/0+/)||[""])[0].length + '"' : "")+ '/>'; break;
					}
					break;
				case '"':
					while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
					payload += '<number:text>' + escapexml(c.slice(1).replace(/""/g, '"')) + '</number:text>';
					break;
				case '/': payload += '<number:text>' + escapexml(c) + '</number:text>'; break;
				case "a":
					if(nf.slice(i, i+3).toLowerCase() == "a/p") { payload += '<number:am-pm/>'; i += 2; break; } // Note: ODF does not support A/P
					if(nf.slice(i, i+5).toLowerCase() == "am/pm")  { payload += '<number:am-pm/>'; i += 4; break; }
					/* falls through */
				default: console.error("unrecognized character " + c + " in ODF format " + nf);
			}
			break j;
		}

		/* currency flag */
		if(nf.indexOf(/\$/) > -1) { type = "currency"; }

		/* should be in a char loop */
		if(nf[0] == "$") { payload += '<number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>'; nf = nf.slice(1); i = 0; }
		i = 0; if(nf[i] == '"') {
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			if(nf[i+1] == "*") {
				i++;
				payload += '<number:fill-character>' + escapexml(c.replace(/""/g, '"')) + '</number:fill-character>';
			} else {
				payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
			}
			nf = nf.slice(i+1); i = 0;
		}

		/* number TODO: interstitial text e.g. 000)000-0000 */
		var np = nf.match(/([#0][0#,]*)(\.[0#]*|)(E[+]?0*|)/i);
		if(!np || !np[0]) console.error("Could not find numeric part of " + nf);
		else {
			var base = np[1].replace(/,/g, "");
			payload += '<number:' + (np[3] ? "scientific-" : "")+ 'number' +
				' number:min-integer-digits="' + (base.indexOf("0") == -1 ? "0" : base.length - base.indexOf("0")) + '"' +
				(np[0].indexOf(",") > -1 ? ' number:grouping="true"' : "") +
				(np[2] && ' number:decimal-places="' + (np[2].length - 1) + '"' || ' number:decimal-places="0"') +
				(np[3] && np[3].indexOf("+") > -1 ? ' number:forced-exponent-sign="true"' : "" ) +
				(np[3] ? ' number:min-exponent-digits="' + np[3].match(/0+/)[0].length + '"' : "" ) +
				'>' +
				/* TODO: interstitial text placeholders */
				'</number:' + (np[3] ? "scientific-" : "") + 'number>';
			i = np.index + np[0].length;
		}

		/* residual text */
		if(nf[i] == '"') {
			c = "";
			while(nf[++i] != '"' || nf[++i] == '"') c += nf[i]; --i;
			payload += '<number:text>' + escapexml(c.replace(/""/g, '"')) + '</number:text>';
		}
	}

	if(!payload) { console.error("Could not generate ODS number format for |" + nf + "|"); return ""; }
	return writextag("number:" + type + "-style", payload, nopts);
}

function write_names_ods(Names, SheetNames, idx) {
	var scoped = Names.filter(function(name) { return name.Sheet == (idx == -1 ? null : idx); });
	if(!scoped.length) return "";
	return "      <table:named-expressions>\n" + scoped.map(function(name) {
		var odsref =  csf_to_ods_3D(name.Ref);
		return "        " + writextag("table:named-range", null, {
			"table:name": name.Name,
			"table:cell-range-address": odsref,
			"table:base-cell-address": odsref.replace(/[\.]?[^\.]*$/, ".$A$1")
		});
	}).join("\n") + "\n      </table:named-expressions>\n";
}
var write_content_ods/*:{(wb:any, opts:any):string}*/ = /* @__PURE__ */(function() {
	/* 6.1.2 White Space Characters */
	var write_text_p = function(text/*:string*/)/*:string*/ {
		return escapexml(text)
			.replace(/  +/g, function($$){return '<text:s text:c="'+$$.length+'"/>';})
			.replace(/\t/g, "<text:tab/>")
			.replace(/\n/g, "</text:p><text:p>")
			.replace(/^ /, "<text:s/>").replace(/ $/, "<text:s/>");
	};

	var null_cell_xml = '          <table:table-cell />\n';
	var covered_cell_xml = '          <table:covered-table-cell/>\n';
	var write_ws = function(ws, wb/*:Workbook*/, i/*:number*/, opts, nfs)/*:string*/ {
		/* Section 9 Tables */
		var o/*:Array<string>*/ = [];
		o.push('      <table:table table:name="' + escapexml(wb.SheetNames[i]) + '" table:style-name="ta1">\n');
		var R=0,C=0, range = decode_range(ws['!ref']||"A1");
		var marr/*:Array<Range>*/ = ws['!merges'] || [], mi = 0;
		var dense = Array.isArray(ws);
		if(ws["!cols"]) {
			for(C = 0; C <= range.e.c; ++C) o.push('        <table:table-column' + (ws["!cols"][C] ? ' table:style-name="co' + ws["!cols"][C].ods + '"' : '') + '></table:table-column>\n');
		}
		var H = "", ROWS = ws["!rows"]||[];
		for(R = 0; R < range.s.r; ++R) {
			H = ROWS[R] ? ' table:style-name="ro' + ROWS[R].ods + '"' : "";
			o.push('        <table:table-row' + H + '></table:table-row>\n');
		}
		for(; R <= range.e.r; ++R) {
			H = ROWS[R] ? ' table:style-name="ro' + ROWS[R].ods + '"' : "";
			o.push('        <table:table-row' + H + '>\n');
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
						textp = cell.v == null ? "" : cell.v;
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
					var _tgt = cell.l.Target;
					_tgt = _tgt.charAt(0) == "#" ? "#" + csf_to_ods_3D(_tgt.slice(1)) : _tgt;
					// TODO: choose correct parent path format based on link delimiters
					if(_tgt.charAt(0) != "#" && !_tgt.match(/^\w+:/)) _tgt = '../' + _tgt;
					text_p = writextag('text:a', text_p, {'xlink:href': _tgt.replace(/&/g, "&amp;")});
				}
				if(nfs[cell.z]) ct["table:style-name"] = "ce" + nfs[cell.z].slice(1);
				o.push('          ' + writextag('table:table-cell', writextag('text:p', text_p, {}), ct) + '\n');
			}
			o.push('        </table:table-row>\n');
		}
		if((wb.Workbook||{}).Names) o.push(write_names_ods(wb.Workbook.Names, wb.SheetNames, i));
		o.push('      </table:table>\n');
		return o.join("");
	};

	var write_automatic_styles_ods = function(o/*:Array<string>*/, wb) {
		o.push(' <office:automatic-styles>\n');

		/* column styles */
		var cidx = 0;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			if(ws["!cols"]) {
				for(var C = 0; C < ws["!cols"].length; ++C) if(ws["!cols"][C]) {
					var colobj = ws["!cols"][C];
					if(colobj.width == null && colobj.wpx == null && colobj.wch == null) continue;
					process_col(colobj);
					colobj.ods = cidx;
					var w = ws["!cols"][C].wpx + "px";
					o.push('  <style:style style:name="co' + cidx + '" style:family="table-column">\n');
					o.push('   <style:table-column-properties fo:break-before="auto" style:column-width="' + w + '"/>\n');
					o.push('  </style:style>\n');
					++cidx;
				}
			}
		});

		/* row styles */
		var ridx = 0;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			if(ws["!rows"]) {
				for(var R = 0; R < ws["!rows"].length; ++R) if(ws["!rows"][R]) {
					ws["!rows"][R].ods = ridx;
					var h = ws["!rows"][R].hpx + "px";
					o.push('  <style:style style:name="ro' + ridx + '" style:family="table-row">\n');
					o.push('   <style:table-row-properties fo:break-before="auto" style:row-height="' + h + '"/>\n');
					o.push('  </style:style>\n');
					++ridx;
				}
			}
		});

		/* table */
		o.push('  <style:style style:name="ta1" style:family="table" style:master-page-name="mp1">\n');
		o.push('   <style:table-properties table:display="true" style:writing-mode="lr-tb"/>\n');
		o.push('  </style:style>\n');

		o.push('  <number:date-style style:name="N37" number:automatic-order="true">\n');
		o.push('   <number:month number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:day number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:year/>\n');
		o.push('  </number:date-style>\n');

		/* number formats, table cells, text */
		var nfs = {};
		var nfi = 69;
		wb.SheetNames.map(function(n) { return wb.Sheets[n]; }).forEach(function(ws) {
			if(!ws) return;
			var range = decode_range(ws["!ref"]);
			for(var R = 0; R <= range.e.r; ++R) for(var C = 0; C <= range.e.c; ++C) {
				var c = Array.isArray(ws) ? (ws[R]||[])[C] : ws[encode_cell({r:R,c:C})];
				if(!c || !c.z || c.z.toLowerCase() == "general") continue;
				if(!nfs[c.z]) {
					var out = write_number_format_ods(c.z, "N" + nfi);
					if(out) { nfs[c.z] = "N" + nfi; ++nfi; o.push(out + "\n"); }
				}
			}
		});
		o.push('  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>\n');
		keys(nfs).forEach(function(nf) {
			o.push('<style:style style:name="ce' + nfs[nf].slice(1) + '" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="' + nfs[nf] + '"/>\n');
		});

		/* page-layout */

		o.push(' </office:automatic-styles>\n');
		return nfs;
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

		if(opts.bookType == "fods") {
			o.push('<office:document' + attr + fods + '>\n');
			o.push(write_meta_ods().replace(/<office:document-meta.*?>/, "").replace(/<\/office:document-meta>/, "") + "\n");
			// TODO: settings (equiv of settings.xml for ODS)
		} else o.push('<office:document-content' + attr  + '>\n');
		// o.push('  <office:scripts/>\n');
		var nfs = write_automatic_styles_ods(o, wb);
		o.push('  <office:body>\n');
		o.push('    <office:spreadsheet>\n');
		if(((wb.Workbook||{}).WBProps||{}).date1904) o.push('      <table:calculation-settings table:case-sensitive="false" table:search-criteria-must-apply-to-whole-cell="true" table:use-wildcards="true" table:use-regular-expressions="false" table:automatic-find-labels="false">\n        <table:null-date table:date-value="1904-01-01"/>\n      </table:calculation-settings>\n');
		for(var i = 0; i != wb.SheetNames.length; ++i) o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts, nfs));
		if((wb.Workbook||{}).Names) o.push(write_names_ods(wb.Workbook.Names, wb.SheetNames, -1));
		o.push('    </office:spreadsheet>\n');
		o.push('  </office:body>\n');
		if(opts.bookType == "fods") o.push('</office:document>');
		else o.push('</office:document-content>');
		return o.join("");
	};
})();

function write_ods(wb/*:any*/, opts/*:any*/) {
	if(opts.bookType == "fods") return write_content_ods(wb, opts);

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
	zip_add_file(zip, f, XML_HEADER + write_meta_ods(/*::wb, opts*/));
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

