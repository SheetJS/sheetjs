var RTF = (function() {
	function rtf_to_sheet(d/*:RawData*/, opts)/*:Worksheet*/ {
		switch(opts.type) {
			case 'base64': return rtf_to_sheet_str(Base64.decode(d), opts);
			case 'binary': return rtf_to_sheet_str(d, opts);
			case 'buffer': return rtf_to_sheet_str(d.toString('binary'), opts);
			case 'array':  return rtf_to_sheet_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}

	function rtf_to_sheet_str(str/*:string*/, opts)/*:Worksheet*/ {
		throw new Error("Unsupported RTF");
	}

	function rtf_to_workbook(d/*:RawData*/, opts)/*:Workbook*/ { return sheet_to_workbook(rtf_to_sheet(d, opts), opts); }
	function sheet_to_rtf() { throw new Error("Unsupported"); }

	return {
		to_workbook: rtf_to_workbook,
		to_sheet: rtf_to_sheet,
		from_sheet: sheet_to_rtf
	};
})();
