/* 18.14 Supplementary Workbook Data */
function parse_xlink_xml(/*::data, rel, name:string, _opts*/) {
	//var opts = _opts || {};
	//if(opts.WTF) throw "XLSX External Link";
}

/* [MS-XLSB] 2.1.7.25 External Link */
function parse_xlink_bin(data, rel, name/*:string*/, _opts) {
	if(!data) return data;
	var opts = _opts || {};

	var pass = false, end = false;

	recordhopper(data, function xlink_parse(val, R, RT) {
		if(end) return;
		switch(RT) {
			case 0x0167: /* 'BrtSupTabs' */
			case 0x016B: /* 'BrtExternTableStart' */
			case 0x016C: /* 'BrtExternTableEnd' */
			case 0x016E: /* 'BrtExternRowHdr' */
			case 0x016F: /* 'BrtExternCellBlank' */
			case 0x0170: /* 'BrtExternCellReal' */
			case 0x0171: /* 'BrtExternCellBool' */
			case 0x0172: /* 'BrtExternCellError' */
			case 0x0173: /* 'BrtExternCellString' */
			case 0x01D8: /* 'BrtExternValueMeta' */
			case 0x0241: /* 'BrtSupNameStart' */
			case 0x0242: /* 'BrtSupNameValueStart' */
			case 0x0243: /* 'BrtSupNameValueEnd' */
			case 0x0244: /* 'BrtSupNameNum' */
			case 0x0245: /* 'BrtSupNameErr' */
			case 0x0246: /* 'BrtSupNameSt' */
			case 0x0247: /* 'BrtSupNameNil' */
			case 0x0248: /* 'BrtSupNameBool' */
			case 0x0249: /* 'BrtSupNameFmla' */
			case 0x024A: /* 'BrtSupNameBits' */
			case 0x024B: /* 'BrtSupNameEnd' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;

			default:
				if(R.T){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record 0x" + RT.toString(16));
		}
	}, opts);
}
