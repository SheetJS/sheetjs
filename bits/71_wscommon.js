var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

function get_sst_id(sst, str) {
	for(var i = 0; i != sst.length; ++i) if(sst[i].t === str) { sst.Count ++; return i; }
	sst[sst.length] = {t:str}; sst.Count ++; sst.Unique ++; return sst.length-1;
}

function get_cell_style(styles, cell, opts) {
	var z = opts.revssf[cell.z||"General"];
	for(var i = 0; i != styles.length; ++i) if(styles[i].numFmtId === z) return i;
	styles[styles.length] = {
		numFmtId:z,
		fontId:0,
		fillId:0,
		borderId:0,
		xfId:0,
		applyNumberFormat:1
	};
	return styles.length-1;
}
