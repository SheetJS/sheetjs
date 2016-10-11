var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

function get_sst_id(sst, str) {
	for(var i = 0, len = sst.length; i < len; ++i) if(sst[i].t === str) { sst.Count ++; return i; }
	sst[len] = {t:str}; sst.Count ++; sst.Unique ++; return len;
}

function get_cell_style(styles, cell, opts) {
  if (typeof style_builder != 'undefined') {
    if (/^\d+$/.exec(cell.s)) { return cell.s}  // if its already an integer index, let it be
    if (cell.s && (cell.s == +cell.s)) { return cell.s}  // if its already an integer index, let it be
    var s = cell.s || {};
    if (cell.z) s.numFmt = cell.z;
    return style_builder.addStyle(s);
  }
  else {
    var z = opts.revssf[cell.z != null ? cell.z : "General"];
    for(var i = 0, len = styles.length; i != len; ++i) if(styles[i].numFmtId === z) return i;
    styles[len] = {
      numFmtId:z,
      fontId:0,
      fillId:0,
      borderId:0,
      xfId:0,
      applyNumberFormat:1
    };
    return len;
  }
}

function get_cell_style_csf(cellXf) {

  if (cellXf) {

    var s = {}

    if (typeof cellXf.numFmtId != undefined)  {
      s.numFmt = SSF._table[cellXf.numFmtId];
    }

    if(cellXf.fillId)  {
      s.fill =  styles.Fills[cellXf.fillId];
    }

    if (cellXf.fontId) {
      s.font = styles.Fonts[cellXf.fontId];
    }
    if (cellXf.borderId) {
      s.border = styles.Borders[cellXf.borderId];
    }
    if (cellXf.applyAlignment==1) {
      s.alignment = cellXf.alignment;
    }


    return JSON.parse(JSON.stringify(s));
  }
  return null;
}

function safe_format(p, fmtid, fillid, opts) {
	try {
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v,_ssfopts);
				else p.w = SSF._general_num(p.v,_ssfopts);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v);
				if((dd|0) === dd) p.w = SSF._general_int(dd,_ssfopts);
				else p.w = SSF._general_num(dd,_ssfopts);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF._general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF.format(fmtid,datenum(p.v),_ssfopts);
		else p.w = SSF.format(fmtid,p.v,_ssfopts);
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
}
