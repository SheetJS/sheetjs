/* map from xlml named formats to SSF TODO: localize */
var XLMLFormatMap/*{[string]:string}*/ = ({
	"General Number": "General",
	"General Date": SSF._table[22],
	"Long Date": "dddd, mmmm dd, yyyy",
	"Medium Date": SSF._table[15],
	"Short Date": SSF._table[14],
	"Long Time": SSF._table[19],
	"Medium Time": SSF._table[18],
	"Short Time": SSF._table[20],
	"Currency": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"Fixed": SSF._table[2],
	"Standard": SSF._table[4],
	"Percent": SSF._table[10],
	"Scientific": SSF._table[11],
	"Yes/No": '"Yes";"Yes";"No";@',
	"True/False": '"True";"True";"False";@',
	"On/Off": '"Yes";"Yes";"No";@'
}/*:any*/);

var SSFImplicit/*{[number]:string}*/ = ({
	"5": '"$"#,##0_);\\("$"#,##0\\)',
	"6": '"$"#,##0_);[Red]\\("$"#,##0\\)',
	"7": '"$"#,##0.00_);\\("$"#,##0.00\\)',
	"8": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"23": 'General', "24": 'General', "25": 'General', "26": 'General',
	"27": 'm/d/yy', "28": 'm/d/yy', "29": 'm/d/yy', "30": 'm/d/yy', "31": 'm/d/yy',
	"32": 'h:mm:ss', "33": 'h:mm:ss', "34": 'h:mm:ss', "35": 'h:mm:ss',
	"36": 'm/d/yy',
	"41": '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
	"42": '_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
	"43": '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
	"44": '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
	"50": 'm/d/yy', "51": 'm/d/yy', "52": 'm/d/yy', "53": 'm/d/yy', "54": 'm/d/yy',
	"55": 'm/d/yy', "56": 'm/d/yy', "57": 'm/d/yy', "58": 'm/d/yy',
	"59": '0',
	"60": '0.00',
	"61": '#,##0',
	"62": '#,##0.00',
	"63": '"$"#,##0_);\\("$"#,##0\\)',
	"64": '"$"#,##0_);[Red]\\("$"#,##0\\)',
	"65": '"$"#,##0.00_);\\("$"#,##0.00\\)',
	"66": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"67": '0%',
	"68": '0.00%',
	"69": '# ?/?',
	"70": '# ??/??',
	"71": 'm/d/yy',
	"72": 'm/d/yy',
	"73": 'd-mmm-yy',
	"74": 'd-mmm',
	"75": 'mmm-yy',
	"76": 'h:mm',
	"77": 'h:mm:ss',
	"78": 'm/d/yy h:mm',
	"79": 'mm:ss',
	"80": '[h]:mm:ss',
	"81": 'mmss.0'
}/*:any*/);

/* dateNF parse TODO: move to SSF */
var dateNFregex = /[dD]+|[mM]+|[yYeE]+|[Hh]+|[Ss]+/g;
function dateNF_regex(dateNF/*:string|number*/)/*:RegExp*/ {
	var fmt = typeof dateNF == "number" ? SSF._table[dateNF] : dateNF;
	fmt = fmt.replace(dateNFregex, "(\\d+)");
	return new RegExp("^" + fmt + "$");
}
function dateNF_fix(str/*:string*/, dateNF/*:string*/, match/*:Array<string>*/)/*:string*/ {
	var Y = -1, m = -1, d = -1, H = -1, M = -1, S = -1;
	(dateNF.match(dateNFregex)||[]).forEach(function(n, i) {
		var v = parseInt(match[i+1], 10);
		switch(n.toLowerCase().charAt(0)) {
			case 'y': Y = v; break; case 'd': d = v; break;
			case 'h': H = v; break; case 's': S = v; break;
			case 'm': if(H >= 0) M = v; else m = v; break;
		}
	});
	if(S >= 0 && M == -1 && m >= 0) { M = m; m = -1; }
	var datestr = (("" + (Y>=0?Y: new Date().getFullYear())).slice(-4) + "-" + ("00" + (m>=1?m:1)).slice(-2) + "-" + ("00" + (d>=1?d:1)).slice(-2));
	if(datestr.length == 7) datestr = "0" + datestr;
	if(datestr.length == 8) datestr = "20" + datestr;
	var timestr = (("00" + (H>=0?H:0)).slice(-2) + ":" + ("00" + (M>=0?M:0)).slice(-2) + ":" + ("00" + (S>=0?S:0)).slice(-2));
	if(H == -1 && M == -1 && S == -1) return datestr;
	if(Y == -1 && m == -1 && d == -1) return timestr;
	return datestr + "T" + timestr;
}

