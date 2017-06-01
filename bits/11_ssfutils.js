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

