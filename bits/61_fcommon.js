/* TODO: it will be useful to parse the function str */
var rc_to_a1 = (function(){
	var rcregex = /(^|[^A-Za-z])R(\[?)(-?\d+|)\]?C(\[?)(-?\d+|)\]?/g;
	var rcbase;
	function rcfunc($$,$1,$2,$3,$4,$5) {
		var R = $3.length>0?parseInt($3,10)|0:0, C = $5.length>0?parseInt($5,10)|0:0;
		if(C<0 && $4.length === 0) C=0;
		if($4.length > 0) C += rcbase.c;
		if($2.length > 0) R += rcbase.r;
		return $1 + encode_col(C) + encode_row(R);
	}
	return function rc_to_a1(fstr, base) {
		rcbase = base;
		return fstr.replace(rcregex, rcfunc);
	};
})();

