/* ECMA-376 18.8.30 numFmt*/
/* Note: `toPrecision` uses standard form when prec > E and E >= -6 */
var general_fmt_num = (function make_general_fmt_num() {
	var trailing_zeroes_and_decimal = /(?:\.0*|(\.\d*[1-9])0+)$/;
	function strip_decimal(o/*:string*/)/*:string*/ {
		return (o.indexOf(".") == -1) ? o : o.replace(trailing_zeroes_and_decimal, "$1");
	}

	/* General Exponential always shows 2 digits exp and trims the mantissa */
	var mantissa_zeroes_and_decimal = /(?:\.0*|(\.\d*[1-9])0+)[Ee]/;
	var exp_with_single_digit = /(E[+-])(\d)$/;
	function normalize_exp(o/*:string*/)/*:string*/ {
		if(o.indexOf("E") == -1) return o;
		return o.replace(mantissa_zeroes_and_decimal,"$1E").replace(exp_with_single_digit,"$10$2");
	}

	/* exponent >= -9 and <= 9 */
	function small_exp(v/*:number*/)/*:string*/ {
		var w = (v<0?12:11);
		var o = strip_decimal(v.toFixed(12)); if(o.length <= w) return o;
		o = v.toPrecision(10); if(o.length <= w) return o;
		return v.toExponential(5);
	}

	/* exponent >= 11 or <= -10 likely exponential */
	function large_exp(v/*:number*/)/*:string*/ {
		var o = strip_decimal(v.toFixed(11));
		return (o.length > (v<0?12:11) || o === "0" || o === "-0") ? v.toPrecision(6) : o;
	}

	function general_fmt_num_base(v/*:number*/)/*:string*/ {
		var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;

		if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
		else if(Math.abs(V) <= 9) o = small_exp(v);
		else if(V === 10) o = v.toFixed(10).substr(0,12);
		else o = large_exp(v);

		return strip_decimal(normalize_exp(o.toUpperCase()));
	}

	return general_fmt_num_base;
})();
SSF._general_num = general_fmt_num;

/*
	"General" rules:
	- text is passed through ("@")
	- booleans are rendered as TRUE/FALSE
	- "up to 11 characters" displayed for numbers
	- Default date format (code 14) used for Dates

	The longest 32-bit integer text is "-2147483648", exactly 11 chars
	TODO: technically the display depends on the width of the cell
*/
function general_fmt(v/*:any*/, opts/*:any*/) {
	switch(typeof v) {
		case 'string': return v;
		case 'boolean': return v ? "TRUE" : "FALSE";
		case 'number': return (v|0) === v ? v.toString(10) : general_fmt_num(v);
		case 'undefined': return "";
		case 'object':
			if(v == null) return "";
			if(v instanceof Date) return format(14, datenum_local(v, opts && opts.date1904), opts);
	}
	throw new Error("unsupported value in General format: " + v);
}
SSF._general = general_fmt;
