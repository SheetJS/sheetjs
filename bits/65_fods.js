/* Part 3 TODO: actually parse formulae */
function ods_to_csf_formula(f/*:string*/)/*:string*/ {
	if(f.slice(0,3) == "of:") f = f.slice(3);
	/* 5.2 Basic Expressions */
	if(f.charCodeAt(0) == 61) {
		f = f.slice(1);
		if(f.charCodeAt(0) == 61) f = f.slice(1);
	}
	f = f.replace(/COM\.MICROSOFT\./g, "");
	/* Part 3 Section 5.8 References */
	f = f.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g, function($$, $1) { return $1.replace(/\./g,""); });
	f = f.replace(/\$'([^']|'')+'/g, function($$) { return $$.slice(1); });
	f = f.replace(/\$([^\]\. #$]+)/g, function($$, $1) { return ($1).match(/^([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])?(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})?$/) ? $$ : $1; });
	/* TODO: something other than this */
	f = f.replace(/\[.(#[A-Z]*[?!])\]/g, "$1");
	return f.replace(/[;~]/g,",").replace(/\|/g,";");
}

function csf_to_ods_formula(f/*:string*/)/*:string*/ {
	var o = "of:=" + f.replace(crefregex, "$1[.$2$3$4$5]").replace(/\]:\[/g,":");
	/* TODO: something other than this */
	return o.replace(/;/g, "|").replace(/,/g,";");
}

function ods_to_csf_3D(r/*:string*/)/*:[string, string]*/ {
	r = r.replace(/\$'([^']|'')+'/g, function($$) { return $$.slice(1); });
	r = r.replace(/\$([^\]\. #$]+)/g, function($$, $1) { return ($1).match(/^([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])?(10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6]|[1-9]\d{0,5})?$/) ? $$ : $1; });
	var a = r.split(":");
	var s = a[0].split(".")[0];
	return [s, a[0].split(".")[1] + (a.length > 1 ? (":" + (a[1].split(".")[1] || a[1].split(".")[0])) : "")];
}

function csf_to_ods_3D(r/*:string*/)/*:string*/ {
	return r.replace(/!/,".");
}

