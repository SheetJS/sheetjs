/* Part 3 TODO: actually parse formulae */
function ods_to_csf_formula(f/*:string*/)/*:string*/ {
	if(f.substr(0,3) == "of:") f = f.substr(3);
	/* 5.2 Basic Expressions */
	if(f.charCodeAt(0) == 61) {
		f = f.substr(1);
		if(f.charCodeAt(0) == 61) f = f.substr(1);
	}
	f = f.replace(/COM\.MICROSOFT\./g, "");
	/* Part 3 Section 5.8 References */
	f = f.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g, function($$, $1) { return $1.replace(/\./g,""); });
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
	var a = r.split(":");
	var s = a[0].split(".")[0];
	return [s, a[0].split(".")[1] + (a.length > 1 ? (":" + (a[1].split(".")[1] || a[1].split(".")[0])) : "")];
}

function csf_to_ods_3D(r/*:string*/)/*:string*/ {
	return r.replace(/\./,"!");
}

