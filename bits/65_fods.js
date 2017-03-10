/* Part 3 TODO: actually parse formulae */
function ods_to_csf_formula(f/*:string*/)/*:string*/ {
	if(f.substr(0,3) == "of:") f = f.substr(3);
	/* 5.2 Basic Expressions */
	if(f.charCodeAt(0) == 61) {
		f = f.substr(1);
		if(f.charCodeAt(0) == 61) f = f.substr(1);
	}
	/* Part 3 Section 5.8 References */
	return f.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g, "$1").replace(/\./g, "");
}
