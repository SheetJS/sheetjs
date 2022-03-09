function split_fmt(fmt/*:string*/)/*:Array<string>*/ {
	var out/*:Array<string>*/ = [];
	var in_str = false/*, cc*/;
	for(var i = 0, j = 0; i < fmt.length; ++i) switch((/*cc=*/fmt.charCodeAt(i))) {
		case 34: /* '"' */
			in_str = !in_str; break;
		case 95: case 42: case 92: /* '_' '*' '\\' */
			++i; break;
		case 59: /* ';' */
			out[out.length] = fmt.substr(j,i-j);
			j = i+1;
	}
	out[out.length] = fmt.substr(j);
	if(in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
	return out;
}
SSF._split = split_fmt;
