function load_entry(fmt/*:string*/, idx/*:?number*/)/*:number*/ {
	if(typeof idx != 'number') {
		idx = +idx || -1;
/*::if(typeof idx != 'number') return 0x188; */
		for(var i = 0; i < 0x0188; ++i) {
/*::if(typeof idx != 'number') return 0x188; */
			if(table_fmt[i] == undefined) { if(idx < 0) idx = i; continue; }
			if(table_fmt[i] == fmt) { idx = i; break; }
		}
/*::if(typeof idx != 'number') return 0x188; */
		if(idx < 0) idx = 0x187;
	}
/*::if(typeof idx != 'number') return 0x188; */
	table_fmt[idx] = fmt;
	return idx;
}
SSF.load = load_entry;
