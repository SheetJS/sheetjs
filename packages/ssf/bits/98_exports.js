SSF._table = table_fmt;
SSF.get_table = function get_table()/*:SSFTable*/ { return table_fmt; };
SSF.load_table = function load_table(tbl/*:SSFTable*/)/*:void*/ {
	for(var i=0; i!=0x0188; ++i)
		if(tbl[i] !== undefined) load_entry(tbl[i], i);
};
SSF.init_table = init_table;
SSF.format = format;
SSF.choose_format = choose_fmt;
