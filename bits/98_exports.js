SSF._table = table_fmt;
SSF.load = function load_entry(fmt/*:string*/, idx/*:number*/) { table_fmt[idx] = fmt; };
SSF.format = format;
SSF.get_table = function get_table() { return table_fmt; };
SSF.load_table = function load_table(tbl/*:{[n:number]:string}*/) { for(var i=0; i!=0x0188; ++i) if(tbl[i] !== undefined) SSF.load(tbl[i], i); };
SSF.init_table = init_table;
