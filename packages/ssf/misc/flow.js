/*# vim: set ts=2: */
/*::

type SSFTable = {[key:number|string]:string};
type SSFDate = {
	D:number; T:number;
	y:number; m:number; d:number; q:number;
	H:number; M:number; S:number; u:number;
};

type SSFModule = {
	format:(fmt:string|number, v:any, o:any)=>string;

	is_date:(fmt:string)=>boolean;
	parse_date_code:(v:number,opts:any)=>?SSFDate;

	load:(fmt:string, idx:?number)=>number;
	get_table:()=>SSFTable;
	load_table:(table:any)=>void;
	_table:SSFTable;
	init_table:any;

	_general_num:(v:number)=>string;
	_general:(v:number, o:?any)=>string;
	_eval:any;
	_split:any;
	version:string;
};

*/
