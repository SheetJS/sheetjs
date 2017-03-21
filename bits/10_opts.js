/* Options */
var opts_fmt/*:Array<Array<any> >*/ = [
	["date1904", 0],
	["output", ""],
	["WTF", false]
];
function fixopts(o){
	for(var y = 0; y != opts_fmt.length; ++y) if(o[opts_fmt[y][0]]===undefined) o[opts_fmt[y][0]]=opts_fmt[y][1];
}
SSF.opts = opts_fmt;
