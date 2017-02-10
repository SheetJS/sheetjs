/*::
declare type Block = any;
declare type BufArray = {
	newblk(sz:number):Block;
	next(sz:number):Block;
	end():any;
	push(buf:Block):void;
};

type RecordHopperCB = {(d:any, R:any, RT:number):?boolean;};

type EvertType = {[string]:string};
type EvertNumType = {[string]:number};
type EvertArrType = {[string]:Array<string>};

type StringConv = {(string):string};
*/
