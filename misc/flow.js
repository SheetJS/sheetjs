/*::
type ZIPFile = any;

type XLString = {
	t:string;
	r?:string;
	h?:string;
};

type WorkbookFile = any;

type Workbook = {
	SheetNames: Array<string>;
	Sheets: any;

	Props?: any;
	Custprops?: any;
	Themes?: any;

	SSF?: {[n:number]:string};
	cfb?: any;
};

interface CellAddress {
	r:number;
	c:number;
};

type Cell = any;

type Range = {
	s: CellAddress;
	e: CellAddress;
}
type Worksheet = any;

type Sheet2CSVOpts = any;
type Sheet2JSONOpts = any;

type ParseOpts = any;

type WriteOpts = any;
type WriteFileOpts = any;

type RawData = any;

interface TypeOpts {
	type:string;
}

type XLSXModule = any;

type SST = {
	[n:number]:XLString;
	Count:number;
	Unique:number;
	push(x:XLString):void;
	length:number;
};

type Comment = any;
*/
