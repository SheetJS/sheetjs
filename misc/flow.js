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
	Sheets: {[name:string]:Worksheet};

	Props?: any;
	Custprops?: any;
	Themes?: any;

	Workbook?: WBWBProps;

	SSF?: SSFTable;
	cfb?: any;
	vbaraw?: any;
};

type WBWBProps = {
	Sheets: Array<WBWSProp>;
	Names?: Array<any>;
	WBProps?: WBProps;
};

type WBProps = {
	allowRefreshQuery?: boolean;
	autoCompressPictures?: boolean;
	backupFile?: boolean;
	checkCompatibility?: boolean;
	codeName?: string;
	date1904?: boolean;
	defaultThemeVersion?: number;
	filterPrivacy?: boolean;
	hidePivotFieldList?: boolean;
	promptedSolutions?: boolean;
	publishItems?: boolean;
	refreshAllConnections?: boolean;
	saveExternalLinkValues?: boolean;
	showBorderUnselectedTables?: boolean;
	showInkAnnotation?: boolean;
	showObjects?: string;
	showPivotChartFilter?: boolean;
	updateLinks?: string;
};

type WBWSProp = {
	Hidden?: number;
	name?: string;
};

interface CellAddress {
	r:number;
	c:number;
};
type CellAddrSpec = CellAddress | string;

type Cell = any;

type Range = {
	s: CellAddress;
	e: CellAddress;
}
type Worksheet = any;

type Sheet2CSVOpts = any;
type Sheet2JSONOpts = any;
type Sheet2HTMLOpts = any;

type ParseOpts = any;

type WriteOpts = any;
type WriteFileOpts = any;

type RawData = any;

interface TypeOpts {
	type?:string;
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

type RowInfo = {
	hidden?:boolean; // if true, the row is hidden

	hpx?:number;     // height in screen pixels
	hpt?:number;     // height in points
	level?:number;   // outline / group level
};

type ColInfo = {
	hidden?:boolean; // if true, the column is hidden

	wpx?:number;     // width in screen pixels
	width?:number;    // width in Excel's "Max Digit Width", width*256 is integral
	wch?:number;     // width in characters

	MDW?:number;     // Excel's "Max Digit Width" unit, always integral
};

interface Margins {
	left?:number;
	right?:number;
	top?:number;
	bottom?:number;
	header?:number;
	footer?:number;
};

interface DefinedName {
	Name:string;
	Ref:string;
	Sheet?:number;
	Comment?:string;
};

interface Hyperlink {
	Target:string;
	Tooltip?:string;
};

type SSFTable = any;

type AOA = Array<Array<any> >;
*/
