/*::
declare module 'exit-on-epipe' {};

declare module 'xlsx' { declare var exports:XLSXModule; };
declare module '../' { declare var exports:XLSXModule; };

declare module 'commander' { declare var exports:any; };
declare module './jszip.js' { declare var exports:any; };
declare module './dist/cpexcel.js' { declare var exports:any; };
declare module 'crypto' { declare var exports:any; };
declare module 'fs' { declare var exports:any; };

type ZIP = any;

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

	_general_int:(v:number)=>string;
	_general_num:(v:number)=>string;
	_general:(v:number, o:?any)=>string;
	_eval:any;
	_split:any;
	version:string;
};

// ----------------------------------------------------------------------------
// Note: The following override is needed because ReadShift includes more cases
// ----------------------------------------------------------------------------
type ReadShiftFunc = any;
type WriteShiftFunc = any;

type CFBModule = {
	version:string;
	find:(cfb:CFBContainer, path:string)=>?CFBEntry;
	read:(blob:RawBytes|string, opts:CFBReadOpts)=>CFBContainer;
	write:(cfb:CFBContainer, opts:CFBWriteOpts)=>RawBytes|string;
	writeFile:(cfb:CFBContainer, filename:string, opts:CFBWriteOpts)=>void;
	parse:(file:RawBytes, opts:CFBReadOpts)=>CFBContainer;
	utils:CFBUtils;
};

type CFBFullPathDir = {
	[n:string]: CFBEntry;
}

type CFBUtils = any;

type CheckFieldFunc = {(hexstr:string, fld:string):void;};

type RawBytes = Array<number> | Buffer | Uint8Array;

class CFBlobArray extends Array<number> {
	l:number;
	write_shift:WriteShiftFunc;
	read_shift:ReadShiftFunc;
	chk:CheckFieldFunc;
};
interface CFBlobBuffer extends Buffer {
	l:number;
	slice:(start?:number, end:?number)=>Buffer;
	write_shift:WriteShiftFunc;
	read_shift:ReadShiftFunc;
	chk:CheckFieldFunc;
};
interface CFBlobUint8 extends Uint8Array {
	l:number;
	slice:(start?:number, end:?number)=>Uint8Array;
	write_shift:WriteShiftFunc;
	read_shift:ReadShiftFunc;
	chk:CheckFieldFunc;
};

interface CFBlobber {
	[n:number]:number;
	l:number;
	length:number;
	slice:(start:?number, end:?number)=>RawBytes;
	write_shift:WriteShiftFunc;
	read_shift:ReadShiftFunc;
	chk:CheckFieldFunc;
};

type CFBlob = CFBlobArray | CFBlobBuffer | CFBlobUint8;

type CFBWriteOpts = any;

interface CFBReadOpts {
	type?:string;
};

type CFBFileIndex = Array<CFBEntry>;

type CFBFindPath = (n:string)=>?CFBEntry;

type CFBContainer = {
	raw?:{
		header:any;
		sectors:Array<any>;
	};
	FileIndex:CFBFileIndex;
	FullPathDir:CFBFullPathDir;
	FullPaths:Array<string>;
}

type CFBEntry = {
	name: string;
	type: number;
	ct?: Date;
	mt?: Date;
	color: number;
	clsid: string;
	state: number;
	start: number;
	size: number;
	storage?: "fat" | "minifat";
	L: number;
	R: number;
	C: number;
	content?: CFBlob;
}


// ----------------------------------------------------------------------------
// Note: The following override is needed because Flow is missing Date#getYear
// ----------------------------------------------------------------------------

type Date$LocaleOptions = {
	localeMatcher?: string,
	timeZone?: string,
	hour12?: boolean,
	formatMatcher?: string,
	weekday?: string,
	era?: string,
	year?: string,
	month?: string,
	day?: string,
	hour?: string,
	minute?: string,
	second?: string,
	timeZoneName?: string,
};

declare class Date {
	constructor(): void;
	constructor(timestamp: number): void;
	constructor(dateString: string): void;
	constructor(year: number, month: number, day?: number, hour?: number, minute?: number, second?: number, millisecond?: number): void;
	getDate(): number;
	getDay(): number;
	getYear(): number;
	getFullYear(): number;
	getHours(): number;
	getMilliseconds(): number;
	getMinutes(): number;
	getMonth(): number;
	getSeconds(): number;
	getTime(): number;
	getTimezoneOffset(): number;
	getUTCDate(): number;
	getUTCDay(): number;
	getUTCFullYear(): number;
	getUTCHours(): number;
	getUTCMilliseconds(): number;
	getUTCMinutes(): number;
	getUTCMonth(): number;
	getUTCSeconds(): number;
	setDate(date: number): number;
	setFullYear(year: number, month?: number, date?: number): number;
	setHours(hours: number, min?: number, sec?: number, ms?: number): number;
	setMilliseconds(ms: number): number;
	setMinutes(min: number, sec?: number, ms?: number): number;
	setMonth(month: number, date?: number): number;
	setSeconds(sec: number, ms?: number): number;
	setTime(time: number): number;
	setUTCDate(date: number): number;
	setUTCFullYear(year: number, month?: number, date?: number): number;
	setUTCHours(hours: number, min?: number, sec?: number, ms?: number): number;
	setUTCMilliseconds(ms: number): number;
	setUTCMinutes(min: number, sec?: number, ms?: number): number;
	setUTCMonth(month: number, date?: number): number;
	setUTCSeconds(sec: number, ms?: number): number;
	toDateString(): string;
	toISOString(): string;
	toJSON(key?: any): string;
	toLocaleDateString(locales?: string, options?: Date$LocaleOptions): string;
	toLocaleString(locales?: string, options?: Date$LocaleOptions): string;
	toLocaleTimeString(locales?: string, options?: Date$LocaleOptions): string;
	toTimeString(): string;
	toUTCString(): string;
	valueOf(): number;
	static ():string;
	static now(): number;
	static parse(s: string): number;
	static UTC(year: number, month: number, date?: number, hours?: number, minutes?: number, seconds?: number, ms?: number): number;
	// multiple indexers not yet supported
	[key: $SymbolToPrimitive]: (hint: 'string' | 'default' | 'number') => string | number;
}

*/
