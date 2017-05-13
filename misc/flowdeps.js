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
