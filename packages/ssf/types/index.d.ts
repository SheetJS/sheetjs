/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
// TypeScript Version: 2.2

/** Version string */
export const version: string;

/** Render value using format string or code */
export function format(fmt: string|number, val: any, opts?: any): string;

/** Load format string */
export function load(fmt: string, idx?: number): number;

/** Test if the format is a Date format */
export function is_date(fmt: string): boolean;


/** Format Table */
export interface SSF$Table {
	[key: number]: string;
	[key: string]: string;
}

/** Get format table */
export function get_table(): SSF$Table;

/** Set format table */
export function load_table(tbl: SSF$Table): void;

/** Find the relevant sub-format for a given value*/
export function choose_format(fmt: string, value: any): [number, string];

/** Parsed date */
export interface SSF$Date {
	/** number of whole days since relevant epoch, 0 <= D */
	D: number;
	/** integral year portion, epoch_year <= y */
	y: number;
	/** integral month portion, 1 <= m <= 12 */
	m: number;
	/** integral day portion, subject to gregorian YMD constraints */
	d: number;
	/** integral day of week (0=Sunday .. 6=Saturday) 0 <= q <= 6 */
	q: number;

	/** number of seconds since midnight, 0 <= T < 86400 */
	T: number;
	/** integral number of hours since midnight, 0 <= H < 24 */
	H: number;
	/** integral number of minutes since the last hour, 0 <= M < 60 */
	M: number;
	/** integral number of seconds since the last minute, 0 <= S < 60 */
	S: number;
	/** sub-second part of time, 0 <= u < 1 */
	u: number;
}

/** Parse numeric date code */
export function parse_date_code(v: number, opts?: any): SSF$Date;
