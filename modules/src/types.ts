declare type RawData = Uint8Array | number[];
declare function recordhopper(data: RawData, cb:(val: any, R_n: string, RT: number)=>void): void;
declare interface ReadableData {
	l: number;
	read_shift(t: 4): number;
	read_shift(t: any): any;
}
declare type ParseFunc<T> = (data: ReadableData, length: number) => T;
declare var parse_XLWideString: ParseFunc<string>;

declare interface WritableData {
	l: number;
	write_shift(t: 4, val: number): void;
	write_shift(t: number, val: string|number, f?: string): any;
}
declare type WritableRawData = WritableData & RawData;
interface BufArray {
	end(): RawData;
	next(sz: number): WritableData;
	push(buf: RawData): void;
}
declare function buf_array(): BufArray;
declare function write_record(ba: BufArray, type: string, payload?: RawData, length?: number): void;
declare function new_buf(sz: number): RawData & WritableData & ReadableData;

declare var tagregex: RegExp;
declare var XML_HEADER: string;
declare var RELS: any;
declare function parsexmltag(tag: string, skip_root?: boolean, skip_LC?: boolean): object;
declare function strip_ns(x: string): string;
declare function write_UInt32LE(x: number, o?: WritableData): RawData;
declare function write_XLWideString(data: string, o?: WritableData): RawData;
declare function writeuint16(x: number): RawData;


interface ParseXLMetaOptions {
	WTF?: number|boolean;
}
interface XLMDT {
	name: string;
}
interface XLMeta {
	Types: XLMDT[];
}