/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
/// <reference path="src/types.ts"/>

/* these are type imports and do not show up in the generated JS */
import { CFB$Container, CFB$Entry } from 'cfb';
import { WorkBook, WorkSheet, Range, CellObject, ParsingOptions, WritingOptions } from '../';
import type { utils } from "../";

declare var encode_cell: typeof utils.encode_cell;
declare var encode_range: typeof utils.encode_range;
declare var book_new: typeof utils.book_new;
declare var book_append_sheet: typeof utils.book_append_sheet;
declare var sheet_to_json: typeof utils.sheet_to_json;
declare var decode_range: typeof utils.decode_range;
import * as _CFB from 'cfb';
declare var CFB: typeof _CFB;
//<<import { utils } from "../../";
//<<const { encode_cell, encode_range, book_new, book_append_sheet } = utils;

/* see https://bugs.webkit.org/show_bug.cgi?id=243148 -- affects iOS Safari */
declare var Buffer: any; // Buffer is typeof-guarded but TS still needs this :(
var subarray: "subarray" | "slice" = (() => {
	try {
	if(typeof Uint8Array == "undefined") return "slice";
	if(typeof Uint8Array.prototype.subarray == "undefined") return "slice";
	// NOTE: feature tests are for node < 6.x
	if(typeof Buffer !== "undefined") {
		if(typeof Buffer.prototype.subarray == "undefined") return "slice";
		if((typeof Buffer.from == "function" ? Buffer.from([72,62]) : new Buffer([72,62])) instanceof Uint8Array) return "subarray";
		return "slice";
	}
	return "subarray";
	} catch(e) { return "slice"; }
})();

function u8_to_dataview(array: Uint8Array): DataView { return new DataView(array.buffer, array.byteOffset, array.byteLength); }
//<<export { u8_to_dataview };

function u8str(u8: Uint8Array): string { return /* Buffer.isBuffer(u8) ? u8.toString() :*/ typeof TextDecoder != "undefined" ? new TextDecoder().decode(u8) : utf8read(a2s(u8)); }
function stru8(str: string): Uint8Array { return typeof TextEncoder != "undefined" ? new TextEncoder().encode(str) : s2a(utf8write(str)) as Uint8Array; }
//<<export { u8str, stru8 };

function u8contains(body: Uint8Array, search: Uint8Array): boolean {
	var L = body.indexOf(search[0]);
	if(L == -1) return false;
	outer: for(; L <= body.length - search.length; ++L) {
		for(var j = 0; j < search.length; ++j) if(body[L+j] != search[j]) continue outer;
		return true;
	}
	return false;
}
//<<export { u8contains }

/** Concatenate Uint8Arrays */
function u8concat(u8a: Uint8Array[]): Uint8Array {
	var len = u8a.reduce((acc: number, x: Uint8Array) => acc + x.length, 0);
	var out = new Uint8Array(len);
	var off = 0;
	u8a.forEach(u8 => { out.set(u8, off); off += u8.length; });
	return out;
}
//<<export { u8concat };

/** Count the number of bits set (assuming int32_t interpretation) */
function popcnt(x: number): number {
	x -= ((x >> 1) & 0x55555555);
	x = (x & 0x33333333) + ((x >> 2) & 0x33333333);
	return (((x + (x >> 4)) & 0x0F0F0F0F) * 0x01010101) >>> 24;
}

/** Read a 128-bit decimal from the modern cell storage */
function readDecimal128LE(buf: Uint8Array, offset: number): number {
	var exp = ((buf[offset + 15] & 0x7F) << 7) | (buf[offset + 14] >> 1);
	var mantissa = buf[offset + 14] & 1;
	for(var j = offset + 13; j >= offset; --j) mantissa = mantissa * 256 + buf[j];
	return ((buf[offset+15] & 0x80) ? -mantissa : mantissa) * Math.pow(10, exp - 0x1820);
}
/** Write a 128-bit decimal to the modern cell storage */
function writeDecimal128LE(buf: Uint8Array, offset: number, value: number): void {
	// TODO: something more correct than this
	var exp = Math.floor(value == 0 ? 0 : /*Math.log10*/Math.LOG10E * Math.log(Math.abs(value))) + 0x1820 - 16;
	var mantissa = (value / Math.pow(10, exp - 0x1820));
	buf[offset+15] |= exp >> 7;
	buf[offset+14] |= (exp & 0x7F) << 1;
	for(var i = 0; mantissa >= 1; ++i, mantissa /= 256) buf[offset + i] = mantissa & 0xFF;
	buf[offset+15] |= (value >= 0 ? 0 : 0x80);
}


type Ptr = [number];

/** Parse an integer from the varint that can be exactly stored in a double */
function parse_varint49(buf: Uint8Array, ptr?: Ptr): number {
	var l = ptr ? ptr[0] : 0;
	var usz = buf[l] & 0x7F;
	varint: if(buf[l++] >= 0x80) {
		usz |= (buf[l] & 0x7F) <<  7; if(buf[l++] < 0x80) break varint;
		usz |= (buf[l] & 0x7F) << 14; if(buf[l++] < 0x80) break varint;
		usz |= (buf[l] & 0x7F) << 21; if(buf[l++] < 0x80) break varint;
		usz += (buf[l] & 0x7F) * Math.pow(2, 28); ++l; if(buf[l++] < 0x80) break varint;
		usz += (buf[l] & 0x7F) * Math.pow(2, 35); ++l; if(buf[l++] < 0x80) break varint;
		usz += (buf[l] & 0x7F) * Math.pow(2, 42); ++l; if(buf[l++] < 0x80) break varint;
	}
	if(ptr) ptr[0] = l;
	return usz;
}
/** Write a varint up to 7 bytes / 49 bits */
function write_varint49(v: number): Uint8Array {
	var usz = new Uint8Array(7);
	usz[0] = (v & 0x7F);
	var L = 1;
	sz: if(v > 0x7F) {
		usz[L-1] |= 0x80; usz[L] = (v >> 7) & 0x7F; ++L;
		if(v <= 0x3FFF) break sz;
		usz[L-1] |= 0x80; usz[L] = (v >> 14) & 0x7F; ++L;
		if(v <= 0x1FFFFF) break sz;
		usz[L-1] |= 0x80; usz[L] = (v >> 21) & 0x7F; ++L;
		if(v <= 0xFFFFFFF) break sz;
		usz[L-1] |= 0x80; usz[L] = ((v/0x100) >>> 21) & 0x7F; ++L;
		if(v <= 0x7FFFFFFFF) break sz;
		usz[L-1] |= 0x80; usz[L] = ((v/0x10000) >>> 21) & 0x7F; ++L;
		if(v <= 0x3FFFFFFFFFF) break sz;
		usz[L-1] |= 0x80; usz[L] = ((v/0x1000000) >>> 21) & 0x7F; ++L;
	}
	return usz[subarray](0, L);
}
/** Parse a repeated varint [packed = true] field */
function parse_packed_varints(buf: Uint8Array): number[] {
	var ptr: Ptr = [0];
	var out: number[] = [];
	while(ptr[0] < buf.length) out.push(parse_varint49(buf, ptr));
	return out;
}
/** Write a repeated varint [packed = true] field */
function write_packed_varints(nums: number[]): Uint8Array {
	return u8concat(nums.map(x => write_varint49(x)));
}
//<<export { parse_varint49, write_varint49 };

/** Parse a 32-bit signed integer from the raw varint */
function varint_to_i32(buf: Uint8Array): number {
	var l = 0, i32 = buf[l] & 0x7F;
	varint: if(buf[l++] >= 0x80) {
		i32 |= (buf[l] & 0x7F) <<  7; if(buf[l++] < 0x80) break varint;
		i32 |= (buf[l] & 0x7F) << 14; if(buf[l++] < 0x80) break varint;
		i32 |= (buf[l] & 0x7F) << 21; if(buf[l++] < 0x80) break varint;
		i32 |= (buf[l] & 0x7F) << 28;
	}
	return i32;
}
/** Parse a 64-bit unsigned integer as a pair */
function varint_to_u64(buf: Uint8Array): [number, number] {
	var l = 0, lo = buf[l] & 0x7F, hi = 0;
	varint: if(buf[l++] >= 0x80) {
		lo |= (buf[l] & 0x7F) <<  7; if(buf[l++] < 0x80) break varint;
		lo |= (buf[l] & 0x7F) << 14; if(buf[l++] < 0x80) break varint;
		lo |= (buf[l] & 0x7F) << 21; if(buf[l++] < 0x80) break varint;
		lo |= (buf[l] & 0x7F) << 28; hi = (buf[l] >> 4) & 0x07; if(buf[l++] < 0x80) break varint;
		hi |= (buf[l] & 0x7F) <<  3; if(buf[l++] < 0x80) break varint;
		hi |= (buf[l] & 0x7F) << 10; if(buf[l++] < 0x80) break varint;
		hi |= (buf[l] & 0x7F) << 17; if(buf[l++] < 0x80) break varint;
		hi |= (buf[l] & 0x7F) << 24; if(buf[l++] < 0x80) break varint;
		hi |= (buf[l] & 0x7F) << 31;
	}
	return [lo >>> 0, hi >>> 0];
}
//<<export { varint_to_i32 };

interface ProtoItem {
	data: Uint8Array;
	type: number;
}
type ProtoField = Array<ProtoItem>
type ProtoMessage = Array<ProtoField>;
/** Shallow parse of a Protobuf message */
function parse_shallow(buf: Uint8Array): ProtoMessage {
	var out: ProtoMessage = [], ptr: Ptr = [0];
	while(ptr[0] < buf.length) {
		var off = ptr[0];
		var num = parse_varint49(buf, ptr);
		var type = num & 0x07; num = Math.floor(num / 8);
		var len = 0;
		var res: Uint8Array;
		if(num == 0) break;
		switch(type) {
			case 0: {
				var l = ptr[0];
				while(buf[ptr[0]++] >= 0x80);
				res = buf[subarray](l, ptr[0]);
			} break;
			case 5: len = 4; res = buf[subarray](ptr[0], ptr[0] + len); ptr[0] += len; break;
			case 1: len = 8; res = buf[subarray](ptr[0], ptr[0] + len); ptr[0] += len; break;
			case 2: len = parse_varint49(buf, ptr); res = buf[subarray](ptr[0], ptr[0] + len); ptr[0] += len; break;
			case 3: // Start group
			case 4: // End group
			default: throw new Error(`PB Type ${type} for Field ${num} at offset ${off}`);
		}
		var v: ProtoItem = { data: res, type };
		if(out[num] == null) out[num] = [v];
		else out[num].push(v);
	}
	return out;
}
/** Serialize a shallow parse */
function write_shallow(proto: ProtoMessage): Uint8Array {
	var out: Uint8Array[] = [];
	proto.forEach((field, idx) => {
		if(idx == 0) return;
		field.forEach(item => {
			if(!item.data) return;
			out.push(write_varint49(idx * 8 + item.type));
			if(item.type == 2) out.push(write_varint49(item.data.length));
			out.push(item.data);
		});
	});
	return u8concat(out);
}
//<<export { parse_shallow, write_shallow };

/** Map over each entry in a repeated (or single-value) field */
function mappa<U>(data: ProtoField, cb:(Uint8Array) => U): U[] {
	return data?.map(d =>  cb(d.data)) || [];
}


interface IWAMessage {
	/** Metadata in .TSP.MessageInfo */
	meta: ProtoMessage;
	data: Uint8Array;
}
interface IWAArchiveInfo {
	id: number;
	merge?: boolean;
	messages: IWAMessage[];
}
/** Extract all messages from a IWA file */
function parse_iwa_file(buf: Uint8Array): IWAArchiveInfo[] {
	var out: IWAArchiveInfo[] = [], ptr: Ptr = [0];
	while(ptr[0] < buf.length) {
		/* .TSP.ArchiveInfo */
		var len = parse_varint49(buf, ptr);
		var ai = parse_shallow(buf[subarray](ptr[0], ptr[0] + len));
		ptr[0] += len;

		var res: IWAArchiveInfo = {
			/* TODO: technically ID is optional */
			id: varint_to_i32(ai[1][0].data),
			messages: []
		};
		ai[2].forEach(b => {
			var mi = parse_shallow(b.data);
			var fl = varint_to_i32(mi[3][0].data);
			res.messages.push({
				meta: mi,
				data: buf[subarray](ptr[0], ptr[0] + fl)
			});
			ptr[0] += fl;
		});
		if(ai[3]?.[0]) res.merge = (varint_to_i32(ai[3][0].data) >>> 0) > 0;
		out.push(res);
	}
	return out;
}
/** Generate an IWA file from a parsed structure */
function write_iwa_file(ias: IWAArchiveInfo[]): Uint8Array {
	var bufs: Uint8Array[] = [];
	ias.forEach(ia => {
		var ai: ProtoMessage = [ [],
			[ {data: write_varint49(ia.id), type: 0} ],
			[]
		];
		if(ia.merge != null) ai[3] = [ { data: write_varint49(+!!ia.merge), type: 0 } ];
		var midata: Uint8Array[] = [];

		ia.messages.forEach(mi => {
			midata.push(mi.data);
			mi.meta[3] = [ { type: 0, data: write_varint49(mi.data.length) } ];
			ai[2].push({data: write_shallow(mi.meta), type: 2});
		});

		var aipayload = write_shallow(ai);
		bufs.push(write_varint49(aipayload.length));
		bufs.push(aipayload);
		midata.forEach(mid => bufs.push(mid));
	});
	return u8concat(bufs);
}
//<<export { IWAMessage, IWAArchiveInfo, parse_iwa_file, write_iwa_file };

/** Decompress a snappy chunk */
function parse_snappy_chunk(type: number, buf: Uint8Array): Uint8Array[] {
	if(type != 0) throw new Error(`Unexpected Snappy chunk type ${type}`);
	var ptr: Ptr = [0];

	var usz = parse_varint49(buf, ptr);
	var chunks: Uint8Array[] = [];
	while(ptr[0] < buf.length) {
		var tag = buf[ptr[0]] & 0x3;
		if(tag == 0) {
			var len = buf[ptr[0]++] >> 2;
			if(len < 60) ++len;
			else {
				var c = len - 59;
				len = buf[ptr[0]];
				if(c > 1) len |= (buf[ptr[0]+1]<<8);
				if(c > 2) len |= (buf[ptr[0]+2]<<16);
				if(c > 3) len |= (buf[ptr[0]+3]<<24);
				len >>>=0; len++;
				ptr[0] += c;
			}
			chunks.push(buf[subarray](ptr[0], ptr[0] + len)); ptr[0] += len; continue;
		} else {
			var offset = 0, length = 0;
			if(tag == 1) {
				length = ((buf[ptr[0]] >> 2) & 0x7) + 4;
				offset = (buf[ptr[0]++] & 0xE0) << 3;
				offset |= buf[ptr[0]++];
			} else {
				length = (buf[ptr[0]++] >> 2) + 1;
				if(tag == 2) { offset = buf[ptr[0]] | (buf[ptr[0]+1]<<8); ptr[0] += 2; }
				else { offset = (buf[ptr[0]] | (buf[ptr[0]+1]<<8) | (buf[ptr[0]+2]<<16) | (buf[ptr[0]+3]<<24))>>>0; ptr[0] += 4; }
			}
			if(offset == 0) throw new Error("Invalid offset 0");
			var j = chunks.length - 1, off = offset;
			while(j >=0 && off >= chunks[j].length) { off -= chunks[j].length; --j; }
			if(j < 0) {
				if(off == 0) off = chunks[(j = 0)].length;
				else throw new Error("Invalid offset beyond length");
			}
			// Node 0.8 Buffer slice does not support negative indices
			if(length < off) chunks.push(chunks[j][subarray](chunks[j].length-off, chunks[j].length-off + length));
			else {
				if(off > 0) { chunks.push(chunks[j][subarray](chunks[j].length-off)); length -= off; } ++j;
				while(length >= chunks[j].length) { chunks.push(chunks[j]); length -= chunks[j].length; ++j; }
				if(length) chunks.push(chunks[j][subarray](0, length));
			}
			if(chunks.length > 100) chunks = [u8concat(chunks)];
		}
	}
	if(chunks.reduce((acc, u8) => acc + u8.length, 0) != usz) throw new Error(`Unexpected length: ${chunks.reduce((acc, u8) => acc + u8.length, 0)} != ${usz}`);
	return chunks;
	//var o = u8concat(chunks);
	//if(o.length != usz) throw new Error(`Unexpected length: ${o.length} != ${usz}`);
	//return o;
}

/** Decompress IWA file */
function decompress_iwa_file(buf: Uint8Array): Uint8Array {
	if(Array.isArray(buf)) buf = new Uint8Array(buf);
	var out: Uint8Array[] = [];
	var l = 0;
	while(l < buf.length) {
		var t = buf[l++];
		var len = buf[l] | (buf[l+1]<<8) | (buf[l+2] << 16); l += 3;
		out.push.apply(out, parse_snappy_chunk(t, buf[subarray](l, l + len)));
		l += len;
	}
	if(l !== buf.length) throw new Error("data is not a valid framed stream!");
	return u8concat(out);
}

/** Compress IWA file */
function compress_iwa_file(buf: Uint8Array): Uint8Array {
	var out: Uint8Array[] = [];
	var l = 0;
	while(l < buf.length) {
		var c = Math.min(buf.length - l, 0xFFFFFFF);
		var frame = new Uint8Array(4);
		out.push(frame);
		var usz = write_varint49(c);
		var L = usz.length;
		out.push(usz);

		if(c <= 60) { L++; out.push(new Uint8Array([(c - 1)<<2])); }
		else if(c <= 0x100)       { L += 2; out.push(new Uint8Array([0xF0, (c-1) & 0xFF])); }
		else if(c <= 0x10000)     { L += 3; out.push(new Uint8Array([0xF4, (c-1) & 0xFF, ((c-1) >> 8) & 0xFF])); }
		else if(c <= 0x1000000)   { L += 4; out.push(new Uint8Array([0xF8, (c-1) & 0xFF, ((c-1) >> 8) & 0xFF, ((c-1) >> 16) & 0xFF])); }
		else if(c <= 0x100000000) { L += 5; out.push(new Uint8Array([0xFC, (c-1) & 0xFF, ((c-1) >> 8) & 0xFF, ((c-1) >> 16) & 0xFF, ((c-1) >>> 24) & 0xFF])); }

		out.push(buf[subarray](l, l + c)); L += c;

		frame[0] = 0;
		frame[1] = L & 0xFF; frame[2] = (L >>  8) & 0xFF; frame[3] = (L >> 16) & 0xFF;
		l += c;
	}
	return u8concat(out);
}
//<<export { decompress_iwa_file, compress_iwa_file };

/** .TST.DataStore */
interface DataLUT {
	/** shared string table */
	sst: string[];
	/** rich string table table */
	rsst: string[];
	/** old format table */
	ofmt: ProtoMessage[];
	/** new format table */
	nfmt: ProtoMessage[];
}
var numbers_lut_new = (): DataLUT => ({ sst: [], rsst: [], ofmt: [], nfmt: [] });

function numbers_format_cell(cell: CellObject, t: number, flags: number, ofmt: ProtoMessage, nfmt: ProtoMessage): void {
	var ctype = t & 0xFF, ver = t >> 8;
	var fmt = ver >= 5 ? nfmt : ofmt;
	dur: if((flags & (ver > 4 ? 8: 4)) && cell.t == "n" && ctype == 7) {
		var dstyle =   (fmt[7]?.[0])  ? parse_varint49(fmt[7][0].data)  : -1;
		if(dstyle == -1) break dur;
		var dmin =     (fmt[15]?.[0]) ? parse_varint49(fmt[15][0].data) : -1;
		var dmax =     (fmt[16]?.[0]) ? parse_varint49(fmt[16][0].data) : -1;
		var auto =     (fmt[40]?.[0]) ? parse_varint49(fmt[40][0].data) : -1;
		var d: number = cell.v as number, dd = d;
		autodur: if(auto) { // TODO: check if numbers reformats on load
			if(d == 0) { dmin = dmax = 2; break autodur; }
			if(d >= 604800) dmin = 1;
			else if(d >= 86400) dmin = 2;
			else if(d >= 3600) dmin = 4;
			else if(d >= 60) dmin = 8;
			else if(d >= 1) dmin = 16;
			else dmin = 32;
			if(Math.floor(d) != d) dmax = 32;
			else if(d % 60) dmax = 16;
			else if(d % 3600) dmax = 8;
			else if(d % 86400) dmax = 4;
			else if(d % 604800) dmax = 2;
			if(dmax < dmin) dmax = dmin;
		}
		if(dmin == -1 || dmax == -1) break dur;
		var dstr: string[] = [], zstr: string[] = [];
		/* TODO: plurality, SSF equivalents */
		if(dmin == 1) {
			dd = d / 604800;
			if(dmax == 1) { zstr.push('d"d"'); } else { dd |= 0; d -= 604800 * dd; }
			dstr.push(dd + (dstyle == 2 ? " week" + (dd == 1 ? "" : "s") : dstyle == 1 ? "w": ""));
		}
		if(dmin <= 2 && dmax >= 2) {
			dd = d / 86400;
			if(dmax > 2) { dd |= 0; d -= 86400 * dd; }
			zstr.push("d" + '"d"');
			dstr.push(dd + (dstyle == 2 ? " day" + (dd == 1 ? "" : "s") : dstyle == 1 ? "d" : ""));
		}
		if(dmin <= 4 && dmax >= 4) {
			dd = d / 3600;
			if(dmax > 4) { dd |= 0; d -= 3600 * dd; }
			zstr.push((dmin >= 4 ? "[h]" : "h") + '"h"');
			dstr.push(dd + (dstyle == 2 ? " hour" + (dd == 1 ? "" : "s") : dstyle == 1 ? "h" : ""));
		}
		if(dmin <= 8 && dmax >= 8) {
			dd = d / 60;
			if(dmax > 8) { dd |= 0; d -= 60 * dd; }
			zstr.push((dmin >= 8 ? "[m]" : "m") + '"m"');
			if(dstyle == 0) dstr.push(((dmin == 8 && dmax == 8 || dd >= 10) ? "" : "0") + dd);
			else dstr.push(dd + (dstyle == 2 ? " minute" + (dd == 1 ? "" : "s") : dstyle == 1 ? "m" : ""));
		}
		if(dmin <= 16 && dmax >= 16) {
			dd = d;
			if(dmax > 16) { dd |= 0; d -= dd; }
			zstr.push((dmin >= 16 ? "[s]" : "s") + '"s"');
			if(dstyle == 0) dstr.push((dmax == 16 && dmin == 16 || dd >= 10 ? "" : "0") + dd);
			else dstr.push(dd + (dstyle == 2 ? " second" + (dd == 1 ? "" : "s") : dstyle == 1 ? "s" : ""));
		}
		if(dmax >= 32) {
			dd = Math.round(1000 * d);
			if(dmin < 32) zstr.push(".000" + '"ms"');
			if(dstyle == 0) dstr.push((dd >= 100 ? "" : dd >= 10 ? "0" : "00") + dd);
			else dstr.push(dd + (dstyle == 2 ? " millisecond" + (dd == 1 ? "" : "s") : dstyle == 1 ? "ms" : ""));
		}
		cell.w = dstr.join(dstyle == 0 ? ":" : " "); cell.z = zstr.join(dstyle == 0 ? '":"': " ");
		if(dstyle == 0) cell.w = cell.w.replace(/:(\d\d\d)$/, ".$1");
	}
}

/** Parse "old storage" (version 0..4) */
function parse_old_storage(buf: Uint8Array, lut: DataLUT, v: 0|1|2|3|4): CellObject | void {
	var dv = u8_to_dataview(buf);
	var flags = dv.getUint32(4, true);

	var ridx = -1, sidx = -1, zidx = -1, ieee = NaN, dt = new Date(2001, 0, 1);
	var doff = (v > 1 ? 12 : 8);
	if(flags & 0x0002) { zidx = dv.getUint32(doff,  true); doff += 4;}
	doff += popcnt(flags & (v > 1 ? 0x0D8C : 0x018C)) * 4;

	if(flags & 0x0200) { ridx = dv.getUint32(doff,  true); doff += 4; }
	doff += popcnt(flags & (v > 1 ? 0x3000 : 0x1000)) * 4;
	if(flags & 0x0010) { sidx = dv.getUint32(doff,  true); doff += 4; }
	if(flags & 0x0020) { ieee = dv.getFloat64(doff, true); doff += 8; }
	if(flags & 0x0040) { dt.setTime(dt.getTime() +  dv.getFloat64(doff, true) * 1000); doff += 8; }

	if(v > 1) {
		flags = dv.getUint32(8, true) >>> 16;
		/* TODO: stress test if a cell can have multiple sub-type formats */
		if(flags & 0xFF) { if(zidx == -1) zidx = dv.getUint32(doff, true); doff += 4; }
	}

	var ret: CellObject;
	var t = buf[v >= 4 ? 1 : 2];
	switch(t) {
		case 0: return void 0; // return { t: "z" }; // blank?
		case 2: ret = { t: "n", v: ieee }; break; // number
		case 3: ret = { t: "s", v: lut.sst[sidx] }; break; // string
		case 5: ret = { t: "d", v: dt }; break; // date-time
		case 6: ret = { t: "b", v: ieee > 0 }; break; // boolean
		case 7: ret = { t: "n", v: ieee }; break; // duration in seconds
		case 8: ret = { t: "e", v: 0}; break; // "formula error" TODO: enumerate and map errors to csf equivalents
		case 9: { // "rich text"
			if(ridx > -1) ret = { t: "s", v: lut.rsst[ridx] };
			else throw new Error(`Unsupported cell type ${buf[subarray](0,4)}`);
		} break;
		default: throw new Error(`Unsupported cell type ${buf[subarray](0,4)}`);
	}

	if(zidx > -1) numbers_format_cell(ret, t | (v<<8), flags, lut.ofmt[zidx], lut.nfmt[zidx]);
	if(t == 7) (ret.v as number) /= 86400;
	return ret;
}

/** Parse "new storage" (version 5) */
function parse_new_storage(buf: Uint8Array, lut: DataLUT): CellObject | void {
	var dv = u8_to_dataview(buf);
	// TODO: bytes 2:3 appear to be unused?
	var flags = dv.getUint32(4, true);
	var fields = dv.getUint32(8, true);
	var doff = 12;

	var ridx = -1, sidx = -1, zidx = -1, d128 = NaN, ieee = NaN, dt = new Date(2001, 0, 1);

	//          0x00001F data
	if(fields & 0x000001) { d128 = readDecimal128LE(buf, doff); doff += 16; }
	if(fields & 0x000002) { ieee = dv.getFloat64(doff, true); doff += 8; }
	if(fields & 0x000004) { dt.setTime(dt.getTime() +  dv.getFloat64(doff, true) * 1000); doff += 8; }
	if(fields & 0x000008) { sidx = dv.getUint32(doff,  true); doff += 4; }
	if(fields & 0x000010) { ridx = dv.getUint32(doff,  true); doff += 4; }

	var ret: CellObject;
	var t = buf[1];
	switch(t) {
		case  0: return void 0; // return { t: "z" }; // blank?
		case  2: ret = { t: "n", v: d128 }; break; // number
		case  3: ret = { t: "s", v: lut.sst[sidx] }; break; // string
		case  5: ret = { t: "d", v: dt }; break; // date-time
		case  6: ret = { t: "b", v: ieee > 0 }; break; // boolean
		case  7: ret = { t: "n", v: ieee }; break;  // duration in "s", fixed later
		case  8: ret = { t: "e", v: 0 }; break; // "formula error" TODO: enumerate and map errors to csf equivalents
		case  9: ret = { t: "s", v: lut.rsst[ridx] }; break;// "rich text"
		case 10: ret = { t: "n", v: d128 }; break; // currency
		default: throw new Error(`Unsupported cell type ${buf[1]} : ${fields & 0x1F} : ${buf[subarray](0,4)}`);
	}

	//          0x0001E0 styling

	//          0x000E00 formula

	//          0x001000 something related to cell format
	doff += popcnt(fields & 0x001FE0) * 4;

	/* TODO: stress test if a cell can have multiple sub-type formats */
	//          0x07E000 formats
	if(fields & 0x07E000) { if(zidx == -1) zidx = dv.getUint32(doff, true); doff += 4; }

	//          0x080000 comment
	//          0x100000 warning

	if(zidx > -1) numbers_format_cell(ret, t | (5<<8), fields >> 13, lut.ofmt[zidx], lut.nfmt[zidx] );
	if(t == 7) (ret.v as number) /= 86400; // duration -> SheetJS absolute time
	return ret;
}

/** Write a cell "new storage" (version 5) */
function write_new_storage(cell: CellObject, sst: string[]): Uint8Array {
	var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
	out[0] = 5;
	switch(cell.t) {
		case "n": out[1] = 2; writeDecimal128LE(out, l, cell.v as number); flags |= 1; l += 16; break;
		case "b": out[1] = 6; dv.setFloat64(l, cell.v ? 1 : 0, true); flags |= 2; l += 8; break;
		case "s":
			var s = cell.v == null ? "" : String(cell.v);
			var isst = sst.indexOf(s);
			if(isst == -1) sst[isst = sst.length] = s;
			out[1] = 3; dv.setUint32(l, isst, true); flags |= 8; l += 4; break;
		default: throw "unsupported cell type " + cell.t;
	}
	dv.setUint32(8, flags, true);
	return out[subarray](0, l);
}
/** Write a cell "old storage" (version 4) */
function write_old_storage(cell: CellObject, sst: string[]): Uint8Array {
	var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
	out[0] = 4;
	switch(cell.t) {
		case "n": out[2] = 2; dv.setFloat64(l, cell.v as number, true); flags |= 0x20; l += 8; break;
		case "b": out[2] = 6; dv.setFloat64(l, cell.v ? 1 : 0, true); flags |= 0x20; l += 8; break;
		case "s":
			var s = cell.v == null ? "" : String(cell.v);
			var isst = sst.indexOf(s);
			if(isst == -1) sst[isst = sst.length] = s;
			out[2] = 3; dv.setUint32(l, isst, true); flags |= 0x10; l += 4; break;
		default: throw "unsupported cell type " + cell.t;
	}
	dv.setUint32(8, flags, true);
	return out[subarray](0, l);
}
//<<export { write_new_storage, write_old_storage };
function parse_cell_storage(buf: Uint8Array, lut: DataLUT): CellObject | void {
	switch(buf[0]) {
		case 0: case 1:
		case 2: case 3: case 4: return parse_old_storage(buf, lut, buf[0]);
		case 5: return parse_new_storage(buf, lut);
		default: throw new Error(`Unsupported payload version ${buf[0]}`);
	}
}

/** .TSS.StylesheetArchive */
//function parse_TSS_StylesheetArchive(M: IWAMessage[][], root: IWAMessage): void {
//	var pb = parse_shallow(root.data);
//}

/** Parse .TSP.Reference */
function parse_TSP_Reference(buf: Uint8Array): number {
	var pb = parse_shallow(buf);
	return parse_varint49(pb[1][0].data);
}
/** Write .TSP.Reference */
function write_TSP_Reference(idx: number): Uint8Array {
	return write_shallow([
		[],
		[ { type: 0, data: write_varint49(idx) } ]
	]);
}
//<<export { parse_TSP_Reference, write_TSP_Reference };

/** Insert Object Reference */
function numbers_add_oref(iwa: IWAArchiveInfo, ref: number): void {
	var orefs: number[] = iwa.messages[0].meta[5]?.[0] ? parse_packed_varints(iwa.messages[0].meta[5][0].data) : [];
	var orefidx = orefs.indexOf(ref);
	if(orefidx == -1) {
		orefs.push(ref);
		iwa.messages[0].meta[5] =[ {type: 2, data: write_packed_varints(orefs) }];
	}
}
/** Delete Object Reference */
function numbers_del_oref(iwa: IWAArchiveInfo, ref: number): void {
	var orefs: number[] = iwa.messages[0].meta[5]?.[0] ? parse_packed_varints(iwa.messages[0].meta[5][0].data) : [];
	iwa.messages[0].meta[5] =[ {type: 2, data: write_packed_varints(orefs.filter(r => r != ref)) }];
}

type MessageSpace = {[id: number]: IWAMessage[]};

/** Parse .TST.TableDataList */
function parse_TST_TableDataList(M: MessageSpace, root: IWAMessage): any[] {
	var pb = parse_shallow(root.data);
	// .TST.TableDataList.ListType
	var type = varint_to_i32(pb[1][0].data);

	var entries = pb[3];
	var data: any[] = [];
	(entries||[]).forEach(entry => {
		// .TST.TableDataList.ListEntry
		var le = parse_shallow(entry.data);
		if(!le[1]) return; // sometimes the list has a spurious entry at index 0
		var key = varint_to_i32(le[1][0].data)>>>0;
		switch(type) {
			case 1: data[key] = u8str(le[3][0].data); break;
			case 8: {
				// .TSP.RichTextPayloadArchive
				var rt = M[parse_TSP_Reference(le[9][0].data)][0];
				var rtp = parse_shallow(rt.data);

				// .TSWP.StorageArchive
				var rtpref = M[parse_TSP_Reference(rtp[1][0].data)][0];
				var mtype = varint_to_i32(rtpref.meta[1][0].data);
				if(mtype != 2001) throw new Error(`2000 unexpected reference to ${mtype}`);
				var tswpsa = parse_shallow(rtpref.data);

				data[key] = tswpsa[3].map(x => u8str(x.data)).join("");
			} break;
			case 2: data[key] = parse_shallow(le[6][0].data); break;
			default: throw type;
		}
	});
	return data;
}

type TileStorageType = -1 | 0 | 1;
interface TileRowInfo {
	/** Row Index */
	R: number;
	/** Cell Storage */
	cells: Uint8Array[];
}
/** Parse .TSP.TileRowInfo */
function parse_TST_TileRowInfo(u8: Uint8Array, type: TileStorageType): TileRowInfo {
	var pb = parse_shallow(u8);
	var R = varint_to_i32(pb[1][0].data) >>> 0;
	var cnt = varint_to_i32(pb[2][0].data) >>> 0;
	// var version = pb?.[5]?.[0] && (varint_to_i32(pb[5][0].data) >>> 0);
	var wide_offsets = pb[8]?.[0]?.data && varint_to_i32(pb[8][0].data) > 0 || false;

	/* select storage by type (1 = modern / 0 = old / -1 = try modern, old) */
	var used_storage_u8: Uint8Array, used_storage: Uint8Array;
	if(pb[7]?.[0]?.data && type != 0) { used_storage_u8 = pb[7]?.[0]?.data; used_storage = pb[6]?.[0]?.data; }
	else if(pb[4]?.[0]?.data && type != 1) { used_storage_u8 = pb[4]?.[0]?.data; used_storage = pb[3]?.[0]?.data; }
	else throw `NUMBERS Tile missing ${type} cell storage`;

	/* find all offsets -- 0xFFFF means cells are not present */
	var width = wide_offsets ? 4 : 1;
	var used_storage_offsets = u8_to_dataview(used_storage_u8);
	var offsets: Array<[number, number]> = [];
	for(var C = 0; C < used_storage_u8.length / 2; ++C) {
		var off = used_storage_offsets.getUint16(C*2, true);
		if(off < 65535) offsets.push([C, off]);
	}
	if(offsets.length != cnt) throw `Expected ${cnt} cells, found ${offsets.length}`;

	var cells: Uint8Array[] = [];
	for(C = 0; C < offsets.length - 1; ++C) cells[offsets[C][0]] = used_storage[subarray](offsets[C][1] * width, offsets[C+1][1] * width);
	if(offsets.length >= 1) cells[offsets[offsets.length - 1][0]] = used_storage[subarray](offsets[offsets.length - 1][1] * width);
	return { R, cells };
}

interface TileInfo {
	data: Uint8Array[][];
	nrows: number;
}
/** Parse .TST.Tile */
function parse_TST_Tile(M: MessageSpace, root: IWAMessage): TileInfo {
	var pb = parse_shallow(root.data);
	// ESBuild issue 2136
	// 	var storage: TileStorageType = (pb?.[7]?.[0]) ? ((varint_to_i32(pb[7][0].data)>>>0) > 0 ? 1 : 0 ) : -1;
	var storage: TileStorageType = -1;
	if(pb?.[7]?.[0]) { if(varint_to_i32(pb[7][0].data)>>>0) storage = 1; else storage = 0; }
	var ri = mappa(pb[5], (u8: Uint8Array) => parse_TST_TileRowInfo(u8, storage));
	return {
		nrows: varint_to_i32(pb[4][0].data)>>>0,
		data: ri.reduce((acc, x) => {
			if(!acc[x.R]) acc[x.R] = [];
			x.cells.forEach((cell, C) => {
				if(acc[x.R][C]) throw new Error(`Duplicate cell r=${x.R} c=${C}`);
				acc[x.R][C] = cell;
			});
			return acc;
		}, [] as Uint8Array[][])
	};
}

/** Parse .TST.TableModelArchive (6001) */
function parse_TST_TableModelArchive(M: MessageSpace, root: IWAMessage, ws: WorkSheet) {
	var pb = parse_shallow(root.data);
	var range: Range = { s: {r:0, c:0}, e: {r:0, c:0} };
	range.e.r = (varint_to_i32(pb[6][0].data) >>> 0) - 1;
	if(range.e.r < 0) throw new Error(`Invalid row varint ${pb[6][0].data}`);
	range.e.c = (varint_to_i32(pb[7][0].data) >>> 0) - 1;
	if(range.e.c < 0) throw new Error(`Invalid col varint ${pb[7][0].data}`);
	ws["!ref"] = encode_range(range);
	var dense = Array.isArray(ws);
	// .TST.DataStore
	var store = parse_shallow(pb[4][0].data);
	var lut: DataLUT = numbers_lut_new();
	if(store[4]?.[0]) lut.sst = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[4][0].data)][0]);
	if(store[11]?.[0]) lut.ofmt = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[11][0].data)][0]);
	if(store[17]?.[0]) lut.rsst = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[17][0].data)][0]);
	if(store[22]?.[0]) lut.nfmt = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[22][0].data)][0]);

	// .TST.TileStorage
	var tile = parse_shallow(store[3][0].data);
	var _R = 0;
	/* TODO: should this list be sorted by id ? */
	tile[1].forEach(t => {
		var tl = (parse_shallow(t.data));
		// var id = varint_to_i32(tl[1][0].data);
		var ref = M[parse_TSP_Reference(tl[2][0].data)][0];
		var mtype = varint_to_i32(ref.meta[1][0].data);
		if(mtype != 6002) throw new Error(`6001 unexpected reference to ${mtype}`);
		var _tile = parse_TST_Tile(M, ref);
		_tile.data.forEach((row, R) => {
			row.forEach((buf, C) => {
				var res = parse_cell_storage(buf, lut);
				if(res) {
					if(dense) {
						if(!ws[_R + R]) ws[_R + R] = [];
						ws[_R + R][C] = res;
					} else {
						var addr = encode_cell({r:_R + R,c:C});
						ws[addr] = res;
					}
				}
			});
		});
		_R += _tile.nrows;
	});

	if(store[13]?.[0]) {
		var ref = M[parse_TSP_Reference(store[13][0].data)][0];
		var mtype = varint_to_i32(ref.meta[1][0].data);
		if(mtype != 6144) throw new Error(`Expected merge type 6144, found ${mtype}`);
		ws["!merges"] = parse_shallow(ref.data)?.[1].map(pi => {
			var merge = parse_shallow(pi.data);
			var origin = u8_to_dataview(parse_shallow(merge[1][0].data)[1][0].data), size = u8_to_dataview(parse_shallow(merge[2][0].data)[1][0].data);
			return {
				s: { r: origin.getUint16(0, true), c: origin.getUint16(2, true) },
				e: {
					r: origin.getUint16(0, true) + size.getUint16(0, true) - 1,
					c: origin.getUint16(2, true) + size.getUint16(2, true) - 1
				}
			};
		});
	}
}

/** Parse .TST.TableInfoArchive (6000) */
function parse_TST_TableInfoArchive(M: MessageSpace, root: IWAMessage, opts?: ParsingOptions): WorkSheet {
	var pb = parse_shallow(root.data);
	// ESBuild #2375
	var out: WorkSheet;
	if(!opts?.dense) out =  ({ "!ref": "A1" });
	else out = ([] as any);
	out["!ref"] = "A1";
	var tableref = M[parse_TSP_Reference(pb[2][0].data)];
	var mtype = varint_to_i32(tableref[0].meta[1][0].data);
	if(mtype != 6001) throw new Error(`6000 unexpected reference to ${mtype}`);
	parse_TST_TableModelArchive(M, tableref[0], out);
	return out;
}

interface NSheet {
	name: string;
	sheets: WorkSheet[];
}
/** Parse .TN.SheetArchive (2) */
function parse_TN_SheetArchive(M: MessageSpace, root: IWAMessage, opts?: ParsingOptions): NSheet {
	var pb = parse_shallow(root.data);
	var out: NSheet = {
		name: (pb[1]?.[0] ? u8str(pb[1][0].data) : ""),
		sheets: []
	};
	var shapeoffs = mappa(pb[2], parse_TSP_Reference);
	shapeoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 6000) out.sheets.push(parse_TST_TableInfoArchive(M, m, opts));
		});
	});
	return out;
}

/** Parse .TN.DocumentArchive */
function parse_TN_DocumentArchive(M: MessageSpace, root: IWAMessage, opts?: ParsingOptions): WorkBook {
	var out = book_new();
	var pb = parse_shallow(root.data);
	if(pb[2]?.[0]) throw new Error("Keynote presentations are not supported");

	// var stylesheet = mappa(pb[4], parse_TSP_Reference)[0];
	// if(varint_to_i32(M[stylesheet][0].meta[1][0].data) == 401) parse_TSS_StylesheetArchive(M, M[stylesheet][0]);
	// var sidebar = mappa(pb[5], parse_TSP_Reference);
	// var theme = mappa(pb[6], parse_TSP_Reference);
	// var docinfo = parse_shallow(pb[8][0].data);
	// var tskinfo = parse_shallow(docinfo[1][0].data);
	// var author_storage = mappa(tskinfo[7], parse_TSP_Reference);

	var sheetoffs = mappa(pb[1], parse_TSP_Reference);
	sheetoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 2) {
				var root = parse_TN_SheetArchive(M, m, opts);
				root.sheets.forEach((sheet, idx) => { book_append_sheet(out, sheet, idx == 0 ? root.name : root.name + "_" + idx, true); });
			}
		});
	});
	if(out.SheetNames.length == 0) throw new Error("Empty NUMBERS file");
	out.bookType = "numbers";
	return out;
}

/** Parse NUMBERS file */
function parse_numbers_iwa(cfb: CFB$Container, opts?: ParsingOptions ): WorkBook {
	var M: MessageSpace = {}, indices: number[] = [];
	cfb.FullPaths.forEach(p => { if(p.match(/\.iwpv2/)) throw new Error(`Unsupported password protection`); });

	/* collect entire message space */
	cfb.FileIndex.forEach(s => {
		if(!s.name.match(/\.iwa$/)) return;
		if(s.content[0] == 98) return; // TODO: OperationStorage.iwa
		var o: Uint8Array;
		try { o = decompress_iwa_file(s.content as Uint8Array); } catch(e) { return console.log("?? " + s.content.length + " " + (e.message || e)); }
		var packets: IWAArchiveInfo[];
		try { packets = parse_iwa_file(o); } catch(e) { return console.log("## " + (e.message || e)); }
		packets.forEach(packet => { M[packet.id] = packet.messages; indices.push(packet.id); });
	});
	if(!indices.length) throw new Error("File has no messages");

	/* find document root */
	if(M?.[1]?.[0].meta?.[1]?.[0].data && varint_to_i32(M[1][0].meta[1][0].data) == 10000) throw new Error("Pages documents are not supported");
	var docroot: IWAMessage | false = M?.[1]?.[0]?.meta?.[1]?.[0].data && varint_to_i32(M[1][0].meta[1][0].data) == 1 && M[1][0];
	if(!docroot) indices.forEach((idx) => {
		M[idx].forEach((iwam) => {
			var mtype = varint_to_i32(iwam.meta[1][0].data) >>> 0;
			if(mtype == 1) {
				if(!docroot) docroot = iwam;
				else throw new Error("Document has multiple roots");
			}
		});
	});
	if(!docroot) throw new Error("Cannot find Document root");
	return parse_TN_DocumentArchive(M, docroot, opts);
}

//<<export { parse_numbers_iwa };

interface DependentInfo {
	deps: number[];
	location: string;
	type: number;
}
/** Write .TST.TileRowInfo */
function write_TST_TileRowInfo(data: any[], SST: string[], wide: boolean): ProtoMessage {
	var tri: ProtoMessage = [
		[],
		[ { type: 0, data: write_varint49(0) }],
		[ { type: 0, data: write_varint49(0) }],
		[ { type: 2, data: new Uint8Array([]) }],
		[ { type: 2, data: new Uint8Array(Array.from({length:510}, () => 255)) }],
		[ { type: 0, data: write_varint49(5) }],
		[ { type: 2, data: new Uint8Array([]) }],
		[ { type: 2, data: new Uint8Array(Array.from({length:510}, () => 255)) }],
		[ { type: 0, data: write_varint49(1) }],
	] as ProtoMessage;
	if(!tri[6]?.[0] || !tri[7]?.[0]) throw "Mutation only works on post-BNC storages!";
	//var wide_offsets = tri[8]?.[0]?.data && varint_to_i32(tri[8][0].data) > 0 || false;
	var cnt = 0;
	if(tri[7][0].data.length < 2 * data.length) {
		var new_7 = new Uint8Array(2 * data.length);
		new_7.set(tri[7][0].data);
		tri[7][0].data = new_7;
	}
	//if(wide) {
	//	tri[3] = [{type: 2, data: new Uint8Array([240, 159, 164, 160]) }];
	//	tri[4] = [{type: 2, data: new Uint8Array([240, 159, 164, 160]) }];
	/* } else*/ if(tri[4][0].data.length < 2 * data.length) {
		var new_4 = new Uint8Array(2 * data.length);
		new_4.set(tri[4][0].data);
		tri[4][0].data = new_4;
	}
	var dv = u8_to_dataview(tri[7][0].data), last_offset = 0, cell_storage: Uint8Array[] = [];
	var _dv = u8_to_dataview(tri[4][0].data), _last_offset = 0, _cell_storage: Uint8Array[] = [];
	var width = wide ? 4 : 1;
	for(var C = 0; C < data.length; ++C) {
		if(data[C] == null) { dv.setUint16(C*2, 0xFFFF, true); _dv.setUint16(C*2, 0xFFFF); continue; }
		dv.setUint16(C*2, last_offset / width, true);
		/*if(!wide)*/ _dv.setUint16(C*2, _last_offset / width, true);
		var celload: Uint8Array, _celload: Uint8Array;
		switch(typeof data[C]) {
			case "string":
				celload = write_new_storage({t: "s", v: data[C]}, SST);
				/*if(!wide)*/ _celload = write_old_storage({t: "s", v: data[C]}, SST);
				break;
			case "number":
				celload = write_new_storage({t: "n", v: data[C]}, SST);
				/*if(!wide)*/ _celload = write_old_storage({t: "n", v: data[C]}, SST);
				break;
			case "boolean":
				celload = write_new_storage({t: "b", v: data[C]}, SST);
				/*if(!wide)*/ _celload = write_old_storage({t: "b", v: data[C]}, SST);
				break;
			default:
				// TODO: write the actual date code
				if(data[C] instanceof Date) {
					celload = write_new_storage({t: "s", v: (data[C] as Date).toISOString()}, SST);
					/*if(!wide)*/ _celload = write_old_storage({t: "s", v: (data[C] as Date).toISOString()}, SST);
					break;
				}
				throw new Error("Unsupported value " + data[C]);
		}
		cell_storage.push(celload); last_offset += celload.length;
		/*if(!wide)*/ { _cell_storage.push(_celload); _last_offset += _celload.length; }
		++cnt;
	}
	tri[2][0].data = write_varint49(cnt);
	tri[5][0].data = write_varint49(5);

	for(; C < tri[7][0].data.length/2; ++C) {
		dv.setUint16(C*2, 0xFFFF, true);
		/*if(!wide)*/ _dv.setUint16(C*2, 0xFFFF, true);
	}
	tri[6][0].data = u8concat(cell_storage);
	/*if(!wide)*/ tri[3][0].data = u8concat(_cell_storage);
	tri[8] = [{type: 0, data: write_varint49(wide ? 1 : 0)}];
	return tri;
}

/** Write IWA Message */
function write_iwam(type: number, payload: Uint8Array): IWAMessage {
	return {
		meta: [ [],
			[ { type: 0, data: write_varint49(type) } ],
			// [ { type: 2, data: new Uint8Array([1, 0, 5]) }]
		],
		data: payload
	};
}

type Dependents = {[x:number]: DependentInfo; last?: number;};

function get_unique_msgid(dep: DependentInfo, dependents: Dependents) {
	if(!dependents.last) dependents.last = 927262;
	for(var i = dependents.last; i < 2000000; ++i) if(!dependents[i]) {
		dependents[dependents.last = i] = dep;
		return i;
	}
	throw new Error("Too many messages");
}

/** Build an approximate dependency tree */
function build_numbers_deps(cfb: CFB$Container): Dependents {
	var dependents: Dependents = {};
	var indices: number[] = [];

	cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
		var fi = row[0], fp = row[1];
		if(fi.type != 2) return;
		if(!fi.name.match(/\.iwa/)) return;
		if(fi.name.match(/OperationStorage/)) return;

		parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array)).forEach(packet => {
			indices.push(packet.id);
			dependents[packet.id] = { deps: [], location: fp, type: varint_to_i32(packet.messages[0].meta[1][0].data) };
		});
	});

	/* precompute a varint for each id */
	indices.sort((x,y) => x-y);
	var indices_varint: Array<[number, Uint8Array]> = indices.filter(x => x > 1).map(x => [x, write_varint49(x)] );

	/* build dependent tree */
	cfb.FileIndex.forEach(fi => {
		if(!fi.name.match(/\.iwa/)) return;
		if(fi.name.match(/OperationStorage/)) return;
		parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array)).forEach(ia => {
			// this is a huge hack based on the observation that most messages of interest have id > 900000
			// TODO: use the actual references
			indices_varint.forEach(ivi => {
				if(ia.messages.some(mess => varint_to_i32(mess.meta[1][0].data) != 11006 && u8contains(mess.data, ivi[1]))) {
					dependents[ivi[0]].deps.push(ia.id);
				}
			});
		});
	});

	return dependents;
}

/** Write NUMBERS workbook */
function write_numbers_iwa(wb: WorkBook, opts?: WritingOptions): CFB$Container {
	if(!opts || !opts.numbers) throw new Error("Must pass a `numbers` option -- check the README");

	/* read template and build packet metadata */
	var cfb: CFB$Container = CFB.read(opts.numbers, { type: "base64" });
	var deps: Dependents = build_numbers_deps(cfb);

	/* .TN.DocumentArchive */
	var docroot: IWAArchiveInfo = numbers_iwa_find(cfb, deps, 1);
	if(docroot == null) throw `Could not find message ${1} in Numbers template`;
	var sheetrefs = mappa(parse_shallow(docroot.messages[0].data)[1], parse_TSP_Reference);
	if(sheetrefs.length > 1) throw new Error("Template NUMBERS file must have exactly one sheet")
	wb.SheetNames.forEach((name, idx) => {
		if(idx >= 1) {
			numbers_add_ws(cfb, deps, idx + 1);
			docroot = numbers_iwa_find(cfb, deps, 1);
			sheetrefs = mappa(parse_shallow(docroot.messages[0].data)[1], parse_TSP_Reference);
		}
		write_numbers_ws(cfb, deps, wb.Sheets[name], name, idx, sheetrefs[idx])
	});
	return cfb;
}

/** Find a particular message by ID, perform actions and commit */
function numbers_iwa_doit(cfb: CFB$Container, deps: Dependents, id: number, cb:(ai:IWAArchiveInfo, x:IWAArchiveInfo[])=>void) {
	var entry = CFB.find(cfb, deps[id].location);
	if(!entry) throw `Could not find ${deps[id].location} in Numbers template`;
	var x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
	var ainfo: IWAArchiveInfo = x.find(packet => packet.id == id) as IWAArchiveInfo;
	// TODO: it's assumed this exists
	cb(ainfo, x);
	entry.content = compress_iwa_file(write_iwa_file(x)); entry.size = entry.content.length;
}

/** Find a particular message by ID */
function numbers_iwa_find(cfb: CFB$Container, deps: Dependents, id: number) {
	var entry = CFB.find(cfb, deps[id].location);
	if(!entry) throw `Could not find ${deps[id].location} in Numbers template`;
	var x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
	var ainfo: IWAArchiveInfo = x.find(packet => packet.id == id) as IWAArchiveInfo;
	// TODO: it's assumed this exists
	return ainfo;
}

/** Deep copy of the essential parts of a worksheet */
function numbers_add_ws(cfb: CFB$Container, deps: Dependents, wsidx: number) {
	var sheetref = -1, newsheetref = -1;

	var remap: {[x: number]: number} = {};
	/* .TN.DocumentArchive -> new .TN.SheetArchive */
	numbers_iwa_doit(cfb, deps, 1, (docroot: IWAArchiveInfo, arch: IWAArchiveInfo[]) => {
		var doc = parse_shallow(docroot.messages[0].data);

		/* new sheet reference */
		sheetref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[1][0].data);
		newsheetref = get_unique_msgid({ deps: [1], location: deps[sheetref].location, type: 2 }, deps);
		remap[sheetref] = newsheetref;

		/* connect root -> sheet */
		numbers_add_oref(docroot, newsheetref);
		doc[1].push({ type: 2, data: write_TSP_Reference(newsheetref) });

		/* copy sheet */
		var sheet = numbers_iwa_find(cfb, deps, sheetref);
		sheet.id = newsheetref;
		if(deps[1].location == deps[newsheetref].location) arch.push(sheet);
		else numbers_iwa_doit(cfb, deps, newsheetref, (_, x) => x.push(sheet));

		docroot.messages[0].data = write_shallow(doc);
	});

	/* .TN.SheetArchive -> new .TST.TableInfoArchive */
	var tiaref = -1;
	numbers_iwa_doit(cfb, deps, newsheetref, (sheetroot: IWAArchiveInfo, arch: IWAArchiveInfo[]) => {
		var sa = parse_shallow(sheetroot.messages[0].data);

		/* remove everything except for name and drawables */
		for(var i = 3; i <= 69; ++i) delete sa[i];
		/* remove all references to drawables */
		var drawables = mappa(sa[2], parse_TSP_Reference);
		drawables.forEach(n => numbers_del_oref(sheetroot, n));

		/* add new tia reference */
		tiaref = get_unique_msgid({ deps: [newsheetref], location: deps[drawables[0]].location, type: deps[drawables[0]].type }, deps);
		numbers_add_oref(sheetroot, tiaref);
		remap[drawables[0]] = tiaref;
		sa[2] = [ { type: 2, data: write_TSP_Reference(tiaref) } ];

		/* copy tia */
		var tia = numbers_iwa_find(cfb, deps, drawables[0]);
		tia.id = tiaref;
		if(deps[drawables[0]].location == deps[newsheetref].location) arch.push(tia);
		else {
			var loc = deps[newsheetref].location;
			loc = loc.replace(/^Root Entry\//,""); // NOTE: the Root Entry prefix is an artifact of the CFB container library
			loc = loc.replace(/^Index\//, "").replace(/\.iwa$/,"");
			numbers_iwa_doit(cfb, deps, 2, (ai => {
				var mlist = parse_shallow(ai.messages[0].data);

				/* add reference from SheetArchive file to TIA */
				var parentidx = mlist[3].findIndex(m => {
					var mm = parse_shallow(m.data);
					if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
					if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
					return false;
				});
				var parent = parse_shallow(mlist[3][parentidx].data);
				if(!parent[6]) parent[6] = [];
				parent[6].push({
					type: 2,
					data: write_shallow([
						[],
						[{type: 0, data: write_varint49(tiaref) }]
					])
				});
				mlist[3][parentidx].data = write_shallow(parent);

				ai.messages[0].data = write_shallow(mlist);
			}));
			numbers_iwa_doit(cfb, deps, tiaref, (_, x) => x.push(tia));
		}

		sheetroot.messages[0].data = write_shallow(sa);
	});

	/* .TST.TableInfoArchive -> new .TST.TableModelArchive */
	var tmaref = -1;
	numbers_iwa_doit(cfb, deps, tiaref, (tiaroot: IWAArchiveInfo, arch: IWAArchiveInfo[]) => {
		var tia = parse_shallow(tiaroot.messages[0].data);

		/* update parent reference */
		var da = parse_shallow(tia[1][0].data);
		for(var i = 3; i <= 69; ++i) delete da[i];
		var dap = parse_TSP_Reference(da[2][0].data);
		da[2][0].data = write_TSP_Reference(remap[dap]);
		tia[1][0].data = write_shallow(da);

		/* remove old tma reference */
		var oldtmaref = parse_TSP_Reference(tia[2][0].data);
		numbers_del_oref(tiaroot, oldtmaref);

		/* add new tma reference */
		tmaref = get_unique_msgid({ deps: [tiaref], location: deps[oldtmaref].location, type: deps[oldtmaref].type }, deps);
		numbers_add_oref(tiaroot, tmaref);
		remap[oldtmaref] = tmaref;
		tia[2][0].data = write_TSP_Reference(tmaref);

		/* copy tma */
		var tma = numbers_iwa_find(cfb, deps, oldtmaref);
		tma.id = tmaref;
		if(deps[tiaref].location == deps[tmaref].location) arch.push(tma);
		else numbers_iwa_doit(cfb, deps, tmaref, (_, x) => x.push(tma));

		tiaroot.messages[0].data = write_shallow(tia);
	});

	/* identifier for finding the TableModelArchive in the archive */
	var loc = deps[tmaref].location;
	loc = loc.replace(/^Root Entry\//,""); // NOTE: the Root Entry prefix is an artifact of the CFB container library
	loc = loc.replace(/^Index\//, "").replace(/\.iwa$/,"");

	/* .TST.TableModelArchive */
	numbers_iwa_doit(cfb, deps, tmaref, (tmaroot: IWAArchiveInfo, arch: IWAArchiveInfo[]) => {
		/* TODO: formulae currently break due to missing CE details */
		var tma = parse_shallow(tmaroot.messages[0].data);
		var uuid = u8str(tma[1][0].data), new_uuid = uuid.replace(/-[A-Z0-9]*/, `-${wsidx.toString(16).padStart(4, "0")}`);
		tma[1][0].data = stru8(new_uuid);

		/* NOTE: These lists should be revisited every time the template is changed */

		/* bare fields */
		[ 12, 13, 29, 31, 32, 33, 39, 44, 47, 81, 82, 84 ].forEach(n => delete tma[n]);

		if(tma[45]) {
			// .TST.SortRuleReferenceTrackerArchive
			var srrta = parse_shallow(tma[45][0].data);
			var ref = parse_TSP_Reference(srrta[1][0].data);
			numbers_del_oref(tmaroot, ref);
			delete tma[45];
		}

		if(tma[70]) {
			// .TST.HiddenStatesOwnerArchive
			var hsoa = parse_shallow(tma[70][0].data);
			hsoa[2]?.forEach(item => {
				// .TST.HiddenStatesArchive
				var hsa = parse_shallow(item.data);
				[2,3].map(n => hsa[n][0]).forEach(hseadata => {
					// .TST.HiddenStateExtentArchive
					var hsea = parse_shallow(hseadata.data);
					if(!hsea[8]) return;
					var ref = parse_TSP_Reference(hsea[8][0].data);
					numbers_del_oref(tmaroot, ref);
				});
			});
			delete tma[70];
		}

		[ 46, // deleting field 46 (base_column_row_uids) forces Numbers to refresh cell table
			30, 34, 35, 36, 38, 48, 49, 60, 61, 62, 63, 64, 71, 72, 73, 74, 75, 85, 86, 87, 88, 89
		].forEach(n => {
			if(!tma[n]) return;
			var ref = parse_TSP_Reference(tma[n][0].data);
			delete tma[n];
			numbers_del_oref(tmaroot, ref);
		});

		/* update .TST.DataStore */
		var store = parse_shallow(tma[4][0].data);
		{
			/* TODO: actually scan through dep tree and update */

			/* blind copy of single references */
			[ 2, 4, 5, 6, 11, 12, 13, 15, 16, 17, 18, 19, 20, 21, 22 ].forEach(n => {
				if(!store[n]?.[0]) return;
				var oldref = parse_TSP_Reference(store[n][0].data);
				var newref = get_unique_msgid({ deps: [tmaref], location: deps[oldref].location, type: deps[oldref].type }, deps);
				numbers_del_oref(tmaroot, oldref);
				numbers_add_oref(tmaroot, newref);
				remap[oldref] = newref;
				var msg = numbers_iwa_find(cfb, deps, oldref);
				msg.id = newref;
				if(deps[oldref].location == deps[tmaref].location) arch.push(msg);
				else {
					deps[newref].location = deps[oldref].location.replace(oldref.toString(), newref.toString());
					if(deps[newref].location == deps[oldref].location) deps[newref].location = deps[newref].location.replace(/\.iwa/, `-${newref}.iwa`);
					CFB.utils.cfb_add(cfb, deps[newref].location, compress_iwa_file(write_iwa_file([ msg ])));

					var newloc = deps[newref].location;
					newloc = newloc.replace(/^Root Entry\//,""); // NOTE: the Root Entry prefix is an artifact of the CFB container library
					newloc = newloc.replace(/^Index\//, "").replace(/\.iwa$/,"");

					numbers_iwa_doit(cfb, deps, 2, (ai => {
						var mlist = parse_shallow(ai.messages[0].data);
						mlist[3].push({type: 2, data: write_shallow([
							[],
							[{type: 0, data: write_varint49(newref)}],
							[{type: 2, data: stru8(newloc.replace(/-.*$/, "")) }],
							[{type: 2, data: stru8(newloc)}],
							[{type: 2, data: new Uint8Array([2, 0, 0])}],
							[{type: 2, data: new Uint8Array([2, 0, 0])}],
							[],
							[],
							[],
							[],
							[{type: 0, data: write_varint49(0)}],
							[],
							[{type: 0, data: write_varint49(0 /* TODO: save_token */)}],
						])});
						mlist[1] = [{type: 0, data: write_varint49(Math.max(newref + 1, parse_varint49(mlist[1][0].data) ))}];

						/* add reference from TableModelArchive file to Tile */
						var parentidx = mlist[3].findIndex(m => {
							var mm = parse_shallow(m.data);
							if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
							if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
							return false;
						});
						var parent = parse_shallow(mlist[3][parentidx].data);
						if(!parent[6]) parent[6] = [];
						parent[6].push({
							type: 2,
							data: write_shallow([
								[],
								[{type: 0, data: write_varint49(newref) }]
							])
						});
						mlist[3][parentidx].data = write_shallow(parent);

						ai.messages[0].data = write_shallow(mlist);
					}));
				}
				store[n][0].data = write_TSP_Reference(newref);
			});

			/* copy row header storage */
			var row_headers = parse_shallow(store[1][0].data);
			{
				row_headers[2]?.forEach(tspref => {
					var oldref = parse_TSP_Reference(tspref.data);
					var newref = get_unique_msgid({ deps: [tmaref], location: deps[oldref].location, type: deps[oldref].type }, deps);
					numbers_del_oref(tmaroot, oldref);
					numbers_add_oref(tmaroot, newref);
					remap[oldref] = newref;
					var msg = numbers_iwa_find(cfb, deps, oldref);
					msg.id = newref;
					if(deps[oldref].location == deps[tmaref].location) {
						arch.push(msg);
					} else {
						deps[newref].location = deps[oldref].location.replace(oldref.toString(), newref.toString());
						if(deps[newref].location == deps[oldref].location) deps[newref].location = deps[newref].location.replace(/\.iwa/, `-${newref}.iwa`);
						CFB.utils.cfb_add(cfb, deps[newref].location, compress_iwa_file(write_iwa_file([ msg ])));

						var newloc = deps[newref].location;
						newloc = newloc.replace(/^Root Entry\//,""); // NOTE: the Root Entry prefix is an artifact of the CFB container library
						newloc = newloc.replace(/^Index\//, "").replace(/\.iwa$/,"");

						numbers_iwa_doit(cfb, deps, 2, (ai => {
							var mlist = parse_shallow(ai.messages[0].data);
							mlist[3].push({type: 2, data: write_shallow([
								[],
								[{type: 0, data: write_varint49(newref)}],
								[{type: 2, data: stru8(newloc.replace(/-.*$/, "")) }],
								[{type: 2, data: stru8(newloc)}],
								[{type: 2, data: new Uint8Array([2, 0, 0])}],
								[{type: 2, data: new Uint8Array([2, 0, 0])}],
								[],
								[],
								[],
								[],
								[{type: 0, data: write_varint49(0)}],
								[],
								[{type: 0, data: write_varint49(0 /* TODO: save_token */)}],
							])});
							mlist[1] = [{type: 0, data: write_varint49(Math.max(newref + 1, parse_varint49(mlist[1][0].data) ))}];

							/* add reference from TableModelArchive file to Tile */
							var parentidx = mlist[3].findIndex(m => {
								var mm = parse_shallow(m.data);
								if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
								if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
								return false;
							});
							var parent = parse_shallow(mlist[3][parentidx].data);
							if(!parent[6]) parent[6] = [];
							parent[6].push({
								type: 2,
								data: write_shallow([
									[],
									[{type: 0, data: write_varint49(newref) }]
								])
							});
							mlist[3][parentidx].data = write_shallow(parent);

							ai.messages[0].data = write_shallow(mlist);
						}));
					}
					tspref.data = write_TSP_Reference(newref);
				})
			}
			store[1][0].data = write_shallow(row_headers);

			/* copy tiles */
			var tiles = parse_shallow(store[3][0].data);
			{
				tiles[1].forEach(t => {
					var tst = parse_shallow(t.data);
					var oldtileref = parse_TSP_Reference(tst[2][0].data);
					var newtileref = remap[oldtileref];
					if(!remap[oldtileref]) {
						newtileref = get_unique_msgid({ deps: [tmaref], location: "", type: deps[oldtileref].type }, deps);
						deps[newtileref].location = `Root Entry/Index/Tables/Tile-${newtileref}.iwa`;
						remap[oldtileref] = newtileref;

						var oldtile = numbers_iwa_find(cfb, deps, oldtileref);
						oldtile.id = newtileref;
						numbers_del_oref(tmaroot, oldtileref);
						numbers_add_oref(tmaroot, newtileref);

						CFB.utils.cfb_add(cfb, `/Index/Tables/Tile-${newtileref}.iwa`, compress_iwa_file(write_iwa_file([ oldtile ])));

						numbers_iwa_doit(cfb, deps, 2, (ai => {
							var mlist = parse_shallow(ai.messages[0].data);
							mlist[3].push({type: 2, data: write_shallow([
								[],
								[{type: 0, data: write_varint49(newtileref)}],
								[{type: 2, data: stru8("Tables/Tile") }],
								[{type: 2, data: stru8(`Tables/Tile-${newtileref}`)}],
								[{type: 2, data: new Uint8Array([2, 0, 0])}],
								[{type: 2, data: new Uint8Array([2, 0, 0])}],
								[],
								[],
								[],
								[],
								[{type: 0, data: write_varint49(0)}],
								[],
								[{type: 0, data: write_varint49(0 /* TODO: save_token */)}],
							])});
							mlist[1] = [{type: 0, data: write_varint49(Math.max(newtileref + 1, parse_varint49(mlist[1][0].data) ))}];

							/* add reference from TableModelArchive file to Tile */
							var parentidx = mlist[3].findIndex(m => {
								var mm = parse_shallow(m.data);
								if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
								if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
								return false;
							});
							var parent = parse_shallow(mlist[3][parentidx].data);
							if(!parent[6]) parent[6] = [];
							parent[6].push({
								type: 2,
								data: write_shallow([
									[],
									[{type: 0, data: write_varint49(newtileref) }]
								])
							});
							mlist[3][parentidx].data = write_shallow(parent);

							ai.messages[0].data = write_shallow(mlist);
						}));
					}
					tst[2][0].data = write_TSP_Reference(newtileref);
					t.data = write_shallow(tst);
				})
			}
			store[3][0].data = write_shallow(tiles);
		}
		tma[4][0].data = write_shallow(store);
		tmaroot.messages[0].data = write_shallow(tma);
	});
}

/** Write NUMBERS worksheet */
function write_numbers_ws(cfb: CFB$Container, deps: Dependents, ws: WorkSheet, wsname: string, sheetidx: number, rootref: number): void {
	/* TODO: support more data types, etc */

	/* .TN.SheetArchive */
	var drawables: number[] = [];
	numbers_iwa_doit(cfb, deps, rootref, (docroot) => {

		var sheetref = parse_shallow(docroot.messages[0].data);
		{
			/* write worksheet name */
			sheetref[1] = [ { type: 2, data: stru8(wsname) }];
			drawables = mappa(sheetref[2], parse_TSP_Reference);
		}
		docroot.messages[0].data = write_shallow(sheetref);
	});

	/* .TST.TableInfoArchive */
	// TODO: verify that the first entry is actually a table
	// TODO: eventually support multiple tables (and replicate in QPW)
	var tia: IWAArchiveInfo = numbers_iwa_find(cfb, deps, drawables[0]);
	/* .TST.TableModelArchive */
	var tmaref = parse_TSP_Reference(parse_shallow(tia.messages[0].data)[2][0].data);
	numbers_iwa_doit(cfb, deps, tmaref, (docroot, x) => write_numbers_tma(cfb, deps, ws, docroot, x, tmaref));
}

var USE_WIDE_ROWS = true;

/** Write .TST.TableModelArchive */
function write_numbers_tma(cfb: CFB$Container, deps: Dependents, ws: WorkSheet, tmaroot: IWAArchiveInfo, tmafile: IWAArchiveInfo[], tmaref: number) {
	var range = decode_range(ws["!ref"] as string);
	range.s.r = range.s.c = 0;

	var trunc = false;
	/* Actual NUMBERS 12.1 range limit ALL1000000 */
	if(range.e.c > 999) { trunc = true; range.e.c = 999; }
	if(range.e.r > 999999) { trunc = true; range.e.r = 999999; }
	if(trunc) console.error(`Truncating to ${encode_range(range)}`);

	/* preprocess data and build up shared string table */
	var data = sheet_to_json<any>(ws, { range, header: 1 });
	var SST = ["~Sh33tJ5~"];

	/* identifier for finding the TableModelArchive in the archive */
	var loc = deps[tmaref].location;
	loc = loc.replace(/^Root Entry\//,""); // NOTE: the Root Entry prefix is an artifact of the CFB container library
	loc = loc.replace(/^Index\//, "").replace(/\.iwa$/,"");

	var pb = parse_shallow(tmaroot.messages[0].data);
	{
		pb[6][0].data = write_varint49(range.e.r + 1); // number_of_rows
		pb[7][0].data = write_varint49(range.e.c + 1); // number_of_columns
		// pb[22] = [ { type: 0, data: write_varint49(1) } ]; // displays table name in sheet

		delete pb[46]; // base_column_row_uids -- deleting forces Numbers to refresh cell table

		/* .TST.DataStore */
		var store = parse_shallow(pb[4][0].data);
		{
			/* rewrite row headers */
			var row_header_ref = parse_TSP_Reference(parse_shallow(store[1][0].data)[2][0].data);
			numbers_iwa_doit(cfb, deps, row_header_ref, (rowhead, _x) => {
				var base_bucket = parse_shallow(rowhead.messages[0].data);
				if(base_bucket?.[2]?.[0]) for(var R = 0; R < data.length; ++R) {
					var _bucket = parse_shallow(base_bucket[2][0].data);
					_bucket[1][0].data = write_varint49(R);
					_bucket[4][0].data = write_varint49(data[R].length);
					base_bucket[2][R] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
				}
				rowhead.messages[0].data = write_shallow(base_bucket);
			});

			/* rewrite col headers */
			var col_header_ref = parse_TSP_Reference(store[2][0].data);
			numbers_iwa_doit(cfb, deps, col_header_ref, (colhead, _x) => {
				var base_bucket = parse_shallow(colhead.messages[0].data);
				for(var C = 0; C <= range.e.c; ++C) {
					var _bucket = parse_shallow(base_bucket[2][0].data);
					_bucket[1][0].data = write_varint49(C);
					_bucket[4][0].data = write_varint49(range.e.r + 1);
					base_bucket[2][C] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
				}
				colhead.messages[0].data = write_shallow(base_bucket);
			});

			var rbtree = parse_shallow(store[9][0].data);
			rbtree[1] = [];

			/* .TST.TileStorage */
			var tilestore = parse_shallow(store[3][0].data);
			{
				/* number of rows per tile */
				var tstride = 256; // NOTE: if this is not 256, Numbers will assert and recalculate
				tilestore[2] = [{type: 0, data: write_varint49(tstride)}];
				//tilestore[3] = [{type: 0, data: write_varint49(USE_WIDE_ROWS ? 1 : 0)}]; // elicits a modification message

				var tileref = parse_TSP_Reference(parse_shallow(tilestore[1][0].data)[2][0].data);

				/* save the save_token from package metadata */
				var save_token = ((): number => {
					/* .TSP.PackageMetadata */
					var metadata = numbers_iwa_find(cfb, deps, 2);
					var mlist = parse_shallow(metadata.messages[0].data);
					/* .TSP.ComponentInfo field 1 is the id, field 12 is the save token */
					var mlst = mlist[3].filter(m => parse_varint49(parse_shallow(m.data)[1][0].data) == tileref);
					return (mlst?.length) ? parse_varint49(parse_shallow(mlst[0].data)[12][0].data) : 0;
				})();

				/* remove existing tile */
				{
					CFB.utils.cfb_del(cfb, deps[tileref].location);

					/* remove existing tile from reference -- TODO: can this have an id other than 2? */
					numbers_iwa_doit(cfb, deps, 2, (ai => {
						var mlist = parse_shallow(ai.messages[0].data);

						mlist[3] = mlist[3].filter(m => parse_varint49(parse_shallow(m.data)[1][0].data) != tileref);

						/* remove reference from TableModelArchive file to Tile */
						var parentidx = mlist[3].findIndex(m => {
							var mm = parse_shallow(m.data);
							if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
							if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
							return false;
						});
						var parent = parse_shallow(mlist[3][parentidx].data);
						if(!parent[6]) parent[6] = [];
						parent[6] = parent[6].filter(m => parse_varint49(parse_shallow(m.data)[1][0].data) != tileref);
						mlist[3][parentidx].data = write_shallow(parent);

						ai.messages[0].data = write_shallow(mlist);
					}));

					numbers_del_oref(tmaroot, tileref);
				}

				/* rewrite entire tile storage */
				tilestore[1] = [] as ProtoField;
				var ntiles = Math.ceil((range.e.r + 1)/tstride);

				for(var tidx = 0; tidx < ntiles; ++tidx) {
					var newtileid = get_unique_msgid({
						deps: [], // TODO: probably should update this
						location: "",
						type: 6002
					}, deps);
					deps[newtileid].location = `Root Entry/Index/Tables/Tile-${newtileid}.iwa`;

					/* create new tile */
					var tiledata: ProtoMessage = [
						[],
						[{type: 0, data: write_varint49(0 /*range.e.c + 1*/)}],
						[{type: 0, data: write_varint49(Math.min(range.e.r + 1, (tidx + 1) * tstride))}],
						[{type: 0, data: write_varint49(0/*cnt*/)}],
						[{type: 0, data: write_varint49(Math.min((tidx+1)*tstride,range.e.r+1) - tidx * tstride)}],
						[],
						[{type: 0, data: write_varint49(5)}],
						[{type: 0, data: write_varint49(1)}],
						[{type: 0, data: write_varint49(USE_WIDE_ROWS ? 1 : 0)}]
					];
					for(var R = tidx * tstride; R <= Math.min(range.e.r, (tidx + 1) * tstride - 1); ++R) {
						var tilerow = write_TST_TileRowInfo(data[R], SST, USE_WIDE_ROWS);
						tilerow[1][0].data = write_varint49(R - tidx * tstride);
						tiledata[5].push({data: write_shallow(tilerow), type: 2});
					}

					/* add to tiles */
					tilestore[1].push({type: 2, data: write_shallow([
						[],
						[{type: 0, data: write_varint49(tidx)}],
						[{type: 2, data: write_TSP_Reference(newtileid)}]
					])});

					/* add to file */
					var newtile: IWAArchiveInfo = {
						id: newtileid,
						messages: [ write_iwam(6002, write_shallow(tiledata))]
					};
					var tilecontent = compress_iwa_file(write_iwa_file([newtile]));
					CFB.utils.cfb_add(cfb, `/Index/Tables/Tile-${newtileid}.iwa`, tilecontent);

					/* update metadata -- TODO: can this have an id other than 2? */
					numbers_iwa_doit(cfb, deps, 2, (ai => {
						var mlist = parse_shallow(ai.messages[0].data);
						mlist[3].push({type: 2, data: write_shallow([
							[],
							[{type: 0, data: write_varint49(newtileid)}],
							[{type: 2, data: stru8("Tables/Tile") }],
							[{type: 2, data: stru8(`Tables/Tile-${newtileid}`)}],
							[{type: 2, data: new Uint8Array([2, 0, 0])}],
							[{type: 2, data: new Uint8Array([2, 0, 0])}],
							[],
							[],
							[],
							[],
							[{type: 0, data: write_varint49(0)}],
							[],
							[{type: 0, data: write_varint49(save_token)}],
						])});
						mlist[1] = [{type: 0, data: write_varint49(Math.max(newtileid + 1, parse_varint49(mlist[1][0].data) ))}];

						/* add reference from TableModelArchive file to Tile */
						var parentidx = mlist[3].findIndex(m => {
							var mm = parse_shallow(m.data);
							if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
							if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
							return false;
						});
						var parent = parse_shallow(mlist[3][parentidx].data);
						if(!parent[6]) parent[6] = [];
						parent[6].push({
							type: 2,
							data: write_shallow([
								[],
								[{type: 0, data: write_varint49(newtileid) }]
							])
						});
						mlist[3][parentidx].data = write_shallow(parent);

						ai.messages[0].data = write_shallow(mlist);
					}));

					/* add to TableModelArchive object references */
					numbers_add_oref(tmaroot, newtileid);

					/* add to row rbtree */
					rbtree[1].push({type: 2, data: write_shallow([
						[],
						[{ type: 0, data: write_varint49(tidx*tstride) }],
						[{ type: 0, data: write_varint49(tidx) }]
					])});
				}
			}
			store[3][0].data = write_shallow(tilestore);
			store[9][0].data = write_shallow(rbtree);
			store[10] = [ { type: 2, data: new Uint8Array([]) }];

			/* write merge list */
			if(ws["!merges"]) {
				var mergeid = get_unique_msgid({
					type: 6144,
					deps: [tmaref],
					location: deps[tmaref].location
				}, deps);

				tmafile.push({
					id: mergeid,
					messages: [ write_iwam(6144, write_shallow([ [],
						ws["!merges"].map((m: Range) => ({type: 2, data: write_shallow([ [],
							[{ type: 2, data: write_shallow([ [],
								[{ type: 5, data: new Uint8Array(new Uint16Array([m.s.r, m.s.c]).buffer) }],
							])}],
							[{ type: 2, data: write_shallow([ [],
								[{ type: 5, data: new Uint8Array(new Uint16Array([m.e.r - m.s.r + 1, m.e.c - m.s.c + 1]).buffer) }],
							]) }]
						])} as ProtoItem))
					])) ]
				} as IWAArchiveInfo);
				store[13] = [ { type: 2, data: write_TSP_Reference(mergeid) } ];

				numbers_iwa_doit(cfb, deps, 2, (ai => {
					var mlist = parse_shallow(ai.messages[0].data);

					/* add reference from TableModelArchive file to merge */
					var parentidx = mlist[3].findIndex(m => {
						var mm = parse_shallow(m.data);
						if(mm[3]?.[0]) return u8str(mm[3][0].data) == loc;
						if(mm[2]?.[0] && u8str(mm[2][0].data) == loc) return true;
						return false;
					});
					var parent = parse_shallow(mlist[3][parentidx].data);
					if(!parent[6]) parent[6] = [];
					parent[6].push({
						type: 2,
						data: write_shallow([
							[],
							[{type: 0, data: write_varint49(mergeid) }]
						])
					});
					mlist[3][parentidx].data = write_shallow(parent);

					ai.messages[0].data = write_shallow(mlist);
				}));

				/* add object reference from TableModelArchive */
				numbers_add_oref(tmaroot, mergeid);

			} else delete store[13]; // TODO: delete references to merge if not needed

			/* rebuild shared string table */
			var sstref = parse_TSP_Reference(store[4][0].data);
			numbers_iwa_doit(cfb, deps, sstref, (sstroot) => {
				var sstdata = parse_shallow(sstroot.messages[0].data);
				{
					sstdata[3] = [];
					SST.forEach((str, i) => {
						if(i == 0) return; // Numbers will assert if index zero
						sstdata[3].push({type: 2, data: write_shallow([ [],
							[ { type: 0, data: write_varint49(i) } ],
							[ { type: 0, data: write_varint49(1) } ],
							[ { type: 2, data: stru8(str) } ]
						])});
					});
				}
				sstroot.messages[0].data = write_shallow(sstdata);
			});

		}
		pb[4][0].data = write_shallow(store);
	}
	tmaroot.messages[0].data = write_shallow(pb);
}
