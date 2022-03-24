/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
/// <reference path="src/types.ts"/>

/* these are type imports and do not show up in the generated JS */
import { CFB$Container, CFB$Entry } from 'cfb';
import { WorkBook, WorkSheet, Range, CellObject } from '../';
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

function u8_to_dataview(array: Uint8Array): DataView { return new DataView(array.buffer, array.byteOffset, array.byteLength); }
//<<export { u8_to_dataview };

function u8str(u8: Uint8Array): string { return /* Buffer.isBuffer(u8) ? u8.toString() :*/ typeof TextDecoder != "undefined" ? new TextDecoder().decode(u8) : utf8read(a2s(u8)); }
function stru8(str: string): Uint8Array { return typeof TextEncoder != "undefined" ? new TextEncoder().encode(str) : s2a(utf8write(str)) as Uint8Array; }
//<<export { u8str, stru8 };

function u8contains(body: Uint8Array, search: Uint8Array): boolean {
  outer: for(var L = 0; L <= body.length - search.length; ++L) {
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
	var exp = Math.floor(value == 0 ? 0 : /*Math.log10*/Math.LOG10E * Math.log(Math.abs(value))) + 0x1820 - 20;
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
	return usz.slice(0, L);
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
				res = buf.slice(l, ptr[0]);
			} break;
			case 5: len = 4; res = buf.slice(ptr[0], ptr[0] + len); ptr[0] += len; break;
			case 1: len = 8; res = buf.slice(ptr[0], ptr[0] + len); ptr[0] += len; break;
			case 2: len = parse_varint49(buf, ptr); res = buf.slice(ptr[0], ptr[0] + len); ptr[0] += len; break;
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
	id?: number;
	merge?: boolean;
	messages?: IWAMessage[];
}
/** Extract all messages from a IWA file */
function parse_iwa_file(buf: Uint8Array): IWAArchiveInfo[] {
	var out: IWAArchiveInfo[] = [], ptr: Ptr = [0];
	while(ptr[0] < buf.length) {
		/* .TSP.ArchiveInfo */
		var len = parse_varint49(buf, ptr);
		var ai = parse_shallow(buf.slice(ptr[0], ptr[0] + len));
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
				data: buf.slice(ptr[0], ptr[0] + fl)
			});
			ptr[0] += fl;
		});
		if(ai[3]?.[0]) res.merge = (varint_to_i32(ai[3][0].data) >>> 0) > 0;
		out.push(res);
	}
	return out;
}
function write_iwa_file(ias: IWAArchiveInfo[]): Uint8Array {
	var bufs: Uint8Array[] = [];
	ias.forEach(ia => {
		var ai: ProtoMessage = [];
		ai[1] = [{data: write_varint49(ia.id), type: 0}];
		ai[2] = [];
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
function parse_snappy_chunk(type: number, buf: Uint8Array): Uint8Array {
	if(type != 0) throw new Error(`Unexpected Snappy chunk type ${type}`);
	var ptr: Ptr = [0];

	var usz = parse_varint49(buf, ptr);
	var chunks = [];
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
			chunks.push(buf.slice(ptr[0], ptr[0] + len)); ptr[0] += len; continue;
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
			chunks = [u8concat(chunks)];
			if(offset == 0) throw new Error("Invalid offset 0");
			if(offset > chunks[0].length) throw new Error("Invalid offset beyond length");
			if(length >= offset) {
				chunks.push(chunks[0].slice(-offset)); length -= offset;
				while(length >= chunks[chunks.length-1].length) {
					chunks.push(chunks[chunks.length - 1]);
					length -= chunks[chunks.length - 1].length;
				}
			}
			chunks.push(chunks[0].slice(-offset, -offset + length));
		}
	}
	var o = u8concat(chunks);
	if(o.length != usz) throw new Error(`Unexpected length: ${o.length} != ${usz}`);
	return o;
}

/** Decompress IWA file */
function decompress_iwa_file(buf: Uint8Array): Uint8Array {
	var out = [];
	var l = 0;
	while(l < buf.length) {
		var t = buf[l++];
		var len = buf[l] | (buf[l+1]<<8) | (buf[l+2] << 16); l += 3;
		out.push(parse_snappy_chunk(t, buf.slice(l, l + len)));
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

		out.push(buf.slice(l, l + c)); L += c;

		frame[0] = 0;
		frame[1] = L & 0xFF; frame[2] = (L >>  8) & 0xFF; frame[3] = (L >> 16) & 0xFF;
		l += c;
	}
	return u8concat(out);
}
//<<export { decompress_iwa_file, compress_iwa_file };

/** Parse "old storage" (version 0..3) */
function parse_old_storage(buf: Uint8Array, sst: string[], rsst: string[], v: 0|1|2|3): CellObject {
  var dv = u8_to_dataview(buf);
	var flags = dv.getUint32(4, true);

  /* TODO: find the correct field position of number formats, formulae, etc */
  var data_offset = (v > 1 ? 12 : 8) + popcnt(flags & (v > 1 ? 0x0D8E : 0x018E)) * 4;

  var ridx = -1, sidx = -1, ieee = NaN, dt = new Date(2001, 0, 1);
  if(flags & 0x0200) { ridx = dv.getUint32(data_offset,  true); data_offset += 4; }
  data_offset += popcnt(flags & (v > 1 ? 0x3000 : 0x1000)) * 4;
  if(flags & 0x0010) { sidx = dv.getUint32(data_offset,  true); data_offset += 4; }
  if(flags & 0x0020) { ieee = dv.getFloat64(data_offset, true); data_offset += 8; }
  if(flags & 0x0040) { dt.setTime(dt.getTime() +  dv.getFloat64(data_offset, true) * 1000); data_offset += 8; }

  var ret: CellObject;
  switch(buf[2]) {
    case 0: break; // return { t: "z" }; // blank?
    case 2: ret = { t: "n", v: ieee }; break; // number
    case 3: ret = { t: "s", v: sst[sidx] }; break; // string
    case 5: ret = { t: "d", v: dt }; break; // date-time
    case 6: ret = { t: "b", v: ieee > 0 }; break; // boolean
    case 7: ret = { t: "n", v: ieee / 86400 }; break; // duration in seconds TODO: emit [hh]:[mm] style format with adjusted value
    case 8: ret = { t: "e", v: 0}; break; // "formula error" TODO: enumerate and map errors to csf equivalents
    case 9: { // "automatic"?
      if(ridx > -1) ret = { t: "s", v: rsst[ridx] };
      else if(sidx > -1) ret = { t: "s", v: sst[sidx] };
      else if(!isNaN(ieee)) ret = { t: "n", v: ieee };
      else throw new Error(`Unsupported cell type ${buf.slice(0,4)}`);
    } break;
    default: throw new Error(`Unsupported cell type ${buf.slice(0,4)}`);
  }
  /* TODO: Some fields appear after the cell data */

  return ret;
}

/** Parse "new storage" (version 5) */
function parse_new_storage(buf: Uint8Array, sst: string[], rsst: string[]): CellObject {
  var dv = u8_to_dataview(buf);
  var flags = dv.getUint32(8, true);

  /* TODO: find the correct field position of number formats, formulae, etc */
  var data_offset = 12;

  var ridx = -1, sidx = -1, d128 = NaN, ieee = NaN, dt = new Date(2001, 0, 1);

  if(flags & 0x0001) { d128 = readDecimal128LE(buf, data_offset); data_offset += 16; }
  if(flags & 0x0002) { ieee = dv.getFloat64(data_offset, true); data_offset += 8; }
  if(flags & 0x0004) { dt.setTime(dt.getTime() +  dv.getFloat64(data_offset, true) * 1000); data_offset += 8; }
  if(flags & 0x0008) { sidx = dv.getUint32(data_offset,  true); data_offset += 4; }
  if(flags & 0x0010) { ridx = dv.getUint32(data_offset,  true); data_offset += 4; }

  var ret: CellObject;
  switch(buf[1]) {
    case 0: break; // return { t: "z" }; // blank?
    case 2: ret = { t: "n", v: d128 }; break; // number
    case 3: ret = { t: "s", v: sst[sidx] }; break; // string
    case 5: ret = { t: "d", v: dt }; break; // date-time
    case 6: ret = { t: "b", v: ieee > 0 }; break; // boolean
    case 7: ret = { t: "n", v: ieee / 86400 }; break;  // duration in seconds TODO: emit [hh]:[mm] style format with adjusted value
    case 8: ret = { t: "e", v: 0}; break; // "formula error" TODO: enumerate and map errors to csf equivalents
    case 9: { // "automatic"?
      if(ridx > -1) ret = { t: "s", v: rsst[ridx] };
      else throw new Error(`Unsupported cell type ${buf[1]} : ${flags & 0x1F} : ${buf.slice(0,4)}`);
    } break;
    case 10: ret = { t: "n", v: d128 }; break; // currency
    default: throw new Error(`Unsupported cell type ${buf[1]} : ${flags & 0x1F} : ${buf.slice(0,4)}`);
  }
  /* TODO: All styling fields appear after the cell data */

  return ret;
}
function write_new_storage(cell: CellObject, sst: string[]): Uint8Array {
	var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
	out[0] = 5;
	switch(cell.t) {
		case "n": out[1] = 2; writeDecimal128LE(out, l, cell.v as number); flags |= 1; l += 16; break;
		case "b": out[1] = 6; dv.setFloat64(l, cell.v ? 1 : 0, true); flags |= 2; l += 8; break;
		case "s":
			if(sst.indexOf(cell.v as string) == -1) throw new Error(`Value ${cell.v} missing from SST!`);
			out[1] = 3; dv.setUint32(l, sst.indexOf(cell.v as string), true); flags |= 8; l += 4; break;
		default: throw "unsupported cell type " + cell.t;
	}
	dv.setUint32(8, flags, true);
	return out.slice(0, l);
}
function write_old_storage(cell: CellObject, sst: string[]): Uint8Array {
	var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
	out[0] = 3;
	switch(cell.t) {
		case "n": out[2] = 2; dv.setFloat64(l, cell.v as number, true); flags |= 0x20; l += 8; break;
		case "b": out[2] = 6; dv.setFloat64(l, cell.v ? 1 : 0, true); flags |= 0x20; l += 8; break;
		case "s":
			if(sst.indexOf(cell.v as string) == -1) throw new Error(`Value ${cell.v} missing from SST!`);
			out[2] = 3; dv.setUint32(l, sst.indexOf(cell.v as string), true); flags |= 0x10; l += 4; break;
		default: throw "unsupported cell type " + cell.t;
	}
	dv.setUint32(4, flags, true);
	return out.slice(0, l);
}
//<<export { write_new_storage, write_old_storage };
function parse_cell_storage(buf: Uint8Array, sst: string[], rsst: string[]): CellObject {
  switch(buf[0]) {
    case 0: case 1:
    case 2: case 3: return parse_old_storage(buf, sst, rsst, buf[0]);
    case 5: return parse_new_storage(buf, sst, rsst);
    default: throw new Error(`Unsupported payload version ${buf[0]}`);
  }
}

/** .TSS.StylesheetArchive */
//function parse_TSS_StylesheetArchive(M: IWAMessage[][], root: IWAMessage): void {
//	var pb = parse_shallow(root.data);
//}

/** .TSP.Reference */
function parse_TSP_Reference(buf: Uint8Array): number {
	var pb = parse_shallow(buf);
	return parse_varint49(pb[1][0].data);
}
function write_TSP_Reference(idx: number): Uint8Array {
	var out: ProtoMessage = [];
	out[1] = [ { type: 0, data: write_varint49(idx) } ];
	return write_shallow(out)
}
//<<export { parse_TSP_Reference, write_TSP_Reference };

type MessageSpace = {[id: number]: IWAMessage[]};

/** .TST.TableDataList */
function parse_TST_TableDataList(M: MessageSpace, root: IWAMessage): string[] {
	var pb = parse_shallow(root.data);
	// .TST.TableDataList.ListType
	var type = varint_to_i32(pb[1][0].data);

	var entries = pb[3];
	var data = [];
	(entries||[]).forEach(entry => {
		// .TST.TableDataList.ListEntry
		var le = parse_shallow(entry.data);
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
		}
	});
	return data;
}

type TileStorageType = -1 | 0 | 1;
interface TileRowInfo {
	/** Row Index */
	R: number;
	/** Cell Storage */
	cells?: Uint8Array[];
}
/** .TSP.TileRowInfo */
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
	for(C = 0; C < offsets.length - 1; ++C) cells[offsets[C][0]] = used_storage.subarray(offsets[C][1] * width, offsets[C+1][1] * width);
	if(offsets.length >= 1) cells[offsets[offsets.length - 1][0]] = used_storage.subarray(offsets[offsets.length - 1][1] * width);
	return { R, cells };
}

interface TileInfo {
	data: Uint8Array[][];
	nrows: number;
}
/** .TST.Tile */
function parse_TST_Tile(M: MessageSpace, root: IWAMessage): TileInfo {
	var pb = parse_shallow(root.data);
	var storage: TileStorageType = (pb?.[7]?.[0]) ? ((varint_to_i32(pb[7][0].data)>>>0) > 0 ? 1 : 0 ) : -1;
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

/** .TST.TableModelArchive (6001) */
function parse_TST_TableModelArchive(M: MessageSpace, root: IWAMessage, ws: WorkSheet) {
	var pb = parse_shallow(root.data);
	var range: Range = { s: {r:0, c:0}, e: {r:0, c:0} };
	range.e.r = (varint_to_i32(pb[6][0].data) >>> 0) - 1;
	if(range.e.r < 0) throw new Error(`Invalid row varint ${pb[6][0].data}`);
	range.e.c = (varint_to_i32(pb[7][0].data) >>> 0) - 1;
	if(range.e.c < 0) throw new Error(`Invalid col varint ${pb[7][0].data}`);
	ws["!ref"] = encode_range(range);

	// .TST.DataStore
	var store = parse_shallow(pb[4][0].data);
	var sst = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[4][0].data)][0]);
	var rsst: string[] = store[17]?.[0] ? parse_TST_TableDataList(M, M[parse_TSP_Reference(store[17][0].data)][0]) : [];

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
				var addr = encode_cell({r:_R + R,c:C});
				var res = parse_cell_storage(buf, sst, rsst);
				if(res) ws[addr] = res;
			});
		});
		_R += _tile.nrows;
	});
}

/** .TST.TableInfoArchive (6000) */
function parse_TST_TableInfoArchive(M: MessageSpace, root: IWAMessage): WorkSheet {
	var pb = parse_shallow(root.data);
	var out: WorkSheet = { "!ref": "A1" };
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
/** .TN.SheetArchive (2) */
function parse_TN_SheetArchive(M: MessageSpace, root: IWAMessage): NSheet {
	var pb = parse_shallow(root.data);
	var out: NSheet = {
		name: (pb[1]?.[0] ? u8str(pb[1][0].data) : ""),
		sheets: []
	};
	var shapeoffs = mappa(pb[2], parse_TSP_Reference);
	shapeoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 6000) out.sheets.push(parse_TST_TableInfoArchive(M, m));
		});
	});
	return out;
}

/** .TN.DocumentArchive */
function parse_TN_DocumentArchive(M: MessageSpace, root: IWAMessage): WorkBook {
	var out = book_new();
	var pb = parse_shallow(root.data);

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
				var root = parse_TN_SheetArchive(M, m);
				root.sheets.forEach((sheet, idx) => { book_append_sheet(out, sheet, idx == 0 ? root.name : root.name + "_" + idx, true); });
			}
		});
	});
	if(out.SheetNames.length == 0) throw new Error("Empty NUMBERS file");
	return out;
}

/** Parse NUMBERS file */
function parse_numbers_iwa(cfb: CFB$Container): WorkBook {
	var M: MessageSpace = {}, indices: number[] = [];
	cfb.FullPaths.forEach(p => { if(p.match(/\.iwpv2/)) throw new Error(`Unsupported password protection`); });

	/* collect entire message space */
	cfb.FileIndex.forEach(s => {
		if(!s.name.match(/\.iwa$/)) return;
		var o: Uint8Array;
		try { o = decompress_iwa_file(s.content as Uint8Array); } catch(e) { return console.log("?? " + s.content.length + " " + (e.message || e)); }
		var packets: IWAArchiveInfo[];
		try { packets = parse_iwa_file(o); } catch(e) { return console.log("## " + (e.message || e)); }
		packets.forEach(packet => { M[packet.id] = packet.messages; indices.push(packet.id); });
	});
	if(!indices.length) throw new Error("File has no messages");

	/* find document root */
	var docroot: IWAMessage = M?.[1]?.[0]?.meta?.[1]?.[0].data && varint_to_i32(M[1][0].meta[1][0].data) == 1 && M[1][0];
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
	return parse_TN_DocumentArchive(M, docroot);
}

//<<export { parse_numbers_iwa };

interface DependentInfo {
  deps: number[];
  location: string;
  type: number;
}

function write_tile_row(tri: ProtoMessage, data: any[], SST: string[]): number {
  if(!tri[6]?.[0] || !tri[7]?.[0]) throw "Mutation only works on post-BNC storages!";
	var wide_offsets = tri[8]?.[0]?.data && varint_to_i32(tri[8][0].data) > 0 || false;
  if(wide_offsets) throw "Math only works with normal offsets";
	var cnt = 0;
  var dv = u8_to_dataview(tri[7][0].data), last_offset = 0, cell_storage: Uint8Array[] = [];
  var _dv = u8_to_dataview(tri[4][0].data), _last_offset = 0, _cell_storage: Uint8Array[] = [];
	for(var C = 0; C < data.length; ++C) {
		if(data[C] == null) { dv.setUint16(C*2, 0xFFFF, true); _dv.setUint16(C*2, 0xFFFF); continue; }
		dv.setUint16(C*2, last_offset, true);
		_dv.setUint16(C*2, _last_offset, true);
		var celload: Uint8Array, _celload: Uint8Array;
		switch(typeof data[C]) {
			case "string":
				celload = write_new_storage({t: "s", v: data[C]}, SST);
				_celload = write_old_storage({t: "s", v: data[C]}, SST);
				break;
			case "number":
				celload = write_new_storage({t: "n", v: data[C]}, SST);
				_celload = write_old_storage({t: "n", v: data[C]}, SST);
				break;
			case "boolean":
				celload = write_new_storage({t: "b", v: data[C]}, SST);
				_celload = write_old_storage({t: "b", v: data[C]}, SST);
				break;
			default: throw new Error("Unsupported value " + data[C]);
		}
		cell_storage.push(celload); last_offset += celload.length;
		_cell_storage.push(_celload); _last_offset += _celload.length;
		++cnt;
	}
	tri[2][0].data = write_varint49(cnt);

	for(; C < tri[7][0].data.length/2; ++C) {
		dv.setUint16(C*2, 0xFFFF, true);
		_dv.setUint16(C*2, 0xFFFF, true);
	}
	tri[6][0].data = u8concat(cell_storage);
	tri[3][0].data = u8concat(_cell_storage);
	return cnt;
}

function write_numbers_iwa(wb: WorkBook, opts: any): CFB$Container {
	if(!opts || !opts.numbers) throw new Error("Must pass a `numbers` option -- check the README");
	/* TODO: support multiple worksheets, larger ranges, more data types, etc */
	var ws = wb.Sheets[wb.SheetNames[0]];
	if(wb.SheetNames.length > 1) console.error("The Numbers writer currently writes only the first table");
	var range = decode_range(ws["!ref"]);
	range.s.r = range.s.c = 0;
	var trunc = false;
	if(range.e.c > 9) { trunc = true; range.e.c = 9; }
	if(range.e.r > 49) { trunc = true; range.e.r = 49; }
	if(trunc) console.error(`The Numbers writer is currently limited to ${encode_range(range)}`);
	var data = sheet_to_json<any>(ws, { range, header: 1 });
	var SST = ["~Sh33tJ5~"];
	data.forEach(row => row.forEach(cell => { if(typeof cell == "string") SST.push(cell); }))

	var dependents: {[x:number]: DependentInfo} = {};
	var indices: number[] = [];

	var cfb: CFB$Container = CFB.read(opts.numbers, { type: "base64" });
	cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
		var fi = row[0], fp = row[1];
		if(fi.type != 2) return;
		if(!fi.name.match(/\.iwa/)) return;

		/* Reframe .iwa files */
		var old_content = fi.content;
		var raw1 = decompress_iwa_file(old_content as Uint8Array);

		var x = parse_iwa_file(raw1);

		x.forEach(packet => {
			indices.push(packet.id);
			dependents[packet.id] = { deps: [], location: fp, type: varint_to_i32(packet.messages[0].meta[1][0].data) };
		});
	});

	indices.sort((x,y) => x-y);
	var indices_varint: Array<[number, Uint8Array]> = indices.filter(x => x > 1).map(x => [x, write_varint49(x)] );

	cfb.FileIndex.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
		var fi = row[0], fp = row[1];
		if(!fi.name.match(/\.iwa/)) return;
		var x = parse_iwa_file(decompress_iwa_file(fi.content as Uint8Array));

		x.forEach(ia => {
			ia.messages.forEach(m => {
				indices_varint.forEach(ivi => {
					if(ia.messages.some(mess => varint_to_i32(mess.meta[1][0].data) != 11006 && u8contains(mess.data, ivi[1]))) {
						dependents[ivi[0]].deps.push(ia.id);
					}
				})
			});
		});
	});

	function get_unique_msgid() {
		for(var i = 927262; i < 2000000; ++i) if(!dependents[i]) return i;
		throw new Error("Too many messages");
	}

	/* .TN.DocumentArchive */
	var entry = CFB.find(cfb, dependents[1].location);
  var x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  var docroot: IWAArchiveInfo;
  for(var xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if(packet.id == 1) docroot = packet;
  }

	/* .TN.SheetArchive */
  var sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[1][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }

	/* .TST.TableInfoArchive */
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }

	/* .TST.TableModelArchive */
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content as Uint8Array));
  for(xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if(packet.id == sheetrootref) docroot = packet;
  }

  var pb = parse_shallow(docroot.messages[0].data);
  {
		pb[6][0].data = write_varint49(range.e.r + 1);
		pb[7][0].data = write_varint49(range.e.c + 1);
		// pb[22] = [ { type: 0, data: write_varint49(1) } ]; // enables table display

		var cruidsref = parse_TSP_Reference(pb[46][0].data);
		var oldbucket = CFB.find(cfb, dependents[cruidsref].location);
		var _x = parse_iwa_file(decompress_iwa_file(oldbucket.content as Uint8Array));
		{
			for(var j = 0; j < _x.length; ++j) {
				if(_x[j].id == cruidsref) break;
			}
			if(_x[j].id != cruidsref) throw "Bad ColumnRowUIDMapArchive";
			var cruids = parse_shallow(_x[j].messages[0].data);
			cruids[1] = []; cruids[2] = [], cruids[3] = [];
			for(var C = 0; C <= range.e.c; ++C) {
				var uuid: ProtoMessage = [];
				uuid[1] = uuid[2] = [ { type: 0, data: write_varint49(C + 420690) } ]
				cruids[1].push({type: 2, data: write_shallow(uuid)})
				cruids[2].push({type: 0, data: write_varint49(C)});
				cruids[3].push({type: 0, data: write_varint49(C)});
			}
			cruids[4] = []; cruids[5] = [], cruids[6] = [];
			for(var R = 0; R <= range.e.r; ++R) {
				uuid = [];
				uuid[1] = uuid[2] = [ { type: 0, data: write_varint49(R + 726270)} ];
				cruids[4].push({type: 2, data: write_shallow(uuid)})
				cruids[5].push({type: 0, data: write_varint49(R)});
				cruids[6].push({type: 0, data: write_varint49(R)});
			}
			_x[j].messages[0].data = write_shallow(cruids);
		}
		oldbucket.content = compress_iwa_file(write_iwa_file(_x)); oldbucket.size = oldbucket.content.length;
		delete pb[46];

    var store = parse_shallow(pb[4][0].data);
    {
			store[7][0].data = write_varint49(range.e.r + 1);
			var row_headers = parse_shallow(store[1][0].data);
			var row_header_ref = parse_TSP_Reference(row_headers[2][0].data);
			oldbucket = CFB.find(cfb, dependents[row_header_ref].location);
			_x = parse_iwa_file(decompress_iwa_file(oldbucket.content as Uint8Array));
			{
				if(_x[0].id != row_header_ref) throw "Bad HeaderStorageBucket";
				var base_bucket = parse_shallow(_x[0].messages[0].data);
				for(R = 0; R < data.length; ++R) {
					var _bucket = parse_shallow(base_bucket[2][0].data);
					_bucket[1][0].data = write_varint49(R);
					_bucket[4][0].data = write_varint49(data[R].length);
					base_bucket[2][R] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
				}
				_x[0].messages[0].data = write_shallow(base_bucket);
			}
			oldbucket.content = compress_iwa_file(write_iwa_file(_x)); oldbucket.size = oldbucket.content.length;

			var col_header_ref = parse_TSP_Reference(store[2][0].data);
			oldbucket = CFB.find(cfb, dependents[col_header_ref].location);
			_x = parse_iwa_file(decompress_iwa_file(oldbucket.content as Uint8Array));
			{
				if(_x[0].id != col_header_ref) throw "Bad HeaderStorageBucket";
				base_bucket = parse_shallow(_x[0].messages[0].data);
				for(C = 0; C <= range.e.c; ++C) {
					_bucket = parse_shallow(base_bucket[2][0].data);
					_bucket[1][0].data = write_varint49(C);
					_bucket[4][0].data = write_varint49(range.e.r + 1);
					base_bucket[2][C] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
				}
				_x[0].messages[0].data = write_shallow(base_bucket);
			}
			oldbucket.content = compress_iwa_file(write_iwa_file(_x)); oldbucket.size = oldbucket.content.length;


      /* ref to string table */
      var sstref = parse_TSP_Reference(store[4][0].data);
      (() => {
        var sentry = CFB.find(cfb, dependents[sstref].location);
        var sx = parse_iwa_file(decompress_iwa_file(sentry.content as Uint8Array));
        var sstroot: IWAArchiveInfo;
        for(var sxi = 0; sxi < sx.length; ++sxi) {
          var packet = sx[sxi];
          if(packet.id == sstref) sstroot = packet;
        }

        var sstdata = parse_shallow(sstroot.messages[0].data);
        {
          sstdata[3] = [];
          var newsst: ProtoMessage = [];
					SST.forEach((str, i) => {
						newsst[1] = [ { type: 0, data: write_varint49(i) } ];
						newsst[2] = [ { type: 0, data: write_varint49(1) } ];
						newsst[3] = [ { type: 2, data: stru8(str) } ];
						sstdata[3].push({type: 2, data: write_shallow(newsst)});
					});
        }
        sstroot.messages[0].data = write_shallow(sstdata);

        var sy = write_iwa_file(sx);
        var raw3 = compress_iwa_file(sy);
        sentry.content = raw3; sentry.size = sentry.content.length;
      })();

      var tile = parse_shallow(store[3][0].data); // TileStorage
      {
        var t = tile[1][0];
        delete tile[2];
        var tl = parse_shallow(t.data); // first Tile
        {
          var tileref = parse_TSP_Reference(tl[2][0].data);
          (() => {
            var tentry = CFB.find(cfb, dependents[tileref].location);
            var tx = parse_iwa_file(decompress_iwa_file(tentry.content as Uint8Array));
            var tileroot: IWAArchiveInfo;
            for(var sxi = 0; sxi < tx.length; ++sxi) {
              var packet = tx[sxi];
              if(packet.id == tileref) tileroot = packet;
            }

            // .TST.Tile
            var tiledata = parse_shallow(tileroot.messages[0].data);
            {
							delete tiledata[6]; delete tile[7];
							var rowload = new Uint8Array(tiledata[5][0].data);
							tiledata[5] = [];
							var cnt = 0;
							for(var R = 0; R <= range.e.r; ++R) {
								var tilerow = parse_shallow(rowload);
								cnt += write_tile_row(tilerow, data[R], SST);
								tilerow[1][0].data = write_varint49(R);
								tiledata[5].push({data: write_shallow(tilerow), type: 2});
							}
							tiledata[1] = [{type: 0, data: write_varint49(range.e.c + 1)}];
							tiledata[2] = [{type: 0, data: write_varint49(range.e.r + 1)}];
							tiledata[3] = [{type: 0, data: write_varint49(cnt)}];
							tiledata[4] = [{type: 0, data: write_varint49(range.e.r + 1)}];
            }
            tileroot.messages[0].data = write_shallow(tiledata);

            var ty = write_iwa_file(tx);
            var raw3 = compress_iwa_file(ty);
            tentry.content = raw3; tentry.size = tentry.content.length;
            //throw dependents[tileref];
          })();
        }
        t.data = write_shallow(tl);
      }
      store[3][0].data = write_shallow(tile);
    }
    pb[4][0].data = write_shallow(store);
  }
  docroot.messages[0].data = write_shallow(pb);

  var y = write_iwa_file(x);
  var raw3 = compress_iwa_file(y);
  entry.content = raw3; entry.size = entry.content.length;

	return cfb;
}
