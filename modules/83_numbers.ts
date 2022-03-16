/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
/// <reference path="src/types.ts"/>

import { CFB$Container } from 'cfb';
import { WorkBook, WorkSheet, Range, CellObject } from '../';
import type { utils } from "../";

declare var encode_cell: typeof utils.encode_cell;
declare var encode_range: typeof utils.encode_range;
declare var book_new: typeof utils.book_new;
declare var book_append_sheet: typeof utils.book_append_sheet;


var u8_to_dataview = (array: Uint8Array): DataView => new DataView(array.buffer, array.byteOffset, array.byteLength);

var u8str = (u8: Uint8Array): string => /* Buffer.isBuffer(u8) ? u8.toString() :*/ typeof TextDecoder != "undefined" ? new TextDecoder().decode(u8) : utf8read(a2s(u8));

var u8concat = (u8a: Uint8Array[]): Uint8Array => {
	var len = u8a.reduce((acc: number, x: Uint8Array) => acc + x.length, 0);
	var out = new Uint8Array(len);
	var off = 0;
	u8a.forEach(u8 => { out.set(u8, off); off += u8.length; });
	return out;
};

/* Hopefully one day this will be added to the language */
var popcnt = (x: number): number => {
  x -= ((x >> 1) & 0x55555555);
  x = (x & 0x33333333) + ((x >> 2) & 0x33333333);
  return (((x + (x >> 4)) & 0x0F0F0F0F) * 0x01010101) >>> 24;
};

/* Used in the modern cell storage */
var readDecimal128LE = (buf: Uint8Array, offset: number): number => {
	var exp = ((buf[offset + 15] & 0x7F) << 7) | (buf[offset + 14] >> 1);
	var mantissa = buf[offset + 14] & 1;
	for(var j = offset + 13; j >= offset; --j) mantissa = mantissa * 256 + buf[j];
	return ((buf[offset+15] & 0x80) ? -mantissa : mantissa) * Math.pow(10, exp - 0x1820);
};

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
/*
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
*/

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

interface ProtoItem {
	offset?: number;
	data: Uint8Array;
	type: number;
}
type ProtoField = Array<ProtoItem>
type ProtoMessage = Array<ProtoField>;
/** Shallow parse of a message */
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
		var v: ProtoItem = { offset: off, data: res, type };
		if(out[num] == null) out[num] = [v];
		else out[num].push(v);
	}
	return out;
}
/** Serialize a shallow parse */
/*
function write_shallow(proto: ProtoMessage): Uint8Array {
	var out: Uint8Array[] = [];
	proto.forEach((field, idx) => {
		field.forEach(item => {
			out.push(write_varint49(idx * 8 + item.type));
			out.push(item.data);
		});
	});
	return u8concat(out);
}
*/

function mappa<U>(data: ProtoField, cb:(Uint8Array) => U): U[] {
	if(!data) return [];
	return data.map((d) => { try {
		return cb(d.data);
	} catch(e) {
		var m = e.message?.match(/at offset (\d+)/);
		if(m) e.message = e.message.replace(/at offset (\d+)/, "at offset " + (+m[1] + d.offset));
		throw e;
	}});
}


interface IWAMessage {
	/** Metadata in .TSP.MessageInfo */
	meta: ProtoMessage;
	data: Uint8Array;
}
interface IWAArchiveInfo {
	id?: number;
	messages?: IWAMessage[];
}

function parse_iwa_file(buf: Uint8Array): IWAArchiveInfo[] {
	var out: IWAArchiveInfo[] = [], ptr: Ptr = [0];
	while(ptr[0] < buf.length) {
		/* .TSP.ArchiveInfo */
		var len = parse_varint49(buf, ptr);
		var ai = parse_shallow(buf.slice(ptr[0], ptr[0] + len));
		ptr[0] += len;

		var res: IWAArchiveInfo = {
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
		out.push(res);
	}
	return out;
}

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

function deframe(buf: Uint8Array): Uint8Array {
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

function parse_old_storage(buf: Uint8Array, sst: string[], rsst: string[]): CellObject {
  var dv = u8_to_dataview(buf);
  var ctype = buf[buf[0] == 4 ? 1 : 2];

  /* TODO: find the correct field position of number formats, formulae, etc */
  var flags = dv.getUint32(4, true);
  var data_offset = 12 + popcnt(flags & 0x0D8E) * 4;

  var ridx = -1, sidx = -1, ieee = NaN, dt = new Date(2001, 0, 1);
  if(flags & 0x0200) { ridx = dv.getUint32(data_offset,  true); data_offset += 4; }
  data_offset += popcnt(flags & 0x3000) * 4;
  if(flags & 0x0010) { sidx = dv.getUint32(data_offset,  true); data_offset += 4; }
  if(flags & 0x0020) { ieee = dv.getFloat64(data_offset, true); data_offset += 8; }
  if(flags & 0x0040) { dt.setTime(dt.getTime() +  dv.getFloat64(data_offset, true) * 1000); data_offset += 8; }

  var ret: CellObject;
  switch(ctype) {
    case 0: break; // return { t: "z" }; // blank?
    case 2: ret = { t: "n", v: ieee }; break; // number
    case 3: ret = { t: "s", v: sst[sidx] }; break; // string
    case 5: ret = { t: "d", v: dt }; break; // date-time
    case 6: ret = { t: "b", v: ieee > 0 }; break; // boolean
    case 7: ret = { t: "n", v: ieee }; break; // duration in seconds TODO: emit [hh]:[mm] style format with adjusted value
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

function parse_storage(buf: Uint8Array, sst: string[], rsst: string[]): CellObject {
  var dv = u8_to_dataview(buf);
  var ctype = buf[1];

  /* TODO: find the correct field position of number formats, formulae, etc */
  var flags = dv.getUint32(8, true);
  var data_offset = 12;

  var ridx = -1, sidx = -1, d128 = NaN, ieee = NaN, dt = new Date(2001, 0, 1);

  if(flags & 0x0001) { d128 = readDecimal128LE(buf, data_offset); data_offset += 16; }
  if(flags & 0x0002) { ieee = dv.getFloat64(data_offset, true); data_offset += 8; }
  if(flags & 0x0004) { dt.setTime(dt.getTime() +  dv.getFloat64(data_offset, true) * 1000); data_offset += 8; }
  if(flags & 0x0008) { sidx = dv.getUint32(data_offset,  true); data_offset += 4; }
  if(flags & 0x0010) { ridx = dv.getUint32(data_offset,  true); data_offset += 4; }

  var ret: CellObject;
  switch(ctype) {
    case 0: break; // return { t: "z" }; // blank?
    case 2: ret = { t: "n", v: d128 }; break; // number
    case 3: ret = { t: "s", v: sst[sidx] }; break; // string
    case 5: ret = { t: "d", v: dt }; break; // date-time
    case 6: ret = { t: "b", v: ieee > 0 }; break; // boolean
    case 7: ret = { t: "n", v: ieee }; break;  // duration in seconds TODO: emit [hh]:[mm] style format with adjusted value
    case 8: ret = { t: "e", v: 0}; break; // "formula error" TODO: enumerate and map errors to csf equivalents
    case 9: { // "automatic"?
      if(ridx > -1) ret = { t: "s", v: rsst[ridx] };
      else throw new Error(`Unsupported cell type ${ctype} : ${flags & 0x1F} : ${buf.slice(0,4)}`);
    } break;
    case 10: ret = { t: "n", v: d128 }; break; // currency
    default: throw new Error(`Unsupported cell type ${ctype} : ${flags & 0x1F} : ${buf.slice(0,4)}`);
  }
  /* TODO: All styling fields appear after the cell data */

  return ret;
}

function parse_cell_storage(buf: Uint8Array, sst: string[], rsst: string[]): CellObject {
  switch(buf[0]) {
    /* TODO: 0-2? */
    case 3: case 4: return parse_old_storage(buf, sst, rsst);
    case 5: return parse_storage(buf, sst, rsst);
    default: throw new Error(`Unsupported payload version ${buf[0]}`);
  }
}

// .TSP.Reference
function parse_TSP_Reference(buf: Uint8Array): number {
	var pb = parse_shallow(buf);
	return parse_varint49(pb[1][0].data);
}

// .TST.TableDataList
function parse_TST_TableDataList(M: IWAMessage[][], root: IWAMessage): string[] {
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

interface TileRowInfo {
	R: number;
	cells?: Uint8Array[];
}
// .TSP.TileRowInfo
function parse_TST_TileRowInfo(u8: Uint8Array): TileRowInfo {
	var pb = parse_shallow(u8);
	var R = varint_to_i32(pb[1][0].data) >>> 0;
	var pre_bnc = pb[3]?.[0]?.data;
	var pre_bnc_offsets = pb[4]?.[0]?.data && u8_to_dataview(pb[4][0].data);
	var storage = pb[6]?.[0]?.data;
	var storage_offsets = pb[7]?.[0]?.data && u8_to_dataview(pb[7][0].data);
	var wide_offsets = pb[8]?.[0]?.data && varint_to_i32(pb[8][0].data) > 0 || false;
	var width = wide_offsets ? 4 : 1;
	var cells = [];
	var off = 0;
	for(var C = 0; C < pre_bnc_offsets.byteLength/2; ++C) {
		/* prefer storage if it is present, otherwise fall back on pre_bnc */
		if(storage && storage_offsets) {
			off = storage_offsets.getUint16(C*2, true) * width;
			if(off < storage.length) { cells[C] = storage.subarray(off, storage_offsets.getUint16(C*2+2, true) * width); continue; }
		}
		if(pre_bnc && pre_bnc_offsets) {
			off = pre_bnc_offsets.getUint16(C*2, true) * width;
			if(off < pre_bnc.length) cells[C] = pre_bnc.subarray(off, pre_bnc_offsets.getUint16(C*2+2, true) * width);
		}
	}
	return { R, cells };
}

// .TST.Tile
function parse_TST_Tile(M: IWAMessage[][], root: IWAMessage): Uint8Array[][] {
	var pb = parse_shallow(root.data);
	var ri = mappa(pb[5], parse_TST_TileRowInfo);
	return ri.reduce((acc, x) => {
		if(!acc[x.R]) acc[x.R] = [];
		x.cells.forEach((cell, C) => {
			if(acc[x.R][C]) throw new Error(`Duplicate cell r=${x.R} c=${C}`);
			acc[x.R][C] = cell;
		});
		return acc;
	}, [] as Uint8Array[][]);
}

// .TST.TableModelArchive
function parse_TST_TableModelArchive(M: IWAMessage[][], root: IWAMessage, ws: WorkSheet) {
	var pb = parse_shallow(root.data);
	var range: Range = { s: {r:0, c:0}, e: {r:0, c:0} };
	range.e.r = (varint_to_i32(pb[6][0].data) >>> 0) - 1;
	if(range.e.r < 0) throw new Error(`Invalid row varint ${pb[6][0].data}`);
	range.e.c = (varint_to_i32(pb[7][0].data) >>> 0) - 1;
	if(range.e.c < 0) throw new Error(`Invalid col varint ${pb[7][0].data}`);
	ws["!ref"] = encode_range(range);

	{
		// .TST.DataStore
		var store = parse_shallow(pb[4][0].data);

		var sst = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[4][0].data)][0]);
		var rsst: string[] = store[17]?.[0] ? parse_TST_TableDataList(M, M[parse_TSP_Reference(store[17][0].data)][0]) : [];
		{
			// .TST.TileStorage
			var tile = parse_shallow(store[3][0].data);
			var tiles: Array<{id: number, ref: Uint8Array[][]}> = [];
			tile[1].forEach(t => {
				var tl = (parse_shallow(t.data));
				var ref = M[parse_TSP_Reference(tl[2][0].data)][0];
				var mtype = varint_to_i32(ref.meta[1][0].data);
				if(mtype != 6002) throw new Error(`6001 unexpected reference to ${mtype}`);
				tiles.push({id: varint_to_i32(tl[1][0].data), ref: parse_TST_Tile(M, ref) });
			});
			tiles.forEach((tile) => {
				tile.ref.forEach((row, R) => {
					row.forEach((buf, C) => {
						var addr = encode_cell({r:R,c:C});
						var res = parse_cell_storage(buf, sst, rsst);
						if(res) ws[addr] = res;
					});
				});
			});
		}
	}
}

// .TST.TableInfoArchive
function parse_TST_TableInfoArchive(M: IWAMessage[][], root: IWAMessage): WorkSheet {
	var pb = parse_shallow(root.data);
	var out: WorkSheet = { "!ref": "A1" };
	var tableref = M[parse_TSP_Reference(pb[2][0].data)];
	var mtype = varint_to_i32(tableref[0].meta[1][0].data);
	if(mtype != 6001) throw new Error(`6000 unexpected reference to ${mtype}`);
	parse_TST_TableModelArchive(M, tableref[0], out);
	return out;
}

// .TN.SheetArchive
interface NSheet {
	name: string;
	sheets: WorkSheet[];
}
function parse_TN_SheetArchive(M: IWAMessage[][], root: IWAMessage): NSheet {
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

// .TN.DocumentArchive
function parse_TN_DocumentArchive(M: IWAMessage[][], root: IWAMessage): WorkBook {
	var out = book_new();
	var pb = parse_shallow(root.data);
	var sheetoffs = mappa(pb[1], parse_TSP_Reference);
	sheetoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 2) {
				var root = parse_TN_SheetArchive(M, m);
				root.sheets.forEach(sheet => { book_append_sheet(out, sheet, root.name); });
			}
		});
	});
	if(out.SheetNames.length == 0) throw new Error("Empty NUMBERS file");
	return out;
}

/* exported parse_numbers_iwa */
function parse_numbers_iwa(cfb: CFB$Container): WorkBook {
	var out: IWAMessage[][] = [];
	cfb.FullPaths.forEach(p => { if(p.match(/\.iwpv2/)) throw new Error(`Unsupported password protection`); });
	/* collect entire message space */
	cfb.FileIndex.forEach(s => {
		if(!s.name.match(/\.iwa$/)) return;
		var o: Uint8Array;
		try { o = deframe(s.content as Uint8Array); } catch(e) { return console.log("?? " + s.content.length + " " + (e.message || e)); }
		var packets: IWAArchiveInfo[];
		try { packets = parse_iwa_file(o); } catch(e) { return console.log("## " + (e.message || e)); }
		packets.forEach(packet => {out[+packet.id] = packet.messages;});
	});
	if(!out.length) throw new Error("File has no messages");
	/* find document root */
	var docroot: IWAMessage;
	out.forEach((iwams) => {
		iwams.forEach((iwam) => {
			var mtype = varint_to_i32(iwam.meta[1][0].data) >>> 0;
			if(mtype == 1) {
				if(!docroot) docroot = iwam;
				else throw new Error("Document has multiple roots");
			}
		});
	});
	if(!docroot) throw new Error("Cannot find Document root");
	return parse_TN_DocumentArchive(out, docroot);
}
