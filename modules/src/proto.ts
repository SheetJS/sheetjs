/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { u8concat } from "./util";

type Ptr = [number];
export { Ptr };

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
export { parse_varint49 };
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
export { write_varint49 };

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
export { varint_to_i32 };

interface ProtoItem {
	offset?: number;
	data: Uint8Array;
	type: number;
}
type ProtoField = Array<ProtoItem>
type ProtoMessage = Array<ProtoField>;
export { ProtoItem, ProtoField, ProtoMessage };
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
export { parse_shallow };
/** Serialize a shallow parse */
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
export { write_shallow };

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
export { mappa };
