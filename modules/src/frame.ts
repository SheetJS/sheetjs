/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { Ptr, parse_varint49 } from './proto';
import { u8concat } from './util';

function is_framed(buf: Uint8Array): boolean {
	var l = 0;
	while(l < buf.length) {
		l++;
		var len = buf[l] | (buf[l+1]<<8) | (buf[l+2] << 16); l += 3;
		l += len;
	}
	return l == buf.length;
}
export { is_framed };

function deframe(buf: Uint8Array): Uint8Array {
	var out = [];
	var l = 0;
	while(l < buf.length) {
		var t = buf[l++];
		var len = buf[l] | (buf[l+1]<<8) | (buf[l+2] << 16); l += 3;
		out.push(process_chunk(t, buf.slice(l, l + len)));
		l += len;
	}
	if(l !== buf.length) throw new Error("data is not a valid framed stream!");
	return u8concat(out);
}
export { deframe };

function process_chunk(type: number, buf: Uint8Array): Uint8Array {
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