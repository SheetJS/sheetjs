/*! otorp (C) 2013-present SheetJS -- http://sheetjs.com */

import { u8indexOf, u8str, u8_to_dataview } from "./util";
import { parse_macho } from "./macho";
import { Descriptor, parse_FileDescriptor, write_FileDescriptor } from './descriptor';


interface OtorpEntry {
	name: string;
	proto: string;
}
export { OtorpEntry };

/** Find and stringify all relevant protobuf defs */
function otorp(buf: Uint8Array, builtins = false): OtorpEntry[] {
	var res = proto_offsets(buf);
	var registry: {[key: string]: Descriptor} = {};
	var names: Set<string> = new Set();
	var out: OtorpEntry[] = [];

	res.forEach((r, i) => {
		if(!builtins && r[1].startsWith("google/protobuf/")) return;
		var b = buf.slice(r[0], i < res.length - 1 ? res[i+1][0] : buf.length);
		var pb = parse_FileDescriptorProto(b/*, r[1]*/);
		names.add(r[1]);
		registry[r[1]] = pb;
	});

	names.forEach(name => {
		/* ensure partial ordering by dependencies */
		names.delete(name);
		var pb = registry[name];
		var doit = (pb.dependency||[]).every((d: string) => !names.has(d));
		if(!doit) { names.add(name); return; }

		var dups = res.filter(r => r[1] == name);
		if(dups.length == 1) return out.push({ name, proto: write_FileDescriptor(pb) });

		/* in a fat binary, compare the defs for x86_64/aarch64 */
		var pbs = dups.map(r => {
			var i = res.indexOf(r);
			var b = buf.slice(r[0], i < res.length - 1 ? res[i+1][0] : buf.length);
			var pb = parse_FileDescriptorProto(b/*, r[1]*/);
			return write_FileDescriptor(pb);
		});
		for(var l = 1; l < pbs.length; ++l) if(pbs[l] != pbs[0]) throw new Error(`Conflicting definitions for ${name} at offsets 0x${dups[0][0].toString(16)} and 0x${dups[l][0].toString(16)}`);
		return out.push({ name, proto: pbs[0] });
	});

	return out;
}
export default otorp;

/** Determine if an address is being referenced */
var is_referenced = (buf: Uint8Array, pos: number): boolean => {
	var dv = u8_to_dataview(buf);

	/* Search for LEA reference (x86) */
	for(var leaddr = 0; leaddr > -1 && leaddr < pos; leaddr = u8indexOf(buf, 0x8D, leaddr + 1))
		if(dv.getUint32(leaddr + 2, true) == pos - leaddr - 6) return true;

	/* Search for absolute reference to address */
	try {
		var headers = parse_macho(buf);
		for(var i = 0; i < headers.length; ++i) {
			if(pos < headers[i].offset || pos > headers[i].offset + headers[i].size) continue;
			var b = headers[i].data;
			var p = pos - headers[i].offset;
			var ref = new Uint8Array([0,0,0,0,0,0,0,0]);
			var dv = u8_to_dataview(ref);
			dv.setInt32(0, p, true);
			if(u8indexOf(b, ref, 0) > 0) return true;
			ref[4] = 0x01;
			if(u8indexOf(b, ref, 0) > 0) return true;
		}
	} catch(e) {}
	return false;
};

type OffsetList = Array<[number, string, number, number]>;
/** Generate a list of potential starting points */
var proto_offsets = (buf: Uint8Array): OffsetList => {
	var meta = parse_macho(buf);
	var out: OffsetList = [];
	var off = 0;
	/* note: this loop only works for names < 128 chars */
	search: while((off = u8indexOf(buf, ".proto", off + 1)) > -1) {
		var pos = off;
		off += 6;
		while(off - pos < 256 && buf[pos] != off - pos - 1) {
			if(buf[pos] > 0x7F || buf[pos] < 0x20) continue search;
			--pos;
		}
		if(off - pos > 250) continue;
		var name = u8str(buf.slice(pos + 1, off));
		if(buf[--pos] != 0x0A) continue;
		if(!is_referenced(buf, pos)) { console.error(`Reference to ${name} not found`); continue; }
		var bin = meta.find(m => m.offset <= pos && m.offset + m.size >= pos);
		out.push([pos, name, bin?.type || -1, bin?.subtype || -1]);
	}
	return out;
};

/** Parse a descriptor that starts with the first byte of the supplied buffer */
var parse_FileDescriptorProto = (buf: Uint8Array): Descriptor => {
	var l = buf.length;
	while(l > 0) try {
		var b = buf.slice(0,l);
		var o = parse_FileDescriptor(b);
		return o;
	} catch(e) {
		var m = e.message.match(/at offset (\d+)/);
		if(m && parseInt(m[1], 10) < buf.length) l = parseInt(m[1], 10) - 1;
		else --l;
	}
	throw new RangeError("no protobuf message in range");
};
