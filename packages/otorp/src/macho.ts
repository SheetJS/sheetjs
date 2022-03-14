/*! otorp (C) 2013-present SheetJS -- http://sheetjs.com */

import { u8_to_dataview } from "./util";

interface MachOEntry {
	type: number;
	subtype: number;
	offset: number;
	size: number;
	align?: number;
	data: Uint8Array;
}
var parse_fat = (buf: Uint8Array): MachOEntry[] => {
	var dv = u8_to_dataview(buf);
	if(dv.getUint32(0, false) !== 0xCAFEBABE) throw new Error("Unsupported file");
	var nfat_arch = dv.getUint32(4, false);
	var out: MachOEntry[] = [];
	for(var i = 0; i < nfat_arch; ++i) {
		var start = i * 20 + 8;

		var cputype = dv.getUint32(start, false);
		var cpusubtype = dv.getUint32(start+4, false);
		var offset = dv.getUint32(start+8, false);
		var size = dv.getUint32(start+12, false);
		var align = dv.getUint32(start+16, false);

		out.push({
			type: cputype,
			subtype: cpusubtype,
			offset,
			size,
			align,
			data: buf.slice(offset, offset + size)
		});
	}
	return out;
};
var parse_macho = (buf: Uint8Array): MachOEntry[] => {
	var dv = u8_to_dataview(buf);
	var magic = dv.getUint32(0, false);
	switch(magic) {
		// fat binary (x86_64 / aarch64)
		case 0xCAFEBABE: return parse_fat(buf);
		// x86_64
		case 0xCFFAEDFE: return [{
			type: dv.getUint32(4, false),
			subtype: dv.getUint32(8, false),
			offset: 0,
			size: buf.length,
			data: buf
		}];
	}
	throw new Error("Unsupported file");
};
export { MachOEntry, parse_macho };