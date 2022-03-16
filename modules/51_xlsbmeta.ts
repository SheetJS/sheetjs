/// <reference path="src/types.ts"/>

/* [MS-XLSB] 2.4.698 BrtMdtinfo */
interface BrtMdtinfo {
	flags: number;
	version: number;
	name: string;
}
function parse_BrtMdtinfo(data: ReadableData, length: number): BrtMdtinfo {
	return {
		flags: data.read_shift(4),
		version: data.read_shift(4),
		name: parse_XLWideString(data, length - 8)
	};
}
function write_BrtMdtinfo(data: BrtMdtinfo): RawData {
	var o = new_buf(12 + 2 * data.name.length);
	o.write_shift(4, data.flags);
	o.write_shift(4, data.version);
	write_XLWideString(data.name, o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.697 BrtMdb */
type Mdir = [number, number]; // "t", "v" in XLSX parlance
type BrtMdb = Mdir[];
function parse_BrtMdb(data: ReadableData/*, length: number*/): BrtMdb {
	var out: Mdir[] = [];
	var cnt = data.read_shift(4);
	while(cnt-- > 0) out.push([data.read_shift(4), data.read_shift(4)]);
	return out;
}
function write_BrtMdb(mdb: BrtMdb): RawData {
	var o = new_buf(4 + 8 * mdb.length);
	o.write_shift(4, mdb.length);
	for(var i = 0; i < mdb.length; ++i) {
		o.write_shift(4, mdb[i][0]);
		o.write_shift(4, mdb[i][1]);
	}
	return o;
}

/* [MS-XLSB] 2.4.72 BrtBeginEsfmd */
function write_BrtBeginEsfmd(cnt: number, name: string): RawData {
	var o = new_buf(8 + 2 * name.length);
	o.write_shift(4, cnt);
	write_XLWideString(name, o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.73 BrtBeginEsmdb */
function parse_BrtBeginEsmdb(data: ReadableData/*, length: number*/): boolean {
	data.l += 4;
	return data.read_shift(4) != 0;
}
function write_BrtBeginEsmdb(cnt: number, cm: boolean): RawData {
	var o = new_buf(8);
	o.write_shift(4, cnt);
	o.write_shift(4, cm ? 1 : 0);
	return o;
}

/* [MS-XLSB] 2.1.7.34 Metadata */
function parse_xlmeta_bin(data, name: string, _opts?: ParseXLMetaOptions): XLMeta {
	var out: XLMeta = { Types: [], Cell: [], Value: [] };
	var opts = _opts || {};
	var state: number[] = [];
	var pass = false;
	var metatype: 0 | 1 | 2 = 2;
	recordhopper(data, (val, R, RT) => {
		switch(RT) {
			// case 0x014C: /* BrtBeginMetadata */
			// case 0x014D: /* BrtEndMetadata */
			// case 0x014E: /* BrtBeginEsmdtinfo */
			// case 0x0150: /* BrtEndEsmdtinfo */
			// case 0x0153: /* BrtBeginEsfmd */
			// case 0x0154: /* BrtEndEsfmd */
			// case 0x0034: /* BrtBeginFmd */
			// case 0x0035: /* BrtEndFmd */
			// case 0x1000: /* BrtBeginDynamicArrayPr */
			// case 0x1001: /* BrtEndDynamicArrayPr */
			// case 0x138A: /* BrtBeginRichValueBlock */
			// case 0x138B: /* BrtEndRichValueBlock */

			case 0x014F: /* BrtMdtinfo */
				out.Types.push({name: (val as BrtMdtinfo).name}); break;

			case 0x0033: /* BrtMdb */
				(val as BrtMdb).forEach(r => {
					if(metatype == 1) out.Cell.push({type: out.Types[r[0] - 1].name, index: r[1] });
					else if(metatype == 0) out.Value.push({type: out.Types[r[0] - 1].name, index: r[1] });
				}); break;

			case 0x0151: /* BrtBeginEsmdb */
				metatype = (val as boolean) ? 1 /* cell */ : 0 /* value */; break;
			case 0x0152: /* BrtEndEsmdb */
				metatype = 2; break;

			case 0x0023: /* BrtFRTBegin */
				state.push(RT); pass = true; break;
			case 0x0024: /* BrtFRTEnd */
				state.pop(); pass = false; break;
			default:
				if(R.T){/* empty */}
				else if(!pass || (opts.WTF && state[state.length-1] != 0x0023 /* BrtFRTBegin */)) throw new Error("Unexpected record 0x" + RT.toString(16));
		}
	});
	return out;
}
function write_xlmeta_bin() {
	var ba = buf_array();
	write_record(ba, 0x014C /* BrtBeginMetadata */);
	write_record(ba, 0x014E /* BrtBeginEsmdtinfo */, write_UInt32LE(1));
	write_record(ba, 0x014F /* BrtMdtinfo */, write_BrtMdtinfo({
		name: "XLDAPR",
		version: 120000,
		flags: 0xD06AC0B0
	}));
	write_record(ba, 0x0150 /* BrtEndEsmdtinfo */);
	/* [ESSTR] [ESMDX] */
	write_record(ba, 0x0153 /* BrtBeginEsfmd */, write_BrtBeginEsfmd(1, "XLDAPR"));
	write_record(ba, 0x0034 /* BrtBeginFmd */);
	write_record(ba, 0x0023 /* BrtFRTBegin */, write_UInt32LE(0x0202));
	write_record(ba, 0x1000 /* BrtBeginDynamicArrayPr */, write_UInt32LE(0));
	write_record(ba, 0x1001 /* BrtEndDynamicArrayPr */, writeuint16(1));
	write_record(ba, 0x0024 /* BrtFRTEnd */);
	write_record(ba, 0x0035 /* BrtEndFmd */);
	write_record(ba, 0x0154 /* BrtEndEsfmd */);
	write_record(ba, 0x0151 /* BrtBeginEsmdb */, write_BrtBeginEsmdb(1, true));
	write_record(ba, 0x0033 /* BrtMdb */, write_BrtMdb([[1, 0]]));
	write_record(ba, 0x0152 /* BrtEndEsmdb */);
	/* *FRT */
	write_record(ba, 0x014D /* BrtEndMetadata */);
	return ba.end();
}
