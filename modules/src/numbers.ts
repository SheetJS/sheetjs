/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { CFB$Container } from 'cfb';
import { WorkBook, WorkSheet, CellAddress, Range, CellObject } from '../../';
import { u8str, u8_to_dataview } from './util';
import { parse_shallow, varint_to_i32, parse_varint49, mappa } from './proto';
import { deframe } from './frame';
import { IWAArchiveInfo, IWAMessage, parse_iwa } from './iwa';

/* written here to avoid a full import of the 'xlsx' library */
var encode_col = (C: number): string => {
	var s="";
	for(++C; C; C=Math.floor((C-1)/26)) s = String.fromCharCode(((C-1)%26) + 65) + s;
	return s;
};
var encode_cell = (c: CellAddress): string => `${encode_col(c.c)}${c.r+1}`;
var encode_range = (r: Range): string => encode_cell(r.s) + ":" + encode_cell(r.e);
var book_new = (): WorkBook => ({Sheets:{}, SheetNames:[]});
var book_append_sheet = (wb: WorkBook, ws: WorkSheet, name?: string): void => {
	if(!name) for(var i = 1; i < 9999; ++i) if(wb.SheetNames.indexOf(name = `Sheet ${i}`) == -1) break;
	wb.SheetNames.push(name); wb.Sheets[name] = ws;
};

function parse_numbers(cfb: CFB$Container): WorkBook {
	var out: IWAMessage[][] = [];
	/* collect entire message space */
	cfb.FileIndex.forEach(s => {
		if(!s.name.match(/\.iwa$/)) return;
		var o: Uint8Array;
		try { o = deframe(s.content as Uint8Array); } catch(e) { return console.log("?? " + s.content.length + " " + (e.message || e)); }
		var packets: IWAArchiveInfo[];
		try { packets = parse_iwa(o); } catch(e) { return console.log("## " + (e.message || e)); }
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
	return parse_docroot(out, docroot);
}
export default parse_numbers;

// .TSP.Reference
function parse_Reference(buf: Uint8Array): number {
	var pb = parse_shallow(buf);
	return parse_varint49(pb[1][0].data);
}

// .TST.TableDataList
function parse_TST_TableDataList(M: IWAMessage[][], root: IWAMessage): string[] {
	var pb = parse_shallow(root.data);
	var entries = pb[3];
	var data = [];
	entries?.forEach(entry => {
		var le = parse_shallow(entry.data);
		var key = varint_to_i32(le[1][0].data)>>>0;
		data[key] = u8str(le[3][0].data);
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
	var storage = pb[3][0].data;
	var offsets = u8_to_dataview(pb[4][0].data);
	var cells = [];
	for(var C = 0; C < offsets.byteLength/2; ++C) {
		var off = offsets.getUint16(C*2, true);
		if(off > storage.length) continue;
		cells[C] = storage.subarray(off, offsets.getUint16(C*2+2, true));
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

		var sst = parse_TST_TableDataList(M, M[parse_Reference(store[4][0].data)][0]);
		{
			// .TST.TileStorage
			var tile = parse_shallow(store[3][0].data);
			var tiles: Array<{id: number, ref: Uint8Array[][]}> = [];
			tile[1].forEach(t => {
				var tl = (parse_shallow(t.data));
				var ref = M[parse_Reference(tl[2][0].data)][0];
				var mtype = varint_to_i32(ref.meta[1][0].data);
				if(mtype != 6002) throw new Error(`6001 unexpected reference to ${mtype}`);
				tiles.push({id: varint_to_i32(tl[1][0].data), ref: parse_TST_Tile(M, ref) });
			});
			tiles.forEach((tile) => {
				tile.ref.forEach((row, R) => {
					row.forEach((buf, C) => {
						var dv = u8_to_dataview(buf);
						//var version = buf[0]; // numbers 3.x use "3", 6.x - 11.x use "4"
						/* TODO: find the correct field position of the data type and value. */
						var ctype = buf[2];
						var addr = encode_cell({r:R,c:C});
						switch(ctype) {
							case 0: { // TODO: generic ??
								switch(buf[1]) {
									case 3: ws[addr] = { t: "s", v: sst[dv.getUint32(buf.length - 4,true)] } as CellObject; break;
									case 2: ws[addr] = { t: "n", v: dv.getFloat64(16, true) } as CellObject; break;
									case 0: break; // ws[addr] = { t: "z" } as CellObject; // blank?
									case 5: break; // date-time
									case 7: break; // duration
									case 6: ws[addr] = { t: "b", v: dv.getFloat64(buf.length - 8, true) > 0 } as CellObject; break;
									default: throw new Error(`Unsupported cell type ${buf.slice(0,4)}`);
								}
							} break;
							case 3: { // string
								ws[addr] = { t: "s", v: sst[dv.getUint32(16,true)] } as CellObject;
							} break;
							case 2: { // number
								ws[addr] = { t: "n", v: dv.getFloat64(buf.length - 12, true) } as CellObject;
							} break;
							case 6: { // boolean
								ws[addr] = { t: "b", v: dv.getFloat64(16, true) > 0 } as CellObject; // 1 or 0
							} break;
							default: throw new Error(`Unsupported cell type ${ctype}`);
						}
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
	var tableref = M[parse_Reference(pb[2][0].data)];
	var mtype = varint_to_i32(tableref[0].meta[1][0].data);
	if(mtype != 6001) throw new Error(`6000 unexpected reference to ${mtype}`);
	parse_TST_TableModelArchive(M, tableref[0], out);
	return out;
}

// .TN.SheetArchive
function parse_sheetroot(M: IWAMessage[][], root: IWAMessage): [WorkSheet, string] {
	var pb = parse_shallow(root.data);
	var out: [WorkSheet, string] = [ { "!ref":"A1" }, (pb[1]?.[0] ? u8str(pb[1][0].data) : "") ];
	var shapeoffs = mappa(pb[2], parse_Reference);
	var seen = false;
	shapeoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 6000) {
				if(seen) return; // TODO: multiple Tables in a sheet
				out[0] = parse_TST_TableInfoArchive(M, m);
				seen = true;
			}
		});
	});
	return out;
}

// .TN.DocumentArchive
function parse_docroot(M: IWAMessage[][], root: IWAMessage): WorkBook {
	var out = book_new();
	var pb = parse_shallow(root.data);
	var sheetoffs = mappa(pb[1], parse_Reference);
	sheetoffs.forEach((off) => {
		M[off].forEach((m: IWAMessage) => {
			var mtype = varint_to_i32(m.meta[1][0].data);
			if(mtype == 2) {
				var root = parse_sheetroot(M, m);
				book_append_sheet(out, root[0], root[1]);
			}
		});
	});
	if(out.SheetNames.length == 0) throw new Error("Empty NUMBERS file");
	return out;
}
