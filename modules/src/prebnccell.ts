/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { CellObject } from '../../';
import { u8_to_dataview, popcnt } from './util';

function parseit(buf: Uint8Array, version: number): CellObject {
  var dv = u8_to_dataview(buf);
  var ctype = buf[version == 4 ? 1 : 2];

  /* TODO: find the correct field position of number formats, formulae, etc */
  var flags = dv.getUint32(4, true);
  var data_offset = 12 + popcnt(flags & 0x3F8E) * 4;

  var sidx = -1, ieee = NaN, dt = NaN;
  if(flags & 0x10) { sidx = dv.getUint32(data_offset, true);  data_offset += 4; }
  if(flags & 0x20) { ieee = dv.getFloat64(data_offset, true); data_offset += 8; }
  if(flags & 0x40) { dt   = dv.getFloat64(data_offset, true); data_offset += 8; }

  var ret;
  switch(ctype) {
    case 0: break; // return { t: "z" }; // blank?
    case 2: ret = { t: "n", v: ieee }; break;
    case 3: ret = { t: "s", v: sidx }; break;
    case 5: var dd = new Date(2001, 0, 1); dd.setTime(dd.getTime() + dt * 1000); ret = { t: "d", v: dd }; break; // date-time TODO: relative or absolute?
    case 6: ret = { t: "b", v: ieee > 0 }; break;
    case 7: ret = { t: "n", v: ieee }; break; // duration in seconds TODO: emit [hh]:[mm] style format with adjusted value
    default: throw new Error(`Unsupported cell type ${buf.slice(0,4)}`);
  }
  /* TODO: Some fields appear after the cell data */

  return ret;
}

function parse(buf: Uint8Array): CellObject {
  var version = buf[0]; // numbers 3.5 uses "3", 6.x - 11.x use "4"
  switch(version) {
    case 3: case 4: return parseit(buf, version);
    default: throw new Error(`Unsupported pre-BNC version ${version}`);
  }
}

export { parse };