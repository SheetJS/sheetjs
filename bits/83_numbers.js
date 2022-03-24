/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
function u8_to_dataview(array) {
  return new DataView(array.buffer, array.byteOffset, array.byteLength);
}
function u8str(u8) {
  return typeof TextDecoder != "undefined" ? new TextDecoder().decode(u8) : utf8read(a2s(u8));
}
function stru8(str) {
  return typeof TextEncoder != "undefined" ? new TextEncoder().encode(str) : s2a(utf8write(str));
}
function u8contains(body, search) {
  outer:
    for (var L = 0; L <= body.length - search.length; ++L) {
      for (var j = 0; j < search.length; ++j)
        if (body[L + j] != search[j])
          continue outer;
      return true;
    }
  return false;
}
function u8concat(u8a) {
  var len = u8a.reduce(function(acc, x) {
    return acc + x.length;
  }, 0);
  var out = new Uint8Array(len);
  var off = 0;
  u8a.forEach(function(u8) {
    out.set(u8, off);
    off += u8.length;
  });
  return out;
}
function popcnt(x) {
  x -= x >> 1 & 1431655765;
  x = (x & 858993459) + (x >> 2 & 858993459);
  return (x + (x >> 4) & 252645135) * 16843009 >>> 24;
}
function readDecimal128LE(buf, offset) {
  var exp = (buf[offset + 15] & 127) << 7 | buf[offset + 14] >> 1;
  var mantissa = buf[offset + 14] & 1;
  for (var j = offset + 13; j >= offset; --j)
    mantissa = mantissa * 256 + buf[j];
  return (buf[offset + 15] & 128 ? -mantissa : mantissa) * Math.pow(10, exp - 6176);
}
function writeDecimal128LE(buf, offset, value) {
  var exp = Math.floor(value == 0 ? 0 : Math.LOG10E * Math.log(Math.abs(value))) + 6176 - 20;
  var mantissa = value / Math.pow(10, exp - 6176);
  buf[offset + 15] |= exp >> 7;
  buf[offset + 14] |= (exp & 127) << 1;
  for (var i = 0; mantissa >= 1; ++i, mantissa /= 256)
    buf[offset + i] = mantissa & 255;
  buf[offset + 15] |= value >= 0 ? 0 : 128;
}
function parse_varint49(buf, ptr) {
  var l = ptr ? ptr[0] : 0;
  var usz = buf[l] & 127;
  varint:
    if (buf[l++] >= 128) {
      usz |= (buf[l] & 127) << 7;
      if (buf[l++] < 128)
        break varint;
      usz |= (buf[l] & 127) << 14;
      if (buf[l++] < 128)
        break varint;
      usz |= (buf[l] & 127) << 21;
      if (buf[l++] < 128)
        break varint;
      usz += (buf[l] & 127) * Math.pow(2, 28);
      ++l;
      if (buf[l++] < 128)
        break varint;
      usz += (buf[l] & 127) * Math.pow(2, 35);
      ++l;
      if (buf[l++] < 128)
        break varint;
      usz += (buf[l] & 127) * Math.pow(2, 42);
      ++l;
      if (buf[l++] < 128)
        break varint;
    }
  if (ptr)
    ptr[0] = l;
  return usz;
}
function write_varint49(v) {
  var usz = new Uint8Array(7);
  usz[0] = v & 127;
  var L = 1;
  sz:
    if (v > 127) {
      usz[L - 1] |= 128;
      usz[L] = v >> 7 & 127;
      ++L;
      if (v <= 16383)
        break sz;
      usz[L - 1] |= 128;
      usz[L] = v >> 14 & 127;
      ++L;
      if (v <= 2097151)
        break sz;
      usz[L - 1] |= 128;
      usz[L] = v >> 21 & 127;
      ++L;
      if (v <= 268435455)
        break sz;
      usz[L - 1] |= 128;
      usz[L] = v / 256 >>> 21 & 127;
      ++L;
      if (v <= 34359738367)
        break sz;
      usz[L - 1] |= 128;
      usz[L] = v / 65536 >>> 21 & 127;
      ++L;
      if (v <= 4398046511103)
        break sz;
      usz[L - 1] |= 128;
      usz[L] = v / 16777216 >>> 21 & 127;
      ++L;
    }
  return usz.slice(0, L);
}
function varint_to_i32(buf) {
  var l = 0, i32 = buf[l] & 127;
  varint:
    if (buf[l++] >= 128) {
      i32 |= (buf[l] & 127) << 7;
      if (buf[l++] < 128)
        break varint;
      i32 |= (buf[l] & 127) << 14;
      if (buf[l++] < 128)
        break varint;
      i32 |= (buf[l] & 127) << 21;
      if (buf[l++] < 128)
        break varint;
      i32 |= (buf[l] & 127) << 28;
    }
  return i32;
}
function parse_shallow(buf) {
  var out = [], ptr = [0];
  while (ptr[0] < buf.length) {
    var off = ptr[0];
    var num = parse_varint49(buf, ptr);
    var type = num & 7;
    num = Math.floor(num / 8);
    var len = 0;
    var res;
    if (num == 0)
      break;
    switch (type) {
      case 0:
        {
          var l = ptr[0];
          while (buf[ptr[0]++] >= 128)
            ;
          res = buf.slice(l, ptr[0]);
        }
        break;
      case 5:
        len = 4;
        res = buf.slice(ptr[0], ptr[0] + len);
        ptr[0] += len;
        break;
      case 1:
        len = 8;
        res = buf.slice(ptr[0], ptr[0] + len);
        ptr[0] += len;
        break;
      case 2:
        len = parse_varint49(buf, ptr);
        res = buf.slice(ptr[0], ptr[0] + len);
        ptr[0] += len;
        break;
      case 3:
      case 4:
      default:
        throw new Error("PB Type ".concat(type, " for Field ").concat(num, " at offset ").concat(off));
    }
    var v = { data: res, type: type };
    if (out[num] == null)
      out[num] = [v];
    else
      out[num].push(v);
  }
  return out;
}
function write_shallow(proto) {
  var out = [];
  proto.forEach(function(field, idx) {
    field.forEach(function(item) {
      if (!item.data)
        return;
      out.push(write_varint49(idx * 8 + item.type));
      if (item.type == 2)
        out.push(write_varint49(item.data.length));
      out.push(item.data);
    });
  });
  return u8concat(out);
}
function mappa(data, cb) {
  return (data == null ? void 0 : data.map(function(d) {
    return cb(d.data);
  })) || [];
}
function parse_iwa_file(buf) {
  var _a;
  var out = [], ptr = [0];
  while (ptr[0] < buf.length) {
    var len = parse_varint49(buf, ptr);
    var ai = parse_shallow(buf.slice(ptr[0], ptr[0] + len));
    ptr[0] += len;
    var res = {
      id: varint_to_i32(ai[1][0].data),
      messages: []
    };
    ai[2].forEach(function(b) {
      var mi = parse_shallow(b.data);
      var fl = varint_to_i32(mi[3][0].data);
      res.messages.push({
        meta: mi,
        data: buf.slice(ptr[0], ptr[0] + fl)
      });
      ptr[0] += fl;
    });
    if ((_a = ai[3]) == null ? void 0 : _a[0])
      res.merge = varint_to_i32(ai[3][0].data) >>> 0 > 0;
    out.push(res);
  }
  return out;
}
function write_iwa_file(ias) {
  var bufs = [];
  ias.forEach(function(ia) {
    var ai = [];
    ai[1] = [{ data: write_varint49(ia.id), type: 0 }];
    ai[2] = [];
    if (ia.merge != null)
      ai[3] = [{ data: write_varint49(+!!ia.merge), type: 0 }];
    var midata = [];
    ia.messages.forEach(function(mi) {
      midata.push(mi.data);
      mi.meta[3] = [{ type: 0, data: write_varint49(mi.data.length) }];
      ai[2].push({ data: write_shallow(mi.meta), type: 2 });
    });
    var aipayload = write_shallow(ai);
    bufs.push(write_varint49(aipayload.length));
    bufs.push(aipayload);
    midata.forEach(function(mid) {
      return bufs.push(mid);
    });
  });
  return u8concat(bufs);
}
function parse_snappy_chunk(type, buf) {
  if (type != 0)
    throw new Error("Unexpected Snappy chunk type ".concat(type));
  var ptr = [0];
  var usz = parse_varint49(buf, ptr);
  var chunks = [];
  while (ptr[0] < buf.length) {
    var tag = buf[ptr[0]] & 3;
    if (tag == 0) {
      var len = buf[ptr[0]++] >> 2;
      if (len < 60)
        ++len;
      else {
        var c = len - 59;
        len = buf[ptr[0]];
        if (c > 1)
          len |= buf[ptr[0] + 1] << 8;
        if (c > 2)
          len |= buf[ptr[0] + 2] << 16;
        if (c > 3)
          len |= buf[ptr[0] + 3] << 24;
        len >>>= 0;
        len++;
        ptr[0] += c;
      }
      chunks.push(buf.slice(ptr[0], ptr[0] + len));
      ptr[0] += len;
      continue;
    } else {
      var offset = 0, length = 0;
      if (tag == 1) {
        length = (buf[ptr[0]] >> 2 & 7) + 4;
        offset = (buf[ptr[0]++] & 224) << 3;
        offset |= buf[ptr[0]++];
      } else {
        length = (buf[ptr[0]++] >> 2) + 1;
        if (tag == 2) {
          offset = buf[ptr[0]] | buf[ptr[0] + 1] << 8;
          ptr[0] += 2;
        } else {
          offset = (buf[ptr[0]] | buf[ptr[0] + 1] << 8 | buf[ptr[0] + 2] << 16 | buf[ptr[0] + 3] << 24) >>> 0;
          ptr[0] += 4;
        }
      }
      chunks = [u8concat(chunks)];
      if (offset == 0)
        throw new Error("Invalid offset 0");
      if (offset > chunks[0].length)
        throw new Error("Invalid offset beyond length");
      if (length >= offset) {
        chunks.push(chunks[0].slice(-offset));
        length -= offset;
        while (length >= chunks[chunks.length - 1].length) {
          chunks.push(chunks[chunks.length - 1]);
          length -= chunks[chunks.length - 1].length;
        }
      }
      chunks.push(chunks[0].slice(-offset, -offset + length));
    }
  }
  var o = u8concat(chunks);
  if (o.length != usz)
    throw new Error("Unexpected length: ".concat(o.length, " != ").concat(usz));
  return o;
}
function decompress_iwa_file(buf) {
  var out = [];
  var l = 0;
  while (l < buf.length) {
    var t = buf[l++];
    var len = buf[l] | buf[l + 1] << 8 | buf[l + 2] << 16;
    l += 3;
    out.push(parse_snappy_chunk(t, buf.slice(l, l + len)));
    l += len;
  }
  if (l !== buf.length)
    throw new Error("data is not a valid framed stream!");
  return u8concat(out);
}
function compress_iwa_file(buf) {
  var out = [];
  var l = 0;
  while (l < buf.length) {
    var c = Math.min(buf.length - l, 268435455);
    var frame = new Uint8Array(4);
    out.push(frame);
    var usz = write_varint49(c);
    var L = usz.length;
    out.push(usz);
    if (c <= 60) {
      L++;
      out.push(new Uint8Array([c - 1 << 2]));
    } else if (c <= 256) {
      L += 2;
      out.push(new Uint8Array([240, c - 1 & 255]));
    } else if (c <= 65536) {
      L += 3;
      out.push(new Uint8Array([244, c - 1 & 255, c - 1 >> 8 & 255]));
    } else if (c <= 16777216) {
      L += 4;
      out.push(new Uint8Array([248, c - 1 & 255, c - 1 >> 8 & 255, c - 1 >> 16 & 255]));
    } else if (c <= 4294967296) {
      L += 5;
      out.push(new Uint8Array([252, c - 1 & 255, c - 1 >> 8 & 255, c - 1 >> 16 & 255, c - 1 >>> 24 & 255]));
    }
    out.push(buf.slice(l, l + c));
    L += c;
    frame[0] = 0;
    frame[1] = L & 255;
    frame[2] = L >> 8 & 255;
    frame[3] = L >> 16 & 255;
    l += c;
  }
  return u8concat(out);
}
function parse_old_storage(buf, sst, rsst, v) {
  var dv = u8_to_dataview(buf);
  var flags = dv.getUint32(4, true);
  var data_offset = (v > 1 ? 12 : 8) + popcnt(flags & (v > 1 ? 3470 : 398)) * 4;
  var ridx = -1, sidx = -1, ieee = NaN, dt = new Date(2001, 0, 1);
  if (flags & 512) {
    ridx = dv.getUint32(data_offset, true);
    data_offset += 4;
  }
  data_offset += popcnt(flags & (v > 1 ? 12288 : 4096)) * 4;
  if (flags & 16) {
    sidx = dv.getUint32(data_offset, true);
    data_offset += 4;
  }
  if (flags & 32) {
    ieee = dv.getFloat64(data_offset, true);
    data_offset += 8;
  }
  if (flags & 64) {
    dt.setTime(dt.getTime() + dv.getFloat64(data_offset, true) * 1e3);
    data_offset += 8;
  }
  var ret;
  switch (buf[2]) {
    case 0:
      break;
    case 2:
      ret = { t: "n", v: ieee };
      break;
    case 3:
      ret = { t: "s", v: sst[sidx] };
      break;
    case 5:
      ret = { t: "d", v: dt };
      break;
    case 6:
      ret = { t: "b", v: ieee > 0 };
      break;
    case 7:
      ret = { t: "n", v: ieee / 86400 };
      break;
    case 8:
      ret = { t: "e", v: 0 };
      break;
    case 9:
      {
        if (ridx > -1)
          ret = { t: "s", v: rsst[ridx] };
        else if (sidx > -1)
          ret = { t: "s", v: sst[sidx] };
        else if (!isNaN(ieee))
          ret = { t: "n", v: ieee };
        else
          throw new Error("Unsupported cell type ".concat(buf.slice(0, 4)));
      }
      break;
    default:
      throw new Error("Unsupported cell type ".concat(buf.slice(0, 4)));
  }
  return ret;
}
function parse_new_storage(buf, sst, rsst) {
  var dv = u8_to_dataview(buf);
  var flags = dv.getUint32(8, true);
  var data_offset = 12;
  var ridx = -1, sidx = -1, d128 = NaN, ieee = NaN, dt = new Date(2001, 0, 1);
  if (flags & 1) {
    d128 = readDecimal128LE(buf, data_offset);
    data_offset += 16;
  }
  if (flags & 2) {
    ieee = dv.getFloat64(data_offset, true);
    data_offset += 8;
  }
  if (flags & 4) {
    dt.setTime(dt.getTime() + dv.getFloat64(data_offset, true) * 1e3);
    data_offset += 8;
  }
  if (flags & 8) {
    sidx = dv.getUint32(data_offset, true);
    data_offset += 4;
  }
  if (flags & 16) {
    ridx = dv.getUint32(data_offset, true);
    data_offset += 4;
  }
  var ret;
  switch (buf[1]) {
    case 0:
      break;
    case 2:
      ret = { t: "n", v: d128 };
      break;
    case 3:
      ret = { t: "s", v: sst[sidx] };
      break;
    case 5:
      ret = { t: "d", v: dt };
      break;
    case 6:
      ret = { t: "b", v: ieee > 0 };
      break;
    case 7:
      ret = { t: "n", v: ieee / 86400 };
      break;
    case 8:
      ret = { t: "e", v: 0 };
      break;
    case 9:
      {
        if (ridx > -1)
          ret = { t: "s", v: rsst[ridx] };
        else
          throw new Error("Unsupported cell type ".concat(buf[1], " : ").concat(flags & 31, " : ").concat(buf.slice(0, 4)));
      }
      break;
    case 10:
      ret = { t: "n", v: d128 };
      break;
    default:
      throw new Error("Unsupported cell type ".concat(buf[1], " : ").concat(flags & 31, " : ").concat(buf.slice(0, 4)));
  }
  return ret;
}
function write_new_storage(cell, sst) {
  var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
  out[0] = 5;
  switch (cell.t) {
    case "n":
      out[1] = 2;
      writeDecimal128LE(out, l, cell.v);
      flags |= 1;
      l += 16;
      break;
    case "b":
      out[1] = 6;
      dv.setFloat64(l, cell.v ? 1 : 0, true);
      flags |= 2;
      l += 8;
      break;
    case "s":
      if (sst.indexOf(cell.v) == -1)
        throw new Error("Value ".concat(cell.v, " missing from SST!"));
      out[1] = 3;
      dv.setUint32(l, sst.indexOf(cell.v), true);
      flags |= 8;
      l += 4;
      break;
    default:
      throw "unsupported cell type " + cell.t;
  }
  dv.setUint32(8, flags, true);
  return out.slice(0, l);
}
function write_old_storage(cell, sst) {
  var out = new Uint8Array(32), dv = u8_to_dataview(out), l = 12, flags = 0;
  out[0] = 3;
  switch (cell.t) {
    case "n":
      out[2] = 2;
      dv.setFloat64(l, cell.v, true);
      flags |= 32;
      l += 8;
      break;
    case "b":
      out[2] = 6;
      dv.setFloat64(l, cell.v ? 1 : 0, true);
      flags |= 32;
      l += 8;
      break;
    case "s":
      if (sst.indexOf(cell.v) == -1)
        throw new Error("Value ".concat(cell.v, " missing from SST!"));
      out[2] = 3;
      dv.setUint32(l, sst.indexOf(cell.v), true);
      flags |= 16;
      l += 4;
      break;
    default:
      throw "unsupported cell type " + cell.t;
  }
  dv.setUint32(4, flags, true);
  return out.slice(0, l);
}
function parse_cell_storage(buf, sst, rsst) {
  switch (buf[0]) {
    case 0:
    case 1:
    case 2:
    case 3:
      return parse_old_storage(buf, sst, rsst, buf[0]);
    case 5:
      return parse_new_storage(buf, sst, rsst);
    default:
      throw new Error("Unsupported payload version ".concat(buf[0]));
  }
}
function parse_TSP_Reference(buf) {
  var pb = parse_shallow(buf);
  return parse_varint49(pb[1][0].data);
}
function write_TSP_Reference(idx) {
  var out = [];
  out[1] = [{ type: 0, data: write_varint49(idx) }];
  return write_shallow(out);
}
function parse_TST_TableDataList(M, root) {
  var pb = parse_shallow(root.data);
  var type = varint_to_i32(pb[1][0].data);
  var entries = pb[3];
  var data = [];
  (entries || []).forEach(function(entry) {
    var le = parse_shallow(entry.data);
    var key = varint_to_i32(le[1][0].data) >>> 0;
    switch (type) {
      case 1:
        data[key] = u8str(le[3][0].data);
        break;
      case 8:
        {
          var rt = M[parse_TSP_Reference(le[9][0].data)][0];
          var rtp = parse_shallow(rt.data);
          var rtpref = M[parse_TSP_Reference(rtp[1][0].data)][0];
          var mtype = varint_to_i32(rtpref.meta[1][0].data);
          if (mtype != 2001)
            throw new Error("2000 unexpected reference to ".concat(mtype));
          var tswpsa = parse_shallow(rtpref.data);
          data[key] = tswpsa[3].map(function(x) {
            return u8str(x.data);
          }).join("");
        }
        break;
    }
  });
  return data;
}
function parse_TST_TileRowInfo(u8, type) {
  var _a, _b, _c, _d, _e, _f, _g, _h, _i, _j, _k, _l, _m, _n;
  var pb = parse_shallow(u8);
  var R = varint_to_i32(pb[1][0].data) >>> 0;
  var cnt = varint_to_i32(pb[2][0].data) >>> 0;
  var wide_offsets = ((_b = (_a = pb[8]) == null ? void 0 : _a[0]) == null ? void 0 : _b.data) && varint_to_i32(pb[8][0].data) > 0 || false;
  var used_storage_u8, used_storage;
  if (((_d = (_c = pb[7]) == null ? void 0 : _c[0]) == null ? void 0 : _d.data) && type != 0) {
    used_storage_u8 = (_f = (_e = pb[7]) == null ? void 0 : _e[0]) == null ? void 0 : _f.data;
    used_storage = (_h = (_g = pb[6]) == null ? void 0 : _g[0]) == null ? void 0 : _h.data;
  } else if (((_j = (_i = pb[4]) == null ? void 0 : _i[0]) == null ? void 0 : _j.data) && type != 1) {
    used_storage_u8 = (_l = (_k = pb[4]) == null ? void 0 : _k[0]) == null ? void 0 : _l.data;
    used_storage = (_n = (_m = pb[3]) == null ? void 0 : _m[0]) == null ? void 0 : _n.data;
  } else
    throw "NUMBERS Tile missing ".concat(type, " cell storage");
  var width = wide_offsets ? 4 : 1;
  var used_storage_offsets = u8_to_dataview(used_storage_u8);
  var offsets = [];
  for (var C = 0; C < used_storage_u8.length / 2; ++C) {
    var off = used_storage_offsets.getUint16(C * 2, true);
    if (off < 65535)
      offsets.push([C, off]);
  }
  if (offsets.length != cnt)
    throw "Expected ".concat(cnt, " cells, found ").concat(offsets.length);
  var cells = [];
  for (C = 0; C < offsets.length - 1; ++C)
    cells[offsets[C][0]] = used_storage.subarray(offsets[C][1] * width, offsets[C + 1][1] * width);
  if (offsets.length >= 1)
    cells[offsets[offsets.length - 1][0]] = used_storage.subarray(offsets[offsets.length - 1][1] * width);
  return { R: R, cells: cells };
}
function parse_TST_Tile(M, root) {
  var _a;
  var pb = parse_shallow(root.data);
  var storage = ((_a = pb == null ? void 0 : pb[7]) == null ? void 0 : _a[0]) ? varint_to_i32(pb[7][0].data) >>> 0 > 0 ? 1 : 0 : -1;
  var ri = mappa(pb[5], function(u8) {
    return parse_TST_TileRowInfo(u8, storage);
  });
  return {
    nrows: varint_to_i32(pb[4][0].data) >>> 0,
    data: ri.reduce(function(acc, x) {
      if (!acc[x.R])
        acc[x.R] = [];
      x.cells.forEach(function(cell, C) {
        if (acc[x.R][C])
          throw new Error("Duplicate cell r=".concat(x.R, " c=").concat(C));
        acc[x.R][C] = cell;
      });
      return acc;
    }, [])
  };
}
function parse_TST_TableModelArchive(M, root, ws) {
  var _a;
  var pb = parse_shallow(root.data);
  var range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  range.e.r = (varint_to_i32(pb[6][0].data) >>> 0) - 1;
  if (range.e.r < 0)
    throw new Error("Invalid row varint ".concat(pb[6][0].data));
  range.e.c = (varint_to_i32(pb[7][0].data) >>> 0) - 1;
  if (range.e.c < 0)
    throw new Error("Invalid col varint ".concat(pb[7][0].data));
  ws["!ref"] = encode_range(range);
  var store = parse_shallow(pb[4][0].data);
  var sst = parse_TST_TableDataList(M, M[parse_TSP_Reference(store[4][0].data)][0]);
  var rsst = ((_a = store[17]) == null ? void 0 : _a[0]) ? parse_TST_TableDataList(M, M[parse_TSP_Reference(store[17][0].data)][0]) : [];
  var tile = parse_shallow(store[3][0].data);
  var _R = 0;
  tile[1].forEach(function(t) {
    var tl = parse_shallow(t.data);
    var ref = M[parse_TSP_Reference(tl[2][0].data)][0];
    var mtype = varint_to_i32(ref.meta[1][0].data);
    if (mtype != 6002)
      throw new Error("6001 unexpected reference to ".concat(mtype));
    var _tile = parse_TST_Tile(M, ref);
    _tile.data.forEach(function(row, R) {
      row.forEach(function(buf, C) {
        var addr = encode_cell({ r: _R + R, c: C });
        var res = parse_cell_storage(buf, sst, rsst);
        if (res)
          ws[addr] = res;
      });
    });
    _R += _tile.nrows;
  });
}
function parse_TST_TableInfoArchive(M, root) {
  var pb = parse_shallow(root.data);
  var out = { "!ref": "A1" };
  var tableref = M[parse_TSP_Reference(pb[2][0].data)];
  var mtype = varint_to_i32(tableref[0].meta[1][0].data);
  if (mtype != 6001)
    throw new Error("6000 unexpected reference to ".concat(mtype));
  parse_TST_TableModelArchive(M, tableref[0], out);
  return out;
}
function parse_TN_SheetArchive(M, root) {
  var _a;
  var pb = parse_shallow(root.data);
  var out = {
    name: ((_a = pb[1]) == null ? void 0 : _a[0]) ? u8str(pb[1][0].data) : "",
    sheets: []
  };
  var shapeoffs = mappa(pb[2], parse_TSP_Reference);
  shapeoffs.forEach(function(off) {
    M[off].forEach(function(m) {
      var mtype = varint_to_i32(m.meta[1][0].data);
      if (mtype == 6e3)
        out.sheets.push(parse_TST_TableInfoArchive(M, m));
    });
  });
  return out;
}
function parse_TN_DocumentArchive(M, root) {
  var out = book_new();
  var pb = parse_shallow(root.data);
  var sheetoffs = mappa(pb[1], parse_TSP_Reference);
  sheetoffs.forEach(function(off) {
    M[off].forEach(function(m) {
      var mtype = varint_to_i32(m.meta[1][0].data);
      if (mtype == 2) {
        var root2 = parse_TN_SheetArchive(M, m);
        root2.sheets.forEach(function(sheet, idx) {
          book_append_sheet(out, sheet, idx == 0 ? root2.name : root2.name + "_" + idx, true);
        });
      }
    });
  });
  if (out.SheetNames.length == 0)
    throw new Error("Empty NUMBERS file");
  return out;
}
function parse_numbers_iwa(cfb) {
  var _a, _b, _c, _d;
  var M = {}, indices = [];
  cfb.FullPaths.forEach(function(p) {
    if (p.match(/\.iwpv2/))
      throw new Error("Unsupported password protection");
  });
  cfb.FileIndex.forEach(function(s) {
    if (!s.name.match(/\.iwa$/))
      return;
    var o;
    try {
      o = decompress_iwa_file(s.content);
    } catch (e) {
      return console.log("?? " + s.content.length + " " + (e.message || e));
    }
    var packets;
    try {
      packets = parse_iwa_file(o);
    } catch (e) {
      return console.log("## " + (e.message || e));
    }
    packets.forEach(function(packet) {
      M[packet.id] = packet.messages;
      indices.push(packet.id);
    });
  });
  if (!indices.length)
    throw new Error("File has no messages");
  var docroot = ((_d = (_c = (_b = (_a = M == null ? void 0 : M[1]) == null ? void 0 : _a[0]) == null ? void 0 : _b.meta) == null ? void 0 : _c[1]) == null ? void 0 : _d[0].data) && varint_to_i32(M[1][0].meta[1][0].data) == 1 && M[1][0];
  if (!docroot)
    indices.forEach(function(idx) {
      M[idx].forEach(function(iwam) {
        var mtype = varint_to_i32(iwam.meta[1][0].data) >>> 0;
        if (mtype == 1) {
          if (!docroot)
            docroot = iwam;
          else
            throw new Error("Document has multiple roots");
        }
      });
    });
  if (!docroot)
    throw new Error("Cannot find Document root");
  return parse_TN_DocumentArchive(M, docroot);
}
function write_tile_row(tri, data, SST) {
  var _a, _b, _c, _d;
  if (!((_a = tri[6]) == null ? void 0 : _a[0]) || !((_b = tri[7]) == null ? void 0 : _b[0]))
    throw "Mutation only works on post-BNC storages!";
  var wide_offsets = ((_d = (_c = tri[8]) == null ? void 0 : _c[0]) == null ? void 0 : _d.data) && varint_to_i32(tri[8][0].data) > 0 || false;
  if (wide_offsets)
    throw "Math only works with normal offsets";
  var cnt = 0;
  var dv = u8_to_dataview(tri[7][0].data), last_offset = 0, cell_storage = [];
  var _dv = u8_to_dataview(tri[4][0].data), _last_offset = 0, _cell_storage = [];
  for (var C = 0; C < data.length; ++C) {
    if (data[C] == null) {
      dv.setUint16(C * 2, 65535, true);
      _dv.setUint16(C * 2, 65535);
      continue;
    }
    dv.setUint16(C * 2, last_offset, true);
    _dv.setUint16(C * 2, _last_offset, true);
    var celload, _celload;
    switch (typeof data[C]) {
      case "string":
        celload = write_new_storage({ t: "s", v: data[C] }, SST);
        _celload = write_old_storage({ t: "s", v: data[C] }, SST);
        break;
      case "number":
        celload = write_new_storage({ t: "n", v: data[C] }, SST);
        _celload = write_old_storage({ t: "n", v: data[C] }, SST);
        break;
      case "boolean":
        celload = write_new_storage({ t: "b", v: data[C] }, SST);
        _celload = write_old_storage({ t: "b", v: data[C] }, SST);
        break;
      default:
        throw new Error("Unsupported value " + data[C]);
    }
    cell_storage.push(celload);
    last_offset += celload.length;
    _cell_storage.push(_celload);
    _last_offset += _celload.length;
    ++cnt;
  }
  tri[2][0].data = write_varint49(cnt);
  for (; C < tri[7][0].data.length / 2; ++C) {
    dv.setUint16(C * 2, 65535, true);
    _dv.setUint16(C * 2, 65535, true);
  }
  tri[6][0].data = u8concat(cell_storage);
  tri[3][0].data = u8concat(_cell_storage);
  return cnt;
}
function write_numbers_iwa(wb, opts) {
  if (!opts || !opts.numbers)
    throw new Error("Must pass a `numbers` option -- check the README");
  var ws = wb.Sheets[wb.SheetNames[0]];
  if (wb.SheetNames.length > 1)
    console.error("The Numbers writer currently writes only the first table");
  var range = decode_range(ws["!ref"]);
  range.s.r = range.s.c = 0;
  var trunc = false;
  if (range.e.c > 9) {
    trunc = true;
    range.e.c = 9;
  }
  if (range.e.r > 49) {
    trunc = true;
    range.e.r = 49;
  }
  if (trunc)
    console.error("The Numbers writer is currently limited to ".concat(encode_range(range)));
  var data = sheet_to_json(ws, { range: range, header: 1 });
  var SST = ["~Sh33tJ5~"];
  data.forEach(function(row) {
    return row.forEach(function(cell) {
      if (typeof cell == "string")
        SST.push(cell);
    });
  });
  var dependents = {};
  var indices = [];
  var cfb = CFB.read(opts.numbers, { type: "base64" });
  cfb.FileIndex.map(function(fi, idx) {
    return [fi, cfb.FullPaths[idx]];
  }).forEach(function(row) {
    var fi = row[0], fp = row[1];
    if (fi.type != 2)
      return;
    if (!fi.name.match(/\.iwa/))
      return;
    var old_content = fi.content;
    var raw1 = decompress_iwa_file(old_content);
    var x2 = parse_iwa_file(raw1);
    x2.forEach(function(packet2) {
      indices.push(packet2.id);
      dependents[packet2.id] = { deps: [], location: fp, type: varint_to_i32(packet2.messages[0].meta[1][0].data) };
    });
  });
  indices.sort(function(x2, y2) {
    return x2 - y2;
  });
  var indices_varint = indices.filter(function(x2) {
    return x2 > 1;
  }).map(function(x2) {
    return [x2, write_varint49(x2)];
  });
  cfb.FileIndex.map(function(fi, idx) {
    return [fi, cfb.FullPaths[idx]];
  }).forEach(function(row) {
    var fi = row[0], fp = row[1];
    if (!fi.name.match(/\.iwa/))
      return;
    var x2 = parse_iwa_file(decompress_iwa_file(fi.content));
    x2.forEach(function(ia) {
      ia.messages.forEach(function(m) {
        indices_varint.forEach(function(ivi) {
          if (ia.messages.some(function(mess) {
            return varint_to_i32(mess.meta[1][0].data) != 11006 && u8contains(mess.data, ivi[1]);
          })) {
            dependents[ivi[0]].deps.push(ia.id);
          }
        });
      });
    });
  });
  function get_unique_msgid() {
    for (var i = 927262; i < 2e6; ++i)
      if (!dependents[i])
        return i;
    throw new Error("Too many messages");
  }
  var entry = CFB.find(cfb, dependents[1].location);
  var x = parse_iwa_file(decompress_iwa_file(entry.content));
  var docroot;
  for (var xi = 0; xi < x.length; ++xi) {
    var packet = x[xi];
    if (packet.id == 1)
      docroot = packet;
  }
  var sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[1][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content));
  for (xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if (packet.id == sheetrootref)
      docroot = packet;
  }
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content));
  for (xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if (packet.id == sheetrootref)
      docroot = packet;
  }
  sheetrootref = parse_TSP_Reference(parse_shallow(docroot.messages[0].data)[2][0].data);
  entry = CFB.find(cfb, dependents[sheetrootref].location);
  x = parse_iwa_file(decompress_iwa_file(entry.content));
  for (xi = 0; xi < x.length; ++xi) {
    packet = x[xi];
    if (packet.id == sheetrootref)
      docroot = packet;
  }
  var pb = parse_shallow(docroot.messages[0].data);
  {
    pb[6][0].data = write_varint49(range.e.r + 1);
    pb[7][0].data = write_varint49(range.e.c + 1);
    var cruidsref = parse_TSP_Reference(pb[46][0].data);
    var oldbucket = CFB.find(cfb, dependents[cruidsref].location);
    var _x = parse_iwa_file(decompress_iwa_file(oldbucket.content));
    {
      for (var j = 0; j < _x.length; ++j) {
        if (_x[j].id == cruidsref)
          break;
      }
      if (_x[j].id != cruidsref)
        throw "Bad ColumnRowUIDMapArchive";
      var cruids = parse_shallow(_x[j].messages[0].data);
      cruids[1] = [];
      cruids[2] = [], cruids[3] = [];
      for (var C = 0; C <= range.e.c; ++C) {
        var uuid = [];
        uuid[1] = uuid[2] = [{ type: 0, data: write_varint49(C + 420690) }];
        cruids[1].push({ type: 2, data: write_shallow(uuid) });
        cruids[2].push({ type: 0, data: write_varint49(C) });
        cruids[3].push({ type: 0, data: write_varint49(C) });
      }
      cruids[4] = [];
      cruids[5] = [], cruids[6] = [];
      for (var R = 0; R <= range.e.r; ++R) {
        uuid = [];
        uuid[1] = uuid[2] = [{ type: 0, data: write_varint49(R + 726270) }];
        cruids[4].push({ type: 2, data: write_shallow(uuid) });
        cruids[5].push({ type: 0, data: write_varint49(R) });
        cruids[6].push({ type: 0, data: write_varint49(R) });
      }
      _x[j].messages[0].data = write_shallow(cruids);
    }
    oldbucket.content = compress_iwa_file(write_iwa_file(_x));
    oldbucket.size = oldbucket.content.length;
    delete pb[46];
    var store = parse_shallow(pb[4][0].data);
    {
      store[7][0].data = write_varint49(range.e.r + 1);
      var row_headers = parse_shallow(store[1][0].data);
      var row_header_ref = parse_TSP_Reference(row_headers[2][0].data);
      oldbucket = CFB.find(cfb, dependents[row_header_ref].location);
      _x = parse_iwa_file(decompress_iwa_file(oldbucket.content));
      {
        if (_x[0].id != row_header_ref)
          throw "Bad HeaderStorageBucket";
        var base_bucket = parse_shallow(_x[0].messages[0].data);
        for (R = 0; R < data.length; ++R) {
          var _bucket = parse_shallow(base_bucket[2][0].data);
          _bucket[1][0].data = write_varint49(R);
          _bucket[4][0].data = write_varint49(data[R].length);
          base_bucket[2][R] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
        }
        _x[0].messages[0].data = write_shallow(base_bucket);
      }
      oldbucket.content = compress_iwa_file(write_iwa_file(_x));
      oldbucket.size = oldbucket.content.length;
      var col_header_ref = parse_TSP_Reference(store[2][0].data);
      oldbucket = CFB.find(cfb, dependents[col_header_ref].location);
      _x = parse_iwa_file(decompress_iwa_file(oldbucket.content));
      {
        if (_x[0].id != col_header_ref)
          throw "Bad HeaderStorageBucket";
        base_bucket = parse_shallow(_x[0].messages[0].data);
        for (C = 0; C <= range.e.c; ++C) {
          _bucket = parse_shallow(base_bucket[2][0].data);
          _bucket[1][0].data = write_varint49(C);
          _bucket[4][0].data = write_varint49(range.e.r + 1);
          base_bucket[2][C] = { type: base_bucket[2][0].type, data: write_shallow(_bucket) };
        }
        _x[0].messages[0].data = write_shallow(base_bucket);
      }
      oldbucket.content = compress_iwa_file(write_iwa_file(_x));
      oldbucket.size = oldbucket.content.length;
      var sstref = parse_TSP_Reference(store[4][0].data);
      (function() {
        var sentry = CFB.find(cfb, dependents[sstref].location);
        var sx = parse_iwa_file(decompress_iwa_file(sentry.content));
        var sstroot;
        for (var sxi = 0; sxi < sx.length; ++sxi) {
          var packet2 = sx[sxi];
          if (packet2.id == sstref)
            sstroot = packet2;
        }
        var sstdata = parse_shallow(sstroot.messages[0].data);
        {
          sstdata[3] = [];
          var newsst = [];
          SST.forEach(function(str, i) {
            newsst[1] = [{ type: 0, data: write_varint49(i) }];
            newsst[2] = [{ type: 0, data: write_varint49(1) }];
            newsst[3] = [{ type: 2, data: stru8(str) }];
            sstdata[3].push({ type: 2, data: write_shallow(newsst) });
          });
        }
        sstroot.messages[0].data = write_shallow(sstdata);
        var sy = write_iwa_file(sx);
        var raw32 = compress_iwa_file(sy);
        sentry.content = raw32;
        sentry.size = sentry.content.length;
      })();
      var tile = parse_shallow(store[3][0].data);
      {
        var t = tile[1][0];
        delete tile[2];
        var tl = parse_shallow(t.data);
        {
          var tileref = parse_TSP_Reference(tl[2][0].data);
          (function() {
            var tentry = CFB.find(cfb, dependents[tileref].location);
            var tx = parse_iwa_file(decompress_iwa_file(tentry.content));
            var tileroot;
            for (var sxi = 0; sxi < tx.length; ++sxi) {
              var packet2 = tx[sxi];
              if (packet2.id == tileref)
                tileroot = packet2;
            }
            var tiledata = parse_shallow(tileroot.messages[0].data);
            {
              delete tiledata[6];
              delete tile[7];
              var rowload = new Uint8Array(tiledata[5][0].data);
              tiledata[5] = [];
              var cnt = 0;
              for (var R2 = 0; R2 <= range.e.r; ++R2) {
                var tilerow = parse_shallow(rowload);
                cnt += write_tile_row(tilerow, data[R2], SST);
                tilerow[1][0].data = write_varint49(R2);
                tiledata[5].push({ data: write_shallow(tilerow), type: 2 });
              }
              tiledata[1] = [{ type: 0, data: write_varint49(range.e.c + 1) }];
              tiledata[2] = [{ type: 0, data: write_varint49(range.e.r + 1) }];
              tiledata[3] = [{ type: 0, data: write_varint49(cnt) }];
              tiledata[4] = [{ type: 0, data: write_varint49(range.e.r + 1) }];
            }
            tileroot.messages[0].data = write_shallow(tiledata);
            var ty = write_iwa_file(tx);
            var raw32 = compress_iwa_file(ty);
            tentry.content = raw32;
            tentry.size = tentry.content.length;
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
  entry.content = raw3;
  entry.size = entry.content.length;
  return cfb;
}
