var NUMBERS = (function() {
  var __defProp = Object.defineProperty;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __markAsModule = function(target) {
    return __defProp(target, "__esModule", { value: true });
  };
  var __export = function(target, all) {
    for (var name in all)
      __defProp(target, name, { get: all[name], enumerable: true });
  };
  var __reExport = function(target, module, copyDefault, desc) {
    if (module && typeof module === "object" || typeof module === "function")
      for (var keys = __getOwnPropNames(module), i = 0, n = keys.length, key; i < n; i++) {
        key = keys[i];
        if (!__hasOwnProp.call(target, key) && (copyDefault || key !== "default"))
          __defProp(target, key, { get: function(k) {
            return module[k];
          }.bind(null, key), enumerable: !(desc = __getOwnPropDesc(module, key)) || desc.enumerable });
      }
    return target;
  };
  var __toCommonJS = /* @__PURE__ */ function(cache) {
    return function(module, temp) {
      return cache && cache.get(module) || (temp = __reExport(__markAsModule({}), module, 1), cache && cache.set(module, temp), temp);
    };
  }(typeof WeakMap !== "undefined" ? /* @__PURE__ */ new WeakMap() : 0);

  // 83_numbers.ts
  var numbers_exports = {};
  __export(numbers_exports, {
    parse_numbers: function() {
      return numbers_default;
    }
  });

  // src/util.ts
  var u8_to_dataview = function(array) {
    return new DataView(array.buffer, array.byteOffset, array.byteLength);
  };
  var u8str = function(u8) {
    return new TextDecoder().decode(u8);
  };
  var u8concat = function(u8a) {
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
  };

  // src/proto.ts
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
        case 1:
          if (!len)
            len = 8;
        case 2:
          if (!len)
            len = parse_varint49(buf, ptr);
          res = buf.slice(ptr[0], ptr[0] + len);
          ptr[0] += len;
          break;
        case 3:
        case 4:
        default:
          throw new Error("PB Type ".concat(type, " for Field ").concat(num, " at offset ").concat(off));
      }
      var v = { offset: off, data: res };
      if (out[num] == null)
        out[num] = [v];
      else
        out[num].push(v);
    }
    return out;
  }
  function mappa(data, cb) {
    if (!data)
      return [];
    return data.map(function(d) {
      var _a;
      try {
        return cb(d.data);
      } catch (e) {
        var m = (_a = e.message) == null ? void 0 : _a.match(/at offset (\d+)/);
        if (m)
          e.message = e.message.replace(/at offset (\d+)/, "at offset " + (+m[1] + d.offset));
        throw e;
      }
    });
  }

  // src/frame.ts
  function deframe(buf) {
    var out = [];
    var l = 0;
    while (l < buf.length) {
      var t = buf[l++];
      var len = buf[l] | buf[l + 1] << 8 | buf[l + 2] << 16;
      l += 3;
      out.push(process_chunk(t, buf.slice(l, l + len)));
      l += len;
    }
    if (l !== buf.length)
      throw new Error("data is not a valid framed stream!");
    return u8concat(out);
  }
  function process_chunk(type, buf) {
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

  // src/iwa.ts
  function parse_iwa(buf) {
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
      out.push(res);
    }
    return out;
  }

  // src/numbers.ts
  var encode_col = function(C) {
    var s = "";
    for (++C; C; C = Math.floor((C - 1) / 26))
      s = String.fromCharCode((C - 1) % 26 + 65) + s;
    return s;
  };
  var encode_cell = function(c) {
    return "".concat(encode_col(c.c)).concat(c.r + 1);
  };
  var encode_range = function(r) {
    return encode_cell(r.s) + ":" + encode_cell(r.e);
  };
  var book_new = function() {
    return { Sheets: {}, SheetNames: [] };
  };
  var book_append_sheet = function(wb, ws, name) {
    if (!name) {
      for (var i = 1; i < 9999; ++i)
        if (wb.SheetNames.indexOf(name = "Sheet ".concat(i)) == -1)
          break;
    }
    wb.SheetNames.push(name);
    wb.Sheets[name] = ws;
  };
  function parse_numbers(cfb) {
    var out = [];
    cfb.FileIndex.forEach(function(s) {
      if (!s.name.match(/\.iwa$/))
        return;
      var o;
      try {
        o = deframe(s.content);
      } catch (e) {
        return console.log("?? " + s.content.length + " " + (e.message || e));
      }
      var packets;
      try {
        packets = parse_iwa(o);
      } catch (e) {
        return console.log("## " + (e.message || e));
      }
      packets.forEach(function(packet) {
        out[+packet.id] = packet.messages;
      });
    });
    if (!out.length)
      throw new Error("File has no messages");
    var docroot;
    out.forEach(function(iwams) {
      iwams.forEach(function(iwam) {
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
    return parse_docroot(out, docroot);
  }
  var numbers_default = parse_numbers;
  function parse_Reference(buf) {
    var pb = parse_shallow(buf);
    return parse_varint49(pb[1][0].data);
  }
  function parse_TST_TableDataList(M, root) {
    var pb = parse_shallow(root.data);
    var entries = pb[3];
    var data = [];
    entries == null ? void 0 : entries.forEach(function(entry) {
      var le = parse_shallow(entry.data);
      var key = varint_to_i32(le[1][0].data) >>> 0;
      data[key] = u8str(le[3][0].data);
    });
    return data;
  }
  function parse_TST_TileRowInfo(u8) {
    var pb = parse_shallow(u8);
    var R = varint_to_i32(pb[1][0].data) >>> 0;
    var storage = pb[3][0].data;
    var offsets = u8_to_dataview(pb[4][0].data);
    var cells = [];
    for (var C = 0; C < offsets.byteLength / 2; ++C) {
      var off = offsets.getUint16(C * 2, true);
      if (off > storage.length)
        continue;
      cells[C] = storage.subarray(off, offsets.getUint16(C * 2 + 2, true));
    }
    return { R: R, cells: cells };
  }
  function parse_TST_Tile(M, root) {
    var pb = parse_shallow(root.data);
    var ri = mappa(pb[5], parse_TST_TileRowInfo);
    return ri.reduce(function(acc, x) {
      if (!acc[x.R])
        acc[x.R] = [];
      x.cells.forEach(function(cell, C) {
        if (acc[x.R][C])
          throw new Error("Duplicate cell r=".concat(x.R, " c=").concat(C));
        acc[x.R][C] = cell;
      });
      return acc;
    }, []);
  }
  function parse_TST_TableModelArchive(M, root, ws) {
    var pb = parse_shallow(root.data);
    var range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    range.e.r = (varint_to_i32(pb[6][0].data) >>> 0) - 1;
    if (range.e.r < 0)
      throw new Error("Invalid row varint ".concat(pb[6][0].data));
    range.e.c = (varint_to_i32(pb[7][0].data) >>> 0) - 1;
    if (range.e.c < 0)
      throw new Error("Invalid col varint ".concat(pb[7][0].data));
    ws["!ref"] = encode_range(range);
    {
      var store = parse_shallow(pb[4][0].data);
      var sst = parse_TST_TableDataList(M, M[parse_Reference(store[4][0].data)][0]);
      {
        var tile = parse_shallow(store[3][0].data);
        var tiles = [];
        tile[1].forEach(function(t) {
          var tl = parse_shallow(t.data);
          var ref = M[parse_Reference(tl[2][0].data)][0];
          var mtype = varint_to_i32(ref.meta[1][0].data);
          if (mtype != 6002)
            throw new Error("6001 unexpected reference to ".concat(mtype));
          tiles.push({ id: varint_to_i32(tl[1][0].data), ref: parse_TST_Tile(M, ref) });
        });
        tiles.forEach(function(tile2) {
          tile2.ref.forEach(function(row, R) {
            row.forEach(function(buf, C) {
              var dv = u8_to_dataview(buf);
              var ctype = buf[2];
              var addr = encode_cell({ r: R, c: C });
              switch (ctype) {
                case 0:
                  {
                    switch (buf[1]) {
                      case 3:
                        ws[addr] = { t: "s", v: sst[dv.getUint32(buf.length - 4, true)] };
                        break;
                      case 2:
                        ws[addr] = { t: "n", v: dv.getFloat64(16, true) };
                        break;
                      case 0:
                        break;
                      case 5:
                        break;
                      case 7:
                        break;
                      case 6:
                        ws[addr] = { t: "b", v: dv.getFloat64(buf.length - 8, true) > 0 };
                        break;
                      default:
                        throw new Error("Unsupported cell type ".concat(buf.slice(0, 4)));
                    }
                  }
                  break;
                case 3:
                  {
                    ws[addr] = { t: "s", v: sst[dv.getUint32(16, true)] };
                  }
                  break;
                case 2:
                  {
                    ws[addr] = { t: "n", v: dv.getFloat64(buf.length - 12, true) };
                  }
                  break;
                case 6:
                  {
                    ws[addr] = { t: "b", v: dv.getFloat64(16, true) > 0 };
                  }
                  break;
                default:
                  throw new Error("Unsupported cell type ".concat(ctype));
              }
            });
          });
        });
      }
    }
  }
  function parse_TST_TableInfoArchive(M, root) {
    var pb = parse_shallow(root.data);
    var out = { "!ref": "A1" };
    var tableref = M[parse_Reference(pb[2][0].data)];
    var mtype = varint_to_i32(tableref[0].meta[1][0].data);
    if (mtype != 6001)
      throw new Error("6000 unexpected reference to ".concat(mtype));
    parse_TST_TableModelArchive(M, tableref[0], out);
    return out;
  }
  function parse_sheetroot(M, root) {
    var _a;
    var pb = parse_shallow(root.data);
    var out = [{ "!ref": "A1" }, ((_a = pb[1]) == null ? void 0 : _a[0]) ? u8str(pb[1][0].data) : ""];
    var shapeoffs = mappa(pb[2], parse_Reference);
    var seen = false;
    shapeoffs.forEach(function(off) {
      M[off].forEach(function(m) {
        var mtype = varint_to_i32(m.meta[1][0].data);
        if (mtype == 6e3) {
          if (seen)
            return;
          out[0] = parse_TST_TableInfoArchive(M, m);
          seen = true;
        }
      });
    });
    return out;
  }
  function parse_docroot(M, root) {
    var out = book_new();
    var pb = parse_shallow(root.data);
    var sheetoffs = mappa(pb[1], parse_Reference);
    sheetoffs.forEach(function(off) {
      M[off].forEach(function(m) {
        var mtype = varint_to_i32(m.meta[1][0].data);
        if (mtype == 2) {
          var root2 = parse_sheetroot(M, m);
          book_append_sheet(out, root2[0], root2[1]);
        }
      });
    });
    if (out.SheetNames.length == 0)
      throw new Error("Empty NUMBERS file");
    return out;
  }
  return __toCommonJS(numbers_exports);
})();
/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
