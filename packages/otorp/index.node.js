var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __markAsModule = (target) => __defProp(target, "__esModule", { value: true });
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __reExport = (target, module2, copyDefault, desc) => {
  if (module2 && typeof module2 === "object" || typeof module2 === "function") {
    for (let key of __getOwnPropNames(module2))
      if (!__hasOwnProp.call(target, key) && (copyDefault || key !== "default"))
        __defProp(target, key, { get: () => module2[key], enumerable: !(desc = __getOwnPropDesc(module2, key)) || desc.enumerable });
  }
  return target;
};
var __toCommonJS = /* @__PURE__ */ ((cache) => {
  return (module2, temp) => {
    return cache && cache.get(module2) || (temp = __reExport(__markAsModule({}), module2, 1), cache && cache.set(module2, temp), temp);
  };
})(typeof WeakMap !== "undefined" ? /* @__PURE__ */ new WeakMap() : 0);

// index.node.ts
var index_node_exports = {};
__export(index_node_exports, {
  otorp: () => otorp_default
});

// src/util.ts
var u8_to_dataview = (array) => new DataView(array.buffer, array.byteOffset, array.byteLength);
var u8str = (u8) => new TextDecoder().decode(u8);
var indent = (str, depth) => str.split(/\n/g).map((x) => x && "  ".repeat(depth) + x).join("\n");
function u8indexOf(u8, data, byteOffset) {
  if (typeof data == "number")
    return u8.indexOf(data, byteOffset);
  var l = byteOffset;
  if (typeof data == "string") {
    outs:
      while ((l = u8.indexOf(data.charCodeAt(0), l)) > -1) {
        ++l;
        for (var j = 1; j < data.length; ++j)
          if (u8[l + j - 1] != data.charCodeAt(j))
            continue outs;
        return l - 1;
      }
  } else {
    outb:
      while ((l = u8.indexOf(data[0], l)) > -1) {
        ++l;
        for (var j = 1; j < data.length; ++j)
          if (u8[l + j - 1] != data[j])
            continue outb;
        return l - 1;
      }
  }
  return -1;
}

// src/macho.ts
var parse_fat = (buf) => {
  var dv = u8_to_dataview(buf);
  if (dv.getUint32(0, false) !== 3405691582)
    throw new Error("Unsupported file");
  var nfat_arch = dv.getUint32(4, false);
  var out = [];
  for (var i = 0; i < nfat_arch; ++i) {
    var start = i * 20 + 8;
    var cputype = dv.getUint32(start, false);
    var cpusubtype = dv.getUint32(start + 4, false);
    var offset = dv.getUint32(start + 8, false);
    var size = dv.getUint32(start + 12, false);
    var align = dv.getUint32(start + 16, false);
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
var parse_macho = (buf) => {
  var dv = u8_to_dataview(buf);
  var magic = dv.getUint32(0, false);
  switch (magic) {
    case 3405691582:
      return parse_fat(buf);
    case 3489328638:
      return [{
        type: dv.getUint32(4, false),
        subtype: dv.getUint32(8, false),
        offset: 0,
        size: buf.length,
        data: buf
      }];
  }
  throw new Error("Unsupported file");
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
        throw new Error(`PB Type ${type} for Field ${num} at offset ${off}`);
    }
    var v = { offset: off, data: res, type };
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
  return data.map((d) => {
    try {
      return cb(d.data);
    } catch (e) {
      var m = e.message?.match(/at offset (\d+)/);
      if (m)
        e.message = e.message.replace(/at offset (\d+)/, "at offset " + (+m[1] + d.offset));
      throw e;
    }
  });
}

// src/descriptor.ts
var TYPES = [
  "error",
  "double",
  "float",
  "int64",
  "uint64",
  "int32",
  "fixed64",
  "fixed32",
  "bool",
  "string",
  "group",
  "message",
  "bytes",
  "uint32",
  "enum",
  "sfixed32",
  "sfixed64",
  "sint32",
  "sint64"
];
function parse_FileOptions(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.javaPackage = u8str(data[1][0].data);
  if (data[8]?.[0])
    out.javaOuterClassname = u8str(data[8][0].data);
  if (data[11]?.[0])
    out.goPackage = u8str(data[11][0].data);
  return out;
}
function parse_EnumValue(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.name = u8str(data[1][0].data);
  if (data[2]?.[0])
    out.number = varint_to_i32(data[2][0].data);
  return out;
}
function parse_Enum(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.name = u8str(data[1][0].data);
  out.value = mappa(data[2], parse_EnumValue);
  return out;
}
var write_Enum = (en) => {
  var out = [`enum ${en.name} {`];
  en.value?.forEach(({ name, number }) => out.push(`  ${name} = ${number};`));
  return out.concat(`}`).join("\n");
};
function parse_FieldOptions(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[2]?.[0])
    out.packed = !!data[2][0].data;
  if (data[3]?.[0])
    out.deprecated = !!data[3][0].data;
  return out;
}
function parse_Field(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.name = u8str(data[1][0].data);
  if (data[2]?.[0])
    out.extendee = u8str(data[2][0].data);
  if (data[3]?.[0])
    out.number = varint_to_i32(data[3][0].data);
  if (data[4]?.[0])
    out.label = varint_to_i32(data[4][0].data);
  if (data[5]?.[0])
    out.type = varint_to_i32(data[5][0].data);
  if (data[6]?.[0])
    out.typeName = u8str(data[6][0].data);
  if (data[7]?.[0])
    out.defaultValue = u8str(data[7][0].data);
  if (data[8]?.[0])
    out.options = parse_FieldOptions(data[8][0].data);
  return out;
}
function write_Field(field) {
  var out = [];
  var label = ["", "optional ", "required ", "repeated "][field.label] || "";
  var type = field.typeName || TYPES[field.type] || "s5s";
  var opts = [];
  if (field.defaultValue)
    opts.push(`default = ${field.defaultValue}`);
  if (field.options?.packed)
    opts.push(`packed = true`);
  if (field.options?.deprecated)
    opts.push(`deprecated = true`);
  var os = opts.length ? ` [${opts.join(", ")}]` : "";
  out.push(`${label}${type} ${field.name} = ${field.number}${os};`);
  return out.length ? indent(out.join("\n"), 1) : "";
}
function write_extensions(ext, xtra = false, coalesce = true) {
  var res = [];
  var xt = [];
  ext.forEach((ext2) => {
    if (!ext2.extendee)
      return;
    var row = coalesce ? xt.find((x) => x[0] == ext2.extendee) : xt[xt.length - 1]?.[0] == ext2.extendee ? xt[xt.length - 1] : null;
    if (row)
      row[1].push(ext2);
    else
      xt.push([ext2.extendee, [ext2]]);
  });
  xt.forEach((extrow) => {
    var out = [`extend ${extrow[0]} {`];
    extrow[1].forEach((ext2) => out.push(write_Field(ext2)));
    res.push(out.concat(`}`).join("\n") + (xtra ? "\n" : ""));
  });
  return res.join("\n");
}
function parse_mtype(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.name = u8str(data[1][0].data);
  if (data[2]?.length >= 1)
    out.field = mappa(data[2], parse_Field);
  if (data[3]?.length >= 1)
    out.nestedType = mappa(data[3], parse_mtype);
  if (data[4]?.length >= 1)
    out.enumType = mappa(data[4], parse_Enum);
  if (data[6]?.length >= 1)
    out.extension = mappa(data[6], parse_Field);
  if (data[5]?.length >= 1)
    out.extensionRange = data[5].map((d) => {
      var data2 = parse_shallow(d.data);
      var out2 = {};
      if (data2[1]?.[0])
        out2.start = varint_to_i32(data2[1][0].data);
      if (data2[2]?.[0])
        out2.end = varint_to_i32(data2[2][0].data);
      return out2;
    });
  return out;
}
var write_mtype = (message) => {
  var out = [`message ${message.name} {`];
  message.nestedType?.forEach((m) => out.push(indent(write_mtype(m), 1)));
  message.enumType?.forEach((en) => out.push(indent(write_Enum(en), 1)));
  message.field?.forEach((field) => out.push(write_Field(field)));
  if (message.extensionRange)
    message.extensionRange.forEach((er) => out.push(`  extensions ${er.start} to ${er.end - 1};`));
  if (message.extension?.length)
    out.push(indent(write_extensions(message.extension), 1));
  return out.concat(`}`).join("\n");
};
function parse_FileDescriptor(buf) {
  var data = parse_shallow(buf);
  var out = {};
  if (data[1]?.[0])
    out.name = u8str(data[1][0].data);
  if (data[2]?.[0])
    out.package = u8str(data[2][0].data);
  if (data[3]?.[0])
    out.dependency = data[3].map((x) => u8str(x.data));
  if (data[4]?.length >= 1)
    out.messageType = mappa(data[4], parse_mtype);
  if (data[5]?.length >= 1)
    out.enumType = mappa(data[5], parse_Enum);
  if (data[7]?.length >= 1)
    out.extension = mappa(data[7], parse_Field);
  if (data[8]?.[0])
    out.options = parse_FileOptions(data[8][0].data);
  return out;
}
var write_FileDescriptor = (pb) => {
  var out = [
    'syntax = "proto2";',
    ""
  ];
  if (pb.dependency)
    pb.dependency.forEach((n) => {
      if (n)
        out.push(`import "${n}";`);
    });
  if (pb.package)
    out.push(`package ${pb.package};
`);
  if (pb.options) {
    var o = out.length;
    if (pb.options.javaPackage)
      out.push(`option java_package = "${pb.options.javaPackage}";`);
    if (pb.options.javaOuterClassname?.replace(/\W/g, ""))
      out.push(`option java_outer_classname = "${pb.options.javaOuterClassname}";`);
    if (pb.options.javaMultipleFiles)
      out.push(`option java_multiple_files = true;`);
    if (pb.options.goPackage)
      out.push(`option go_package = "${pb.options.goPackage}";`);
    if (out.length > o)
      out.push("");
  }
  pb.enumType?.forEach((en) => {
    if (en.name)
      out.push(write_Enum(en) + "\n");
  });
  pb.messageType?.forEach((m) => {
    if (m.name) {
      var o2 = write_mtype(m);
      if (o2)
        out.push(o2 + "\n");
    }
  });
  if (pb.extension?.length) {
    var e = write_extensions(pb.extension, true, false);
    if (e)
      out.push(e);
  }
  return out.join("\n") + "\n";
};

// src/otorp.ts
function otorp(buf, builtins = false) {
  var res = proto_offsets(buf);
  var registry = {};
  var names = /* @__PURE__ */ new Set();
  var out = [];
  res.forEach((r, i) => {
    if (!builtins && r[1].startsWith("google/protobuf/"))
      return;
    var b = buf.slice(r[0], i < res.length - 1 ? res[i + 1][0] : buf.length);
    var pb = parse_FileDescriptorProto(b);
    names.add(r[1]);
    registry[r[1]] = pb;
  });
  names.forEach((name) => {
    names.delete(name);
    var pb = registry[name];
    var doit = (pb.dependency || []).every((d) => !names.has(d));
    if (!doit) {
      names.add(name);
      return;
    }
    var dups = res.filter((r) => r[1] == name);
    if (dups.length == 1)
      return out.push({ name, proto: write_FileDescriptor(pb) });
    var pbs = dups.map((r) => {
      var i = res.indexOf(r);
      var b = buf.slice(r[0], i < res.length - 1 ? res[i + 1][0] : buf.length);
      var pb2 = parse_FileDescriptorProto(b);
      return write_FileDescriptor(pb2);
    });
    for (var l = 1; l < pbs.length; ++l)
      if (pbs[l] != pbs[0])
        throw new Error(`Conflicting definitions for ${name} at offsets 0x${dups[0][0].toString(16)} and 0x${dups[l][0].toString(16)}`);
    return out.push({ name, proto: pbs[0] });
  });
  return out;
}
var otorp_default = otorp;
var is_referenced = (buf, pos) => {
  var dv = u8_to_dataview(buf);
  for (var leaddr = 0; leaddr > -1 && leaddr < pos; leaddr = u8indexOf(buf, 141, leaddr + 1))
    if (dv.getUint32(leaddr + 2, true) == pos - leaddr - 6)
      return true;
  try {
    var headers = parse_macho(buf);
    for (var i = 0; i < headers.length; ++i) {
      if (pos < headers[i].offset || pos > headers[i].offset + headers[i].size)
        continue;
      var b = headers[i].data;
      var p = pos - headers[i].offset;
      var ref = new Uint8Array([0, 0, 0, 0, 0, 0, 0, 0]);
      var dv = u8_to_dataview(ref);
      dv.setInt32(0, p, true);
      if (u8indexOf(b, ref, 0) > 0)
        return true;
      ref[4] = 1;
      if (u8indexOf(b, ref, 0) > 0)
        return true;
    }
  } catch (e) {
  }
  return false;
};
var proto_offsets = (buf) => {
  var meta = parse_macho(buf);
  var out = [];
  var off = 0;
  search:
    while ((off = u8indexOf(buf, ".proto", off + 1)) > -1) {
      var pos = off;
      off += 6;
      while (off - pos < 256 && buf[pos] != off - pos - 1) {
        if (buf[pos] > 127 || buf[pos] < 32)
          continue search;
        --pos;
      }
      if (off - pos > 250)
        continue;
      var name = u8str(buf.slice(pos + 1, off));
      if (buf[--pos] != 10)
        continue;
      if (!is_referenced(buf, pos)) {
        console.error(`Reference to ${name} not found`);
        continue;
      }
      var bin = meta.find((m) => m.offset <= pos && m.offset + m.size >= pos);
      out.push([pos, name, bin?.type || -1, bin?.subtype || -1]);
    }
  return out;
};
var parse_FileDescriptorProto = (buf) => {
  var l = buf.length;
  while (l > 0)
    try {
      var b = buf.slice(0, l);
      var o = parse_FileDescriptor(b);
      return o;
    } catch (e) {
      var m = e.message.match(/at offset (\d+)/);
      if (m && parseInt(m[1], 10) < buf.length)
        l = parseInt(m[1], 10) - 1;
      else
        --l;
    }
  throw new RangeError("no protobuf message in range");
};
module.exports = __toCommonJS(index_node_exports);
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  otorp
});
/*! otorp (C) 2013-present SheetJS -- http://sheetjs.com */
/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
