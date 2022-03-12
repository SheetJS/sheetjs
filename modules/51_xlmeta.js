RELS.XLMETA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
function parse_xlmeta_xml(data, name, opts) {
  var out = { Types: [] };
  if (!data)
    return out;
  var pass = false;
  data.replace(tagregex, function(x, idx) {
    var y = parsexmltag(x);
    switch (strip_ns(y[0])) {
      case "<?xml":
        break;
      case "<metadata":
      case "</metadata>":
        break;
      case "<metadataTypes":
      case "</metadataTypes>":
        break;
      case "<metadataType":
        out.Types.push({ name: y.name });
        break;
      case "<futureMetadata":
        break;
      case "</futureMetadata>":
        break;
      case "<bk>":
        break;
      case "</bk>":
        break;
      case "<rc":
        break;
      case "</rc>":
        break;
      case "<cellMetadata":
      case "</cellMetadata>":
        break;
      case "<valueMetadata":
        break;
      case "</valueMetadata>":
        break;
      case "<extLst":
      case "<extLst>":
      case "</extLst>":
      case "<extLst/>":
        break;
      case "<ext":
        pass = true;
        break;
      case "</ext>":
        pass = false;
        break;
      default:
        if (!pass && opts.WTF)
          throw new Error("unrecognized " + y[0] + " in metadata");
    }
    return x;
  });
  return out;
}
function write_xlmeta_xml() {
  var o = [XML_HEADER];
  o.push('<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">\n  <metadataTypes count="1">\n    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>\n  </metadataTypes>\n  <futureMetadata name="XLDAPR" count="1">\n    <bk>\n      <extLst>\n        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">\n          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>\n        </ext>\n      </extLst>\n    </bk>\n  </futureMetadata>\n  <cellMetadata count="1">\n    <bk>\n      <rc t="1" v="0"/>\n    </bk>\n  </cellMetadata>\n</metadata>');
  return o.join("");
}
function parse_BrtMdtinfo(data, length) {
  return {
    flags: data.read_shift(4),
    version: data.read_shift(4),
    name: parse_XLWideString(data, length - 8)
  };
}
function write_BrtMdtinfo(data) {
  var o = new_buf(12 + 2 * data.name.length);
  o.write_shift(4, data.flags);
  o.write_shift(4, data.version);
  write_XLWideString(data.name, o);
  return o.slice(0, o.l);
}
function write_BrtMdb(mdb) {
  var o = new_buf(4 + 8 * mdb.length);
  o.write_shift(4, mdb.length);
  for (var i = 0; i < mdb.length; ++i) {
    o.write_shift(4, mdb[i][0]);
    o.write_shift(4, mdb[i][1]);
  }
  return o;
}
function write_BrtBeginEsfmd(cnt, name) {
  var o = new_buf(8 + 2 * name.length);
  o.write_shift(4, cnt);
  write_XLWideString(name, o);
  return o.slice(0, o.l);
}
function write_BrtBeginEsmdb(cnt, cm) {
  var o = new_buf(8);
  o.write_shift(4, cnt);
  o.write_shift(4, cm ? 1 : 0);
  return o;
}
function parse_xlmeta_bin(data, name, _opts) {
  var out = { Types: [] };
  var opts = _opts || {};
  var state = [];
  var pass = false;
  recordhopper(data, function(val, R, RT) {
    switch (RT) {
      case 335:
        out.Types.push({ name: val.name });
        break;
      case 51:
        break;
      case 35:
        state.push(RT);
        pass = true;
        break;
      case 36:
        state.pop();
        pass = false;
        break;
      default:
        if (R.T) {
        } else if (!pass || opts.WTF && state[state.length - 1] != 35)
          throw new Error("Unexpected record 0x" + RT.toString(16));
    }
  });
  return out;
}
function write_xlmeta_bin() {
  var ba = buf_array();
  write_record(ba, 332);
  write_record(ba, 334, write_UInt32LE(1));
  write_record(ba, 335, write_BrtMdtinfo({
    name: "XLDAPR",
    version: 12e4,
    flags: 3496657072
  }));
  write_record(ba, 336);
  write_record(ba, 339, write_BrtBeginEsfmd(1, "XLDAPR"));
  write_record(ba, 52);
  write_record(ba, 35, write_UInt32LE(514));
  write_record(ba, 4096, write_UInt32LE(0));
  write_record(ba, 4097, writeuint16(1));
  write_record(ba, 36);
  write_record(ba, 53);
  write_record(ba, 340);
  write_record(ba, 337, write_BrtBeginEsmdb(1, true));
  write_record(ba, 51, write_BrtMdb([[1, 0]]));
  write_record(ba, 338);
  write_record(ba, 333);
  return ba.end();
}
