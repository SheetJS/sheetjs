/*! sheetjs (C) 2013-present SheetJS -- http://sheetjs.com */
import { decompress_iwa_file, compress_iwa_file, u8concat } from "./src/numbers";
import { read, writeFile, utils, find, CFB$Entry } from 'cfb';

var f = process.argv[2]; var o = process.argv[3];
var cfb = read(f, {type: "file"});

var FI = cfb.FileIndex;
var bufs = [], mode = 0;
FI.map((fi, idx): [CFB$Entry, string] => ([fi, cfb.FullPaths[idx]])).forEach(row => {
  var fi = row[0], fp = row[1];
  /* blank all plist files */
  if(fi.name.match(/\.plist/)) {
    console.error(`Blanking plist ${fi.name}`);
    fi.content = new Uint8Array([
      0x62, 0x70, 0x6c, 0x69, 0x73, 0x74, 0x30, 0x30, 0xd0, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x09, 0x0a
    ]);
    return;
  }

  if(fi.type != 2) return;

  /* Remove other metadata */
  if(!fi.name.match(/\.iwa/)) {
    if(fi.name.match(/Sh33tJ5/)) return;
    console.error(`Removing file ${fi.name}`);
    utils.cfb_del(cfb, fp);
    return;
  }

  /* Reframe .iwa files */
  console.error(`Reframing iwa ${fi.name} (${fi.size})`);
  var old_content = fi.content;
  var raw1 = decompress_iwa_file(old_content as Uint8Array);
  var new_content = compress_iwa_file(raw1);
  var raw2 = decompress_iwa_file(new_content);
  for(var i = 0; i < raw1.length; ++i) if(raw1[i] != raw2[i]) throw new Error(`${fi.name} did not properly roundtrip`);
  bufs.push(raw1);
  switch(mode) {
    case 1: utils.cfb_del(cfb, fp); break;
    /* falls through */
    case 0: fi.content = new_content; break;
  }
});

if(mode == 1) {
  var res = compress_iwa_file(u8concat(bufs));
  console.error(`Adding iwa Document.iwa (${res.length})`);
  utils.cfb_add(cfb, "/Index/Document.iwa", res);
}
writeFile(cfb, o, {fileType: "zip"});
